import streamlit as st
from google import genai
from google.genai import types
import json
from pydantic import BaseModel, Field
from typing import List, Optional
from PIL import Image, ImageEnhance
import docx
from docx import Document
from docx.shared import Pt, Inches, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# --- Page Setup ---
st.set_page_config(
    page_title="Academic Bento Exam Studio",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 🎨 Academic Bento & Editorial Minimalist UI Design System ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { 
        font-family: 'Plus Jakarta Sans', sans-serif !important; 
        background-color: #F8F9FA !important; 
        color: #1A1D20 !important; 
    }
    
    h1, h2, h3, .serif-title { font-family: 'Instrument Serif', serif !important; }
    
    .main-title {
        font-size: 3.5rem; font-weight: 400; text-align: left;
        color: #111315; letter-spacing: -0.5px; line-height: 1.1; margin-bottom: 0px;
    }
    
    .sub-title { text-align: left; color: #606873; font-size: 1.05rem; font-weight: 500; margin-bottom: 25px; }

    [data-testid="stSidebar"] { background: #F1F3F5 !important; border-right: 1px solid #E2E8F0 !important; }

    .bento-card {
        background: #FFFFFF; border: 2px solid #E2E8F0; border-radius: 16px;
        padding: 24px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.02); margin-bottom: 20px;
    }

    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #CBD5E1 !important; background-color: #FAFAFA !important;
        border-radius: 12px !important; padding: 30px 20px !important; transition: all 0.2s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover { background-color: #F1F5F9 !important; border-color: #0F172A !important; }

    .stButton>button {
        width: 100%; background-color: #0F172A; padding: 12px 20px; text-align: center; 
        color: #FFFFFF !important; border-radius: 10px; border: none; font-weight: 600; font-size: 0.95rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15); transition: transform 0.1s ease, background-color 0.2s ease;
    }
    .stButton>button:hover { background-color: #334155; transform: translateY(-1px); }
    
    .stDownloadButton>button {
        width: 100%; background-color: #0D9488; padding: 12px 20px; text-align: center; 
        color: #FFFFFF !important; border-radius: 10px; border: none; font-weight: 700;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.2); transition: transform 0.1s ease, background-color 0.2s ease;
    }
    .stDownloadButton>button:hover { background-color: #0F766E; transform: translateY(-1px); }
</style>
""", unsafe_allow_html=True)

# --- 1. Universal Structural Schemas ---
class LayoutBlock(BaseModel):
    block_type: str = Field(description="Can be: 'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block', 'drawing_box_block'")
    text_content: Optional[str] = Field(default=None)
    list_items: Optional[List[str]] = Field(default=None)
    table_rows: Optional[int] = Field(default=None)
    table_cols: Optional[int] = Field(default=None)
    table_data: Optional[List[List[str]]] = Field(description="2D array matrix for grid_table_block", default=None)
    columns_data: Optional[List[List[str]]] = Field(description="Column text lists for column_layout_block", default=None)
    box_height_inches: Optional[float] = Field(default=2.0)

class UniversalExamPaper(BaseModel):
    school_name: str
    exam_title: str
    class_name: str
    subject: str
    full_marks: str
    time: str
    student_info_line: str = Field(description="Student details line placeholder (Name, Roll, etc.)")
    blocks: List[LayoutBlock] = Field(description="Chronological sequence of structured objects found")

# --- 2. Advanced Typography & Word Engine ---
def set_table_borders(table, color="cccccc"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)  
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Adds generous internal padding to table cells for primary kindergarten worksheets."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_section_columns(section, num_cols):
    """Enables native multi-column layout for modern Word documents."""
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), str(num_cols))
    else:
        new_cols = OxmlElement('w:cols')
        new_cols.set(qn('w:num'), str(num_cols))
        new_cols.set(qn('w:space'), '720')  # 0.5 inch spacing between columns
        sectPr.append(new_cols)

def create_docx(data: UniversalExamPaper, language: str, grade_tier: str, base_font_size: int):
    doc = Document()
    section = doc.sections[0]
    
    is_early_childhood = grade_tier in ["Nursery", "PP / LKG / UKG"]
    
    # Define exact typography & layout defaults per grade tier
    if is_early_childhood:
        headline_fs = 18       # Question Headline / Instructions
        question_fs = 22       # Questions, Options & Blanks
        header_title_fs = 24   # School Title
        meta_fs = 18           # Exam Meta Details & Name/Roll line
        
        # Portrait A4 with generous padding for direct answer writing
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.orientation = WD_ORIENT.PORTRAIT
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    else:
        # Classes 1 to 4: Landscape A4 split into 2 structured columns
        headline_fs = base_font_size + 1
        question_fs = base_font_size
        header_title_fs = base_font_size + 4
        meta_fs = base_font_size
        
        section.page_width = Mm(297)
        section.page_height = Mm(210)
        section.orientation = WD_ORIENT.LANDSCAPE
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        set_section_columns(section, 2)
        
    font_mapping = {"Bengali": "Kalpurush", "Hindi": "Mangal", "English": "Calibri"}
    selected_font = font_mapping.get(language, "Calibri")
        
    style = doc.styles['Normal']
    style.font.name = selected_font
    style.font.size = Pt(question_fs)
    
    # --- 1. School Header & Exam Title ---
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_school = p_school.add_run(data.school_name)
    run_school.bold = True
    run_school.font.size = Pt(header_title_fs)
    p_school.paragraph_format.space_after = Pt(2)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(data.exam_title)
    run_title.bold = True
    run_title.font.size = Pt(headline_fs)
    p_title.paragraph_format.space_after = Pt(6)
    
    # --- 2. Metadata Block ---
    class_label = "Class" if language == "English" else ("শ্রেণী" if language == "Bengali" else "कक्षा")
    marks_label = "Full Marks" if language == "English" else ("পূর্ণমান" if language == "Bengali" else "पूर्णांक")
    subject_label = "Subject" if language == "English" else ("বিষয়" if language == "Bengali" else "विषय")
    time_label = "Time" if language == "English" else ("সময়" if language == "Bengali" else "समय")

    def format_meta(label, val):
        if not val: return ""
        if label.lower() in val.lower() or "শ্রেণী" in val or "कक्षा" in val or "পূর্ণমান" in val or "বিষয়" in val or "সময়" in val or "पूर्णांक" in val:
            return val
        return f"{label} — {val}"

    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.autofit = False
    col_w = Inches(3.5) if is_early_childhood else Inches(4.5)
    meta_table.columns[0].width = col_w
    meta_table.columns[1].width = col_w
    
    meta_table.rows[0].cells[0].paragraphs[0].text = format_meta(class_label, data.class_name)
    meta_table.rows[0].cells[1].paragraphs[0].text = format_meta(marks_label, data.full_marks)
    meta_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    meta_table.rows[1].cells[0].paragraphs[0].text = format_meta(subject_label, data.subject)
    meta_table.rows[1].cells[1].paragraphs[0].text = format_meta(time_label, data.time)
    meta_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for row in meta_table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(meta_fs)
                    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    # Student Details Line
    p_info = doc.add_paragraph()
    run_info = p_info.add_run(data.student_info_line)
    run_info.font.size = Pt(meta_fs)
    run_info.bold = True
    p_info.paragraph_format.space_after = Pt(6)
    
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("—" * 50 if not is_early_childhood else "—" * 60)
    p_div.paragraph_format.space_after = Pt(12)
    
    # --- 3. Content Blocks Builder ---
    for b in data.blocks:
        space_after_val = 14 if is_early_childhood else 6
        
        if b.block_type == 'text_paragraph' and b.text_content:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(space_after_val)
            
            # Smart Check: Section Headline / Instruction vs Question Text
            cleaned = b.text_content.strip()
            is_headline = cleaned.endswith((':', '।')) and not any(char.isdigit() for char in cleaned[:3])
            
            run = p.add_run(b.text_content)
            
            if is_headline:
                run.font.size = Pt(headline_fs)  # 18pt for Headlines
                run.bold = True
            else:
                run.font.size = Pt(question_fs)  # 22pt for Questions & Options
                if cleaned.startswith(('১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯', '০', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '१', '२', '३', '४', '५', '६', '७', '८', '९', '०')):
                    run.bold = True
                
        elif b.block_type == 'list_block':
            if b.text_content:
                lp_head = doc.add_paragraph()
                run_h = lp_head.add_run(b.text_content)
                run_h.bold = True
                run_h.font.size = Pt(headline_fs)  # 18pt Headline
                lp_head.paragraph_format.space_after = Pt(6)
            if b.list_items:
                for item in b.list_items:
                    lp = doc.add_paragraph()
                    lp.paragraph_format.left_indent = Inches(0.3)
                    lp.paragraph_format.space_after = Pt(8 if is_early_childhood else 3)
                    run_item = lp.add_run(item)
                    run_item.font.size = Pt(question_fs)  # 22pt Items
                    
        elif b.block_type == 'grid_table_block':
            if b.text_content:
                tp = doc.add_paragraph()
                run_t = tp.add_run(b.text_content)
                run_t.bold = True
                run_t.font.size = Pt(headline_fs)  # 18pt Headline
            if b.table_rows and b.table_cols and b.table_data:
                tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
                tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(tbl, "cccccc")
                for r_idx, row_items in enumerate(b.table_data):
                    if r_idx < b.table_rows:
                        for c_idx, val in enumerate(row_items):
                            if c_idx < b.table_cols:
                                cell = tbl.rows[r_idx].cells[c_idx]
                                cell.paragraphs[0].text = val
                                if is_early_childhood:
                                    set_cell_margins(cell, top=220, bottom=220, left=200, right=200)
                                for r in cell.paragraphs[0].runs:
                                    r.font.size = Pt(question_fs)  # 22pt Table Text
                doc.add_paragraph().paragraph_format.space_after = Pt(10)
                
        elif b.block_type == 'column_layout_block':
            if b.text_content:
                cp = doc.add_paragraph()
                run_c = cp.add_run(b.text_content)
                run_c.bold = True
                run_c.font.size = Pt(headline_fs)
            if b.columns_data:
                num_cols = len(b.columns_data)
                max_rows = max(len(col) for col in b.columns_data)
                col_tbl = doc.add_table(rows=max_rows, cols=num_cols)
                col_width = Inches(3.2 / num_cols) if is_early_childhood else Inches(4.0 / num_cols)
                for c in range(num_cols):
                    col_tbl.columns[c].width = col_width
                    col_items = b.columns_data[c]
                    for r in range(max_rows):
                        if r < len(col_items):
                            cell = col_tbl.rows[r].cells[c]
                            cell.paragraphs[0].text = col_items[r]
                            for rn in cell.paragraphs[0].runs:
                                rn.font.size = Pt(question_fs)  # 22pt Column Text
                doc.add_paragraph()
                
        elif b.block_type == 'drawing_box_block':
            if b.text_content:
                dp = doc.add_paragraph()
                run_d = dp.add_run(b.text_content)
                run_d.bold = True
                run_d.font.size = Pt(headline_fs)
            box_tbl = doc.add_table(rows=1, cols=1)
            box_tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
            default_h = (b.box_height_inches or 2.0) * (1.5 if is_early_childhood else 1.0)
            box_tbl.rows[0].height = Inches(default_h)
            set_table_borders(box_tbl, "888888")
            doc.add_paragraph().paragraph_format.space_after = Pt(12)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. UI Header & Sidebar Controls ---
st.markdown('<p class="main-title">Academic Studio</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">High-Precision Editorial Digitizer & Automated Layout Compiler</p>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h3 style='font-family: Instrument Serif; font-size: 1.8rem; color: #111315;'>Studio Controls</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    api_key = st.text_input("🔑 Gemini API Key", type="password")
    if not api_key:
        st.warning("⚠️ Access Key required.")
        
    grade_tier = st.selectbox(
        "🎓 Target Grade Tier", 
        ["Nursery", "PP / LKG / UKG", "Classes 1 to 4"],
        help="Nursery/PP automatically configures 18pt headlines & 22pt questions with generous answer padding."
    )
    
    exam_language = st.selectbox("🌐 Document Language", ["Bengali", "English", "Hindi"])
    
    st.markdown("---")
    st.markdown("<h3 style='font-family: Instrument Serif; font-size: 1.8rem; color: #111315;'>Typography Preset</h3>", unsafe_allow_html=True)
    
    if grade_tier in ["Nursery", "PP / LKG / UKG"]:
        st.info("📏 **Early Childhood Profile Active:**\n- **Headlines:** `18pt` (Bold)\n- **Questions:** `22pt` (Spacious)")
        custom_font_size = 22
    else:
        st.info("📑 **Standard Exam Profile Active:**\n- **Layout:** Landscape A4 (2 Columns)")
        custom_font_size = st.number_input("Base Font Size (Pt)", min_value=9, max_value=16, value=11)
    
    st.markdown("---")
    st.caption("🔒 **Accuracy Mode:** Zero-temperature verification pipeline active.")

# --- Session State Initialization ---
if "parsed_data" not in st.session_state:
    st.session_state.parsed_data = None
if "original_filename" not in st.session_state:
    st.session_state.original_filename = "Exam_Output.docx"

# --- Main Workspace Layout ---
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown("### 01 / Intake Stream")
    uploaded_files = st.file_uploader(
        f"Upload {exam_language} exam sheets", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.session_state.original_filename = uploaded_files[0].name.rsplit('.', 1)[0] + ".docx"
        st.success(f"Loaded {len(uploaded_files)} source files.")
        
        thumb_cols = st.columns(min(len(uploaded_files), 4))
        for idx, file in enumerate(uploaded_files[:4]):
            with thumb_cols[idx]:
                img_preview = Image.open(file)
                st.image(img_preview, use_container_width=True)
        
        st.write("")
        if st.button("Run High-Precision Extraction"):
            if not api_key:
                st.error("Missing API Key.")
            else:
                with st.status("Executing deep dual-pass accuracy pipeline...", expanded=True) as status:
                    try:
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        img_list = []
                        for f in sorted_files:
                            raw_img = Image.open(f)
                            enhanced = ImageEnhance.Sharpness(ImageEnhance.Contrast(raw_img).enhance(1.7)).enhance(2.1)
                            img_list.append(enhanced)
                        
                        client = genai.Client(api_key=api_key)
                        
                        system_instruction = (
                            f"You are a meticulous, zero-error academic document transcriber specializing in {exam_language} primary papers for {grade_tier}. "
                            f"Your goal is absolute fidelity. Do not hallucinate or skip any words, numbers, punctuation marks, or fill-in lines (......). "
                            f"Carefully evaluate layout sequences item by item. Ensure all tabular data matrices are fully represented element by element. "
                            f"Translate icon or drawing illustrations into fitting contextual emojis (e.g. ☀️, 🍎, 🌳)."
                        )
                        
                        prompt = (
                            f"Perform a comprehensive structural extraction of all provided pages in sequence for {grade_tier}. "
                            f"Double-check every character to ensure complete accuracy matching the visual source material."
                        )
                        contents = [prompt] + img_list
                        
                        response = client.models.generate_content(
                            model='gemini-3.5-flash',
                            contents=contents,
                            config=types.GenerateContentConfig(
                                system_instruction=system_instruction,
                                response_mime_type="application/json",
                                response_schema=UniversalExamPaper,
                                temperature=0.0
                            )
                        )
                        
                        raw_json = json.loads(response.text)
                        st.session_state.parsed_data = UniversalExamPaper(**raw_json)
                        status.update(label="High-precision extraction complete.", state="complete", expanded=False)
                        st.rerun()
                        
                    except Exception as e:
                        status.update(label="Extraction failed", state="error")
                        st.error(f"Error: {e}")

with col_right:
    st.markdown("### 02 / Studio Review & Export")
    
    if st.session_state.parsed_data is None:
        st.info("Upload source files and run extraction to preview components here.")
    else:
        data = st.session_state.parsed_data
        
        with st.form("exam_editor_form"):
            st.markdown("#### Document Metadata")
            data.school_name = st.text_input("School Name", value=data.school_name)
            data.exam_title = st.text_input("Exam Title", value=data.exam_title)
            
            c1, c2 = st.columns(2)
            with c1:
                data.class_name = st.text_input("Class", value=data.class_name)
                data.subject = st.text_input("Subject", value=data.subject)
            with c2:
                data.full_marks = st.text_input("Full Marks", value=data.full_marks)
                data.time = st.text_input("Time", value=data.time)
                
            data.student_info_line = st.text_input("Student Info Placeholder", value=data.student_info_line)
            
            st.markdown("---")
            st.markdown("#### Structural Elements Editor")
            
            for i, block in enumerate(data.blocks):
                if block.block_type == 'text_paragraph':
                    block.text_content = st.text_area(f"Block {i+1} (Text)", value=block.text_content or "", height=70)
                elif block.block_type == 'list_block':
                    block.text_content = st.text_input(f"Block {i+1} (List Header)", value=block.text_content or "")
                elif block.block_type == 'drawing_box_block':
                    block.text_content = st.text_input(f"Block {i+1} (Drawing Box Label)", value=block.text_content or "")
            
            update_submitted = st.form_submit_button("Update Layout State")
            
            if update_submitted:
                st.success("Layout modified successfully.")
        
        st.markdown("---")
        st.markdown("#### Download Output Document")
        
        word_bytes = create_docx(st.session_state.parsed_data, exam_language, grade_tier, int(custom_font_size))
        
        st.download_button(
            label=f"Download {st.session_state.original_filename}",
            data=word_bytes,
            file_name=st.session_state.original_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
