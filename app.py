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
    
    h1, h2, h3, .serif-title {
        font-family: 'Instrument Serif', serif !important;
    }
    
    .main-title {
        font-size: 3.5rem; font-weight: 400; text-align: left;
        color: #111315; letter-spacing: -0.5px; line-height: 1.1;
        margin-bottom: 0px;
    }
    
    .sub-title { 
        text-align: left; color: #606873; font-size: 1.05rem; 
        font-weight: 500; margin-bottom: 25px; 
    }

    [data-testid="stSidebar"] { 
        background: #F1F3F5 !important; 
        border-right: 1px solid #E2E8F0 !important; 
    }

    /* Bento Grid Cards */
    .bento-card {
        background: #FFFFFF;
        border: 2px solid #E2E8F0;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.02);
        margin-bottom: 20px;
    }

    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #CBD5E1 !important; 
        background-color: #FAFAFA !important;
        border-radius: 12px !important; 
        padding: 30px 20px !important; 
        transition: all 0.2s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: #F1F5F9 !important; 
        border-color: #0F172A !important;
    }

    .stButton>button {
        width: 100%; 
        background-color: #0F172A;
        padding: 12px 20px; 
        text-align: center; 
        color: #FFFFFF !important; 
        border-radius: 10px; 
        border: none; 
        font-weight: 600;
        font-size: 0.95rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
        transition: transform 0.1s ease, background-color 0.2s ease;
    }
    .stButton>button:hover { 
        background-color: #334155; 
        transform: translateY(-1px); 
    }
    
    .stDownloadButton>button {
        width: 100%; 
        background-color: #0D9488;
        padding: 12px 20px; 
        text-align: center; 
        color: #FFFFFF !important; 
        border-radius: 10px; 
        border: none; 
        font-weight: 700;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.2);
        transition: transform 0.1s ease, background-color 0.2s ease;
    }
    .stDownloadButton>button:hover { 
        background-color: #0F766E; 
        transform: translateY(-1px); 
    }
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
    box_height_inches: Optional[float] = Field(default=1.5)

class UniversalExamPaper(BaseModel):
    school_name: str
    exam_title: str
    class_name: str
    subject: str
    full_marks: str
    time: str
    student_info_line: str = Field(description="Student details line placeholder (Name, Roll, etc.)")
    blocks: List[LayoutBlock] = Field(description="Chronological sequence of structured objects found")

# --- 2. Word Engine ---
def set_table_borders(table, color="cccccc"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single'); border.set(qn('w:sz'), '4'); border.set(qn('w:space'), '0'); border.set(qn('w:color'), color)  
        tblBorders.append(border)
    tblPr.append(tblBorders)

def create_docx(data: UniversalExamPaper, language: str, font_size: int):
    doc = Document()
    
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
    font_mapping = {"Bengali": "Kalpurush", "Hindi": "Mangal", "English": "Calibri"}
    selected_font = font_mapping.get(language, "Calibri")
        
    style = doc.styles['Normal']
    style.font.name = selected_font
    style.font.size = Pt(font_size)
    
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_school = p_school.add_run(data.school_name)
    run_school.bold = True
    run_school.font.size = Pt(font_size + 4)
    p_school.paragraph_format.space_after = Pt(2)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(data.exam_title)
    run_title.bold = True
    run_title.font.size = Pt(font_size + 1)
    p_title.paragraph_format.space_after = Pt(8)
    
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
    meta_table.columns[0].width = Inches(3.75)
    meta_table.columns[1].width = Inches(3.75)
    
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
                    run.font.size = Pt(font_size)
                    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    p_info = doc.add_paragraph()
    p_info.add_run(data.student_info_line).font.size = Pt(font_size)
    p_info.paragraph_format.space_after = Pt(8)
    
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("—" * 74)
    p_div.paragraph_format.space_after = Pt(12)
    
    for b in data.blocks:
        if b.block_type == 'text_paragraph' and b.text_content:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(b.text_content)
            is_num_start = b.text_content.strip().startswith(('১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯', '০', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '१', '२', '३', '४', '५', '६', '७', '८', '९', '०'))
            if "।" in b.text_content or ":" in b.text_content or is_num_start:
                run.bold = True
                
        elif b.block_type == 'list_block':
            if b.text_content:
                doc.add_paragraph().add_run(b.text_content).bold = True
            if b.list_items:
                for item in b.list_items:
                    lp = doc.add_paragraph()
                    lp.paragraph_format.left_indent = Inches(0.4)
                    lp.add_run(item)
                    
        elif b.block_type == 'grid_table_block':
            if b.text_content:
                doc.add_paragraph().add_run(b.text_content).bold = True
            if b.table_rows and b.table_cols and b.table_data:
                tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
                tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(tbl, "cccccc")
                for r_idx, row_items in enumerate(b.table_data):
                    if r_idx < b.table_rows:
                        for c_idx, val in enumerate(row_items):
                            if c_idx < b.table_cols:
                                tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
                doc.add_paragraph()
                
        elif b.block_type == 'column_layout_block':
            if b.text_content:
                doc.add_paragraph().add_run(b.text_content).bold = True
            if b.columns_data:
                num_cols = len(b.columns_data)
                max_rows = max(len(col) for col in b.columns_data)
                col_tbl = doc.add_table(rows=max_rows, cols=num_cols)
                col_width = Inches(7.5 / num_cols)
                for c in range(num_cols):
                    col_tbl.columns[c].width = col_width
                    col_items = b.columns_data[c]
                    for r in range(max_rows):
                        if r < len(col_items):
                            col_tbl.rows[r].cells[c].paragraphs[0].text = col_items[r]
                doc.add_paragraph()
                
        elif b.block_type == 'drawing_box_block':
            if b.text_content:
                doc.add_paragraph().add_run(b.text_content).bold = True
            box_tbl = doc.add_table(rows=1, cols=1)
            box_tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
            box_tbl.rows[0].height = Inches(b.box_height_inches or 1.5)
            set_table_borders(box_tbl, "888888")
            doc.add_paragraph()

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. UI Header & Sidebar Matrix ---
st.markdown('<p class="main-title">Academic Studio</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Editorial Document Digitizer & Layout Compiler</p>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h3 style='font-family: Instrument Serif; font-size: 1.8rem; color: #111315;'>Studio Controls</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    api_key = st.text_input("🔑 Gemini API Key", type="password")
    if not api_key:
        st.warning("⚠️ Access Key required.")
        
    exam_language = st.selectbox("🌐 Document Language", ["Bengali", "English", "Hindi"])
    
    st.markdown("---")
    st.markdown("<h3 style='font-family: Instrument Serif; font-size: 1.8rem; color: #111315;'>Formatting</h3>", unsafe_allow_html=True)
    custom_font_size = st.number_input("Base Font Size (Pt)", min_value=8, max_value=18, value=12)
    
    st.markdown("---")
    st.caption("📋 Standard A4 Layout Profile with Narrow Margins (0.5 in).")

# --- Session State Initialization ---
if "parsed_data" not in st.session_state:
    st.session_state.parsed_data = None
if "original_filename" not in st.session_state:
    st.session_state.original_filename = "Exam_Output.docx"

# --- Main Bento Layout Workspace ---
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
        
        # Thumbnail Grid Preview
        thumb_cols = st.columns(min(len(uploaded_files), 4))
        for idx, file in enumerate(uploaded_files[:4]):
            with thumb_cols[idx]:
                img_preview = Image.open(file)
                st.image(img_preview, use_container_width=True)
        
        st.write("")
        if st.button("Run Extraction Pipeline"):
            if not api_key:
                st.error("Missing API Key.")
            else:
                with st.status("Analyzing document structure...", expanded=True) as status:
                    try:
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        img_list = []
                        for f in sorted_files:
                            raw_img = Image.open(f)
                            enhanced = ImageEnhance.Sharpness(ImageEnhance.Contrast(raw_img).enhance(1.6)).enhance(2.0)
                            img_list.append(enhanced)
                        
                        client = genai.Client(api_key=api_key)
                        
                        system_instruction = (
                            f"You are an expert OCR parser for {exam_language} primary school exams (Nursery to Class 4). "
                            f"Extract literal content with extreme accuracy. Preserve dotted fill-in-the-blank lines (......) completely. "
                            f"Fully map out all tables and grids using 'grid_table_block' or 'column_layout_block' with populated `table_data`. "
                            f"Translate icon/drawing illustrations into matching Emojis (e.g. ☀️, 🍎, 🌳)."
                        )
                        
                        prompt = f"Analyze all pages sequentially and extract structured objects in {exam_language}."
                        contents = [prompt] + img_list
                        
                        response = client.models.generate_content(
                            model='gemini-3.5-flash',
                            contents=contents,
                            config=types.GenerateContentConfig(
                                system_instruction=system_instruction,
                                response_mime_type="application/json",
                                response_schema=UniversalExamPaper,
                                temperature=0.1
                            )
                        )
                        
                        raw_json = json.loads(response.text)
                        st.session_state.parsed_data = UniversalExamPaper(**raw_json)
                        status.update(label="Extraction complete. Review content on the right.", state="complete", expanded=False)
                        st.rerun()
                        
                    except Exception as e:
                        status.update(label="Extraction failed", state="error")
                        st.error(f"Error: {e}")

with col_right:
    st.markdown("### 02 / Studio Review & Export")
    
    if st.session_state.parsed_data is None:
        st.info("Upload source files and run extraction to preview and edit text components here.")
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
        word_bytes = create_docx(st.session_state.parsed_data, exam_language, int(custom_font_size))
        
        st.download_button(
            label=f"Download {st.session_state.original_filename}",
            data=word_bytes,
            file_name=st.session_state.original_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
