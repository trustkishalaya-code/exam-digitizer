import streamlit as st
from google import genai
from google.genai import types
import json
from pydantic import BaseModel, Field
from typing import List, Optional
from PIL import Image
import docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# --- Page Setup (MUST BE FIRST) ---
st.set_page_config(
    page_title="Exam Digitizer Ultra",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 🎨 Ultra-Premium Custom CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;900&display=swap');

    /* Global Font & Subtle Tweaks */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif !important;
    }
    
    /* Animated Gradient Main Title */
    .main-title {
        font-size: 4rem;
        font-weight: 900;
        text-align: center;
        background: linear-gradient(to right, #00F260, #0575E6, #00F260);
        background-size: 200% auto;
        color: #fff;
        background-clip: text;
        text-fill-color: transparent;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: shine 3s linear infinite;
        margin-bottom: -10px;
        padding-top: 20px;
    }
    
    @keyframes shine {
        to { background-position: 200% center; }
    }
    
    /* Subtitle */
    .sub-title {
        text-align: center;
        color: #8892b0;
        font-size: 1.2rem;
        font-weight: 400;
        margin-bottom: 40px;
    }

    /* Glassmorphism for Sidebar */
    [data-testid="stSidebar"] {
        background: rgba(15, 23, 42, 0.05) !important;
        backdrop-filter: blur(15px);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }

    /* Sexy Interactive File Uploader */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #0575E6 !important;
        background-color: rgba(5, 117, 230, 0.03) !important;
        border-radius: 20px !important;
        padding: 30px !important;
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: rgba(5, 117, 230, 0.08) !important;
        border-color: #00F260 !important;
        transform: translateY(-3px);
        box-shadow: 0 10px 25px rgba(5, 117, 230, 0.2);
    }

    /* Primary Generate Button (Pulse & Glow) */
    .stButton>button {
        width: 100%;
        background-image: linear-gradient(to right, #4776E6 0%, #8E54E9 51%, #4776E6 100%);
        margin: 10px 0px;
        padding: 15px 20px;
        text-align: center;
        text-transform: uppercase;
        transition: 0.5s;
        background-size: 200% auto;
        color: white !important;
        border-radius: 12px;
        border: none;
        font-weight: 900;
        letter-spacing: 1px;
        box-shadow: 0 4px 15px rgba(142, 84, 233, 0.4);
    }
    .stButton>button:hover {
        background-position: right center;
        box-shadow: 0 8px 25px rgba(142, 84, 233, 0.6);
        transform: translateY(-2px);
    }
    
    /* Download Button (Neon Green Success) */
    .stDownloadButton>button {
        width: 100%;
        background-image: linear-gradient(to right, #11998e 0%, #38ef7d 51%, #11998e 100%);
        margin: 10px 0px;
        padding: 15px 20px;
        text-align: center;
        text-transform: uppercase;
        transition: 0.5s;
        background-size: 200% auto;
        color: white !important;
        border-radius: 12px;
        border: none;
        font-weight: 900;
        letter-spacing: 1px;
        box-shadow: 0 4px 15px rgba(56, 239, 125, 0.4);
    }
    .stDownloadButton>button:hover {
        background-position: right center;
        box-shadow: 0 8px 25px rgba(56, 239, 125, 0.6);
        transform: translateY(-2px);
    }
    
    /* Info/Success/Warning Cards Styling */
    .stAlert {
        border-radius: 15px !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)


# --- 1. Define the Universal Layout Schema ---
class LayoutBlock(BaseModel):
    block_type: str = Field(description="Can be: 'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block', 'drawing_box_block'")
    text_content: Optional[str] = Field(default=None)
    list_items: Optional[List[str]] = Field(default=None)
    table_rows: Optional[int] = Field(default=None)
    table_cols: Optional[int] = Field(default=None)
    table_data: Optional[List[List[str]]] = Field(default=None)
    columns_data: Optional[List[List[str]]] = Field(default=None)
    box_height_inches: Optional[float] = Field(default=1.5)

class UniversalExamPaper(BaseModel):
    school_name: str
    exam_title: str
    class_name: str
    subject: str
    full_marks: str
    time: str
    student_info_line: str = Field(description="Student details line placeholder")
    blocks: List[LayoutBlock] = Field(description="Chronological order of layout blocks")


# --- 2. Helper Functions for Word Styling ---
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

def create_docx(data: UniversalExamPaper, language: str):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    font_mapping = {"Bengali": "Kalpurush", "Hindi": "Mangal", "English": "Calibri"}
    selected_font = font_mapping.get(language, "Calibri")
        
    style = doc.styles['Normal']
    style.font.name = selected_font
    style.font.size = Pt(11)
    
    # Header
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_school = p_school.add_run(data.school_name)
    run_school.bold = True
    run_school.font.name = selected_font
    run_school.font.size = Pt(15)
    p_school.paragraph_format.space_after = Pt(2)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(data.exam_title)
    run_title.bold = True
    run_title.font.name = selected_font
    run_title.font.size = Pt(12)
    p_title.paragraph_format.space_after = Pt(8)
    
    # Metadata Grid
    class_label = "Class" if language == "English" else ("শ্রেণী" if language == "Bengali" else "कक्षा")
    marks_label = "Full Marks" if language == "English" else ("পূর্ণমান" if language == "Bengali" else "पूर्णांक")
    subject_label = "Subject" if language == "English" else ("বিষয়" if language == "Bengali" else "विषय")
    time_label = "Time" if language == "English" else ("সময়" if language == "Bengali" else "समय")

    def format_meta(label, val):
        if not val: return ""
        if label.lower() in val.lower() or "শ্রেণী" in val or "कक्षा" in val or "পূর্ণমান" in val or "विषय" in val or "সময়" in val or "समय" in val or "पूर्णांक" in val:
            return val
        return f"{label} — {val}"

    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.autofit = False
    meta_table.columns[0].width = Inches(3.5)
    meta_table.columns[1].width = Inches(3.5)
    
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
                    run.font.name = selected_font
                    run.font.size = Pt(11)
                    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    # Student Details Line
    p_info = doc.add_paragraph()
    run_info = p_info.add_run(data.student_info_line)
    run_info.font.name = selected_font
    run_info.font.size = Pt(11)
    p_info.paragraph_format.space_after = Pt(8)
    
    # Divider
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_div = p_div.add_run("—" * 60)
    run_div.font.name = selected_font
    p_div.paragraph_format.space_after = Pt(12)
    
    # Layout Blocks
    for b in data.blocks:
        if b.block_type == 'text_paragraph':
            if b.text_content:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                is_num_start = b.text_content.strip().startswith((
                    '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯', '০',
                    '1', '2', '3', '4', '5', '6', '7', '8', '9', '0',
                    '१', '२', '३', '४', '५', '६', '७', '८', '९', '०'
                ))
                if "।" in b.text_content or ":" in b.text_content or is_num_start:
                    run.bold = True
                    
        elif b.block_type == 'list_block':
            if b.text_content:
                p = doc.add_paragraph()
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                run.bold = True
                p.paragraph_format.space_after = Pt(4)
            if b.list_items:
                for item in b.list_items:
                    lp = doc.add_paragraph()
                    lp.paragraph_format.left_indent = Inches(0.4)
                    lp.paragraph_format.space_after = Pt(4)
                    run_item = lp.add_run(item)
                    run_item.font.name = selected_font
                    
        elif b.block_type == 'grid_table_block':
            if b.text_content:
                p = doc.add_paragraph()
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                run.bold = True
                p.paragraph_format.space_after = Pt(4)
            if b.table_rows and b.table_cols:
                tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
                tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(tbl, "cccccc")
                if b.table_data:
                    for r_idx, row_items in enumerate(b.table_data):
                        if r_idx < b.table_rows:
                            for c_idx, val in enumerate(row_items):
                                if c_idx < b.table_cols:
                                    cell = tbl.rows[r_idx].cells[c_idx]
                                    cell.paragraphs[0].text = val
                                    for r_run in cell.paragraphs[0].runs:
                                        r_run.font.name = selected_font
                doc.add_paragraph().paragraph_format.space_after = Pt(8)
                
        elif b.block_type == 'column_layout_block':
            if b.text_content:
                p = doc.add_paragraph()
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                run.bold = True
                p.paragraph_format.space_after = Pt(4)
            if b.columns_data:
                num_cols = len(b.columns_data)
                max_rows = max(len(col) for col in b.columns_data)
                col_tbl = doc.add_table(rows=max_rows, cols=num_cols)
                col_tbl.autofit = False
                col_width = Inches(7.0 / num_cols)
                for c in range(num_cols):
                    col_tbl.columns[c].width = col_width
                    col_items = b.columns_data[c]
                    for r in range(max_rows):
                        cell = col_tbl.rows[r].cells[c]
                        if r < len(col_items):
                            cell.paragraphs[0].text = col_items[r]
                            cell.paragraphs[0].paragraph_format.left_indent = Inches(0.1)
                            for r_run in cell.paragraphs[0].runs:
                                r_run.font.name = selected_font
                doc.add_paragraph().paragraph_format.space_after = Pt(8)
                
        elif b.block_type == 'drawing_box_block':
            if b.text_content:
                p = doc.add_paragraph()
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                run.bold = True
                p.paragraph_format.space_after = Pt(4)
            box_tbl = doc.add_table(rows=1, cols=1)
            box_tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
            box_tbl.rows[0].height = Inches(b.box_height_inches or 1.5)
            set_table_borders(box_tbl, "888888")
            doc.add_paragraph().paragraph_format.space_after = Pt(8)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. Web UI Body ---

st.markdown('<p class="main-title">AI Vision Matrix</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Instantly materialize printed exams into beautifully structured Word files.</p>', unsafe_allow_html=True)
st.write("") # Spacer

# Sidebar Settings
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #00F260;'>🔮 Setup</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    api_key = st.text_input("🔑 Gemini Pro API Key", type="password", help="Paste your Google AI Studio key here.")
    if not api_key:
        st.warning("⚠️ API Key is required to run the AI engine.")
        
    exam_language = st.selectbox(
        "🌐 Target Language",
        ["Bengali", "English", "Hindi"],
        index=0,
        help="Select the language written in the images."
    )
    
    st.markdown("---")
    st.info("📌 **Pro-Tips**\n\n1️⃣ File names should be sorted chronologically (e.g., `page1.jpg`, `page2.jpg`).\n\n2️⃣ Ensure images are well-lit with clear text.")

# Main Interactive Area (Using wider columns for the new layout)
col1, col_space, col2 = st.columns([1.2, 0.1, 1])

with col1:
    st.markdown("<h3 style='color: #4776E6;'>📸 1. Feed the Matrix</h3>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        f"Drop {exam_language} image files here", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

with col2:
    st.markdown("<h3 style='color: #8E54E9;'>⚡ 2. Initialize Core</h3>", unsafe_allow_html=True)
    if not uploaded_files:
        st.info("👈 Waiting for visual data input on the left...")
    else:
        st.success(f"📦 {len(uploaded_files)} Data Blocks Ready")
        
        if st.button("✨ Digitize Now"):
            if not api_key:
                st.error("Access Denied: Missing API Key in sidebar.")
            else:
                with st.status(f"🧠 Linking with Gemini AI for {exam_language} parsing...", expanded=True) as status:
                    try:
                        st.write("🔄 Aligning files chronologically...")
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        
                        st.write("👁️ Opening AI Vision protocols...")
                        img_list = [Image.open(f) for f in sorted_files]
                        
                        st.write("⚡ Executing Layout Extraction Algorithms (takes ~10 seconds)...")
                        client = genai.Client(api_key=api_key)
                        
                        system_instruction = (
                            f"You are an expert layout-agnostic document OCR parser specialized in parsing {exam_language} exam papers. "
                            f"Read header info from Page 1, and sequentialize all questions across all pages into standard structured layout blocks. "
                            f"Preserve original {exam_language} spellings and formatting. "
                            f"Map visual formats to 'blocks' array: "
                            f"'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block' (for MCQs/Matching), or 'drawing_box_block'."
                        )
                        
                        prompt = f"Analyze all these pages in order. The text extracted must be completely faithful to the {exam_language} words present."
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
                        
                        st.write("📝 Synthesizing Microsoft Word Document...")
                        raw_json = json.loads(response.text)
                        exam_data = UniversalExamPaper(**raw_json)
                        word_bytes = create_docx(exam_data, exam_language)
                        
                        status.update(label="✅ Operations Complete!", state="complete", expanded=False)
                        
                        # Generate Download
                        st.download_button(
                            label="📥 Download Structured Document",
                            data=word_bytes,
                            file_name=f"Digitized_{exam_language}_Exam.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        st.balloons()
                        
                    except Exception as e:
                        status.update(label="❌ Error in Processing", state="error")
                        st.error(f"Diagnostics: {e}")
