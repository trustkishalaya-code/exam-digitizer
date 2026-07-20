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

# --- Page Setup (MUST BE FIRST) ---
st.set_page_config(
    page_title="AI Exam Digitizer Pro",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 🎨 Premium Custom Typography & UI Layout ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700;900&family=Inter:wght@300;400;500;600&display=swap');

    /* Global Typography Customizations */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
    }
    
    h1, h2, h3, .main-title, .section-header {
        font-family: 'Space Grotesk', sans-serif !important;
    }
    
    /* Cosmic Neon Gradient Title Layout */
    .main-title {
        font-size: 3.8rem;
        font-weight: 900;
        text-align: center;
        background: linear-gradient(135deg, #FF007A 0%, #7928CA 50%, #00DFD8 100%);
        background-size: 200% auto;
        color: #fff;
        background-clip: text;
        text-fill-color: transparent;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: gradientFlow 4s ease infinite;
        margin-bottom: -5px;
        padding-top: 15px;
        letter-spacing: -1px;
    }
    
    @keyframes gradientFlow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    .sub-title {
        text-align: center;
        color: #94A3B8;
        font-size: 1.15rem;
        font-weight: 400;
        margin-bottom: 35px;
        letter-spacing: 0.2px;
    }

    /* Translucent Glass Sidebar */
    [data-testid="stSidebar"] {
        background: rgba(15, 23, 42, 0.4) !important;
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.08);
    }

    /* Interactive Upload Dropzone Wrapper */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #7928CA !important;
        background-color: rgba(121, 40, 202, 0.03) !important;
        border-radius: 16px !important;
        padding: 40px 20px !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: rgba(0, 223, 216, 0.04) !important;
        border-color: #00DFD8 !important;
        transform: translateY(-2px);
        box-shadow: 0 8px 30px rgba(0, 223, 216, 0.15);
    }

    /* Core Compile Button Accent styling */
    .stButton>button {
        width: 100%;
        background-image: linear-gradient(135deg, #7928CA 0%, #FF007A 100%);
        padding: 14px 20px;
        text-align: center;
        transition: 0.4s;
        background-size: 150% auto;
        color: white !important;
        border-radius: 12px;
        border: none;
        font-weight: 700;
        font-family: 'Space Grotesk', sans-serif !important;
        letter-spacing: 0.5px;
        box-shadow: 0 4px 15px rgba(121, 40, 202, 0.3);
    }
    .stButton>button:hover {
        background-position: right center;
        box-shadow: 0 8px 25px rgba(255, 0, 122, 0.5);
        transform: translateY(-2px);
    }
    
    /* High Contrast Success Download Action */
    .stDownloadButton>button {
        width: 100%;
        background-image: linear-gradient(135deg, #00DFD8 0%, #00F260 100%);
        padding: 14px 20px;
        text-align: center;
        transition: 0.4s;
        background-size: 150% auto;
        color: #0F172A !important;
        border-radius: 12px;
        border: none;
        font-weight: 700;
        font-family: 'Space Grotesk', sans-serif !important;
        letter-spacing: 0.5px;
        box-shadow: 0 4px 15px rgba(0, 223, 216, 0.3);
    }
    .stDownloadButton>button:hover {
        background-position: right center;
        box-shadow: 0 8px 25px rgba(0, 242, 96, 0.5);
        transform: translateY(-2px);
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
    blocks: List[LayoutBlock] = Field(description="Chronological sequence of structured objects found")


# --- 2. Word Engine with Native A4 / Narrow Layout Adjustments ---
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
    
    # 📏 Strict Standard Layout: A4 Sheet Dimensions + Narrow (0.5 inch) Margins
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
    
    # Header Parsing Engine
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
    
    # Metadata Setup Table
    class_label = "Class" if language == "English" else ("শ্রেণী" if language == "Bengali" else "कक्षा")
    marks_label = "Full Marks" if language == "English" else ("পূর্ণমান" if language == "Bengali" else "पूर्णांक")
    subject_label = "Subject" if language == "English" else ("विषय" if language == "Bengali" else "विषय")
    time_label = "Time" if language == "English" else ("সময়" if language == "Bengali" else "समय")

    def format_meta(label, val):
        if not val: return ""
        if label.lower() in val.lower() or "শ্রেণী" in val or "कक्षा" in val or "পূর্ণমান" in val or "विषय" in val or "সময়" in val or "पूर्णांक" in val:
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
    
    # Process Chronological Layout Blocks
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
            if b.table_rows and b.table_cols:
                tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
                tbl.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(tbl, "cccccc")
                if b.table_data:
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

# --- 3. Interactive Front-End Layout ---
st.markdown('<p class="main-title">AI Vision Matrix</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Sleek, A4-Standard precision digitizer optimized for primary and secondary school exams.</p>', unsafe_allow_html=True)
st.write("")

# Configuration Engine Sidebar
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #00DFD8; font-family: Space Grotesk;'>🔮 System Engine</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    api_key = st.text_input("🔑 Gemini Pro API Key", type="password")
    if not api_key:
        st.warning("⚠️ Access Key required to authorize deployment.")
        
    exam_language = st.selectbox("🌐 Target Language", ["Bengali", "English", "Hindi"])
    
    st.markdown("---")
    st.markdown("<h3 style='color: #FF007A; font-family: Space Grotesk;'>🎛️ Document Styling</h3>", unsafe_allow_html=True)
    custom_font_size = st.number_input("Font Size (Pt)", min_value=8, max_value=18, value=11)
    
    st.markdown("---")
    st.caption("⚙️ **Forced Layout Profiles Enabled:** Page Formats are strictly configured to **A4 Profile** utilizing modern **Narrow Borders (0.5 in)**.")

# Main Operational Viewport
col1, col_space, col2 = st.columns([1.2, 0.1, 1])

with col1:
    st.markdown("<h3 style='color: #7928CA;'>📸 1. Input Processing Stream</h3>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        f"Drop {exam_language} image segments here", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

with col2:
    st.markdown("<h3 style='color: #FF007A;'>⚡ 2. Transformation Vector</h3>", unsafe_allow_html=True)
    if not uploaded_files:
        st.info("👈 System idle. Feed raw paper matrices to initialize parsing sequence.")
    else:
        st.success(f"📂 {len(uploaded_files)} Page Channels Registered & Online")
        
        if st.button("✨ Compile Document Matrix"):
            if not api_key:
                st.error("Operation Aborted: Missing API credential matrix in configuration.")
            else:
                with st.status(f"🧠 Initiating Vision Intelligence OCR...", expanded=True) as status:
                    try:
                        # 🧬 Step A: Chronological Array Sorting & High-Contrast Visual Preprocessing
                        st.write("🌌 Enhancing image channel contrast values by 1.5x...")
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        img_list = []
                        
                        for f in sorted_files:
                            raw_img = Image.open(f)
                            # Boost contrast aggressively to sharpen low-quality hand drawings or faded text
                            enhancer = ImageEnhance.Contrast(raw_img)
                            enhanced_img = enhancer.enhance(1.5)
                            img_list.append(enhanced_img)
                        
                        client = genai.Client(api_key=api_key)
                        
                        # 🎯 Step B: Upgraded Precision + Emoji-mapping Directives
                        system_instruction = (
                            f"You are a master-level, highly accurate document OCR parser specialized in parsing {exam_language} exam papers for young elementary students. "
                            f"Your primary directive is 100% literal text accuracy. DO NOT summarize, paraphrase, or leave out words. Transcribe exactly as written, "
                            f"preserving all original punctuation, sentence styling, numbering, and mathematical operators.\n\n"
                            f"🖼️ ELEMENTARY ILLUSTRATION TRANSLATION RULE: These exam papers are for lower classes and contain basic drawings. "
                            f"If you see basic vector-style drawings or illustrations representing common objects (e.g., a sun, tree, ball, star, apple, cat, mango, flower), "
                            f"DO NOT replace them with an empty drawing box block. Instead, instantly translate that illustration into its closest matching Emoji or Unicode character matrix "
                            f"(e.g., ☀️, 🌳, ⚽, ⭐, 🍎, 🐈, 🥭, 🌸). Insert these emojis directly within the 'text_content' or 'list_items' exactly where the image appears on the physical paper."
                        )
                        
                        prompt = (
                            f"Analyze all processed pages sequentially. Extract structural layout blocks and all text contents with extreme fidelity in {exam_language}. "
                            f"Pay absolute attention to word spelling, text sequences, and object counts (e.g., if a question has a drawing of 3 stars next to it, output '⭐ ⭐ ⭐')."
                        )
                        
                        contents = [prompt] + img_list
                        
                        st.write("🧠 Executing deep layout parsing matrix...")
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
                        
                        # 📝 Step C: Map Schema Directly into A4 Styled Word Document
                        st.write("📝 Structuralizing output stream into Microsoft Word container...")
                        raw_json = json.loads(response.text)
                        exam_data = UniversalExamPaper(**raw_json)
                        word_bytes = create_docx(exam_data, exam_language, int(custom_font_size))
                        
                        status.update(label="🔮 Execution Completed Successfully!", state="complete", expanded=False)
                        
                        # Stream Dynamic Download Action
                        st.download_button(
                            label="📥 Download A4 Digitized Document",
                            data=word_bytes,
                            file_name=f"Digitized_{exam_language}_Exam.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        st.balloons()
                        
                        # 🔬 Step D: Real-time X-Ray Data Inspector Panel
                        st.markdown("<br>", unsafe_allow_html=True)
                        with st.expander("👀 View Visual X-Ray Data Preview"):
                            st.json(raw_json)
                        
                    except Exception as e:
                        status.update(label="❌ Matrix Processing Failure", state="error")
                        st.error(f"Error Report: {e}")
