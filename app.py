import streamlit as st
from google import genai
from google.genai import types
import json
from pydantic import BaseModel, Field
from typing import List, Optional
from PIL import Image
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# --- Page Setup ---
st.set_page_config(
    page_title="Question Paper Studio",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Session State Initialization ---
if "exam_data" not in st.session_state:
    st.session_state.exam_data = None
if "raw_json" not in st.session_state:
    st.session_state.raw_json = None

# --- Minimalist iOS / One UI Custom CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    /* Massive High-Impact Gradient Title */
    .main-title {
        font-size: clamp(3.5rem, 8vw, 6.5rem); /* Massive on desktop, scales down safely on mobile */
        font-weight: 900;
        text-align: center;
        background: linear-gradient(135deg, #007AFF 0%, #FF2D55 50%, #5856D6 100%);
        background-size: 200% auto;
        color: #fff;
        background-clip: text;
        text-fill-color: transparent;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: gradientFlow 5s ease infinite;
        margin-bottom: 5px;
        padding-top: 20px;
        letter-spacing: -2px;
        line-height: 1.1;
    }
    
    @keyframes gradientFlow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    /* Short, Stylized Explainer */
    .hero-explainer {
        text-align: center;
        font-size: 1.05rem; /* Scaled down appropriately */
        font-weight: 400;
        color: var(--text-color); /* Fixes Dark Mode: Automatically adapts to light/dark theme */
        opacity: 0.75; 
        max-width: 650px;
        margin: 15px auto 40px auto;
        line-height: 1.5;
        letter-spacing: -0.2px;
    }
    
    .hero-explainer strong {
        font-weight: 600;
        color: var(--text-color); /* Fixes Dark Mode: Automatically adapts to light/dark theme */
        opacity: 1; /* Makes the bold text pop slightly more */
    }

    .section-header {
        font-weight: 600;
        font-size: 1.25rem;
        margin-bottom: 15px;
        letter-spacing: -0.3px;
    }

    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #D1D1D6 !important;
        background-color: transparent !important;
        border-radius: 14px !important;
        padding: 40px 20px !important;
        transition: border-color 0.2s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #007AFF !important;
    }

    .stButton>button, .stDownloadButton>button {
        width: 100%;
        background-color: #007AFF !important; 
        color: white !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 12px 20px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        transition: all 0.2s ease;
        box-shadow: none !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #0056b3 !important;
        transform: scale(0.98);
    }
</style>
""", unsafe_allow_html=True)


# --- 1. Universal Layout Schema ---
class LayoutBlock(BaseModel):
    block_type: str = Field(description="Can be: 'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block', 'drawing_box_block'")
    text_content: Optional[str] = Field(default=None)
    list_items: Optional[List[str]] = Field(default=None)
    table_rows: Optional[int] = Field(default=None)
    table_cols: Optional[int] = Field(default=None)
    table_data: Optional[List[List[str]]] = Field(default=None)
    columns_data: Optional[List[List[str]]] = Field(default=None)
    box_height_inches: Optional[float] = Field(default=2.0)
    diagram_description: Optional[str] = Field(default=None, description="If this is a drawing box, describe what the original image showed so the user knows what to paste.")

class UniversalExamPaper(BaseModel):
    school_name: str
    exam_title: str
    class_name: str
    subject: str
    full_marks: str
    time: str
    student_info_line: str
    blocks: List[LayoutBlock] = Field(description="Chronological order of layout blocks")


# --- 2. Helper Functions ---
def optimize_image(img, max_width=1600):
    if img.width > max_width:
        ratio = max_width / img.width
        new_size = (max_width, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    return img.convert('RGB')

def set_table_borders(table, color="cccccc"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single'); border.set(qn('w:sz'), '4'); border.set(qn('w:space'), '0'); border.set(qn('w:color'), color)  
        tblBorders.append(border)
    tblPr.append(tblBorders)

def create_docx(data: UniversalExamPaper, language: str, font_size: int, margin_size: float):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(margin_size)
        section.bottom_margin = Inches(margin_size)
        section.left_margin = Inches(margin_size)
        section.right_margin = Inches(margin_size)
        
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
        if label.lower() in val.lower() or "শ্রেণী" in val or "कक्षा" in val or "পূর্ণমান" in val or "विषय" in val or "সময়" in val or "সময়" in val or "पूर्णांक" in val:
            return val
        return f"{label} — {val}"

    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.autofit = True
    
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
    p_div.add_run("—" * 60)
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
            if b.table_rows and b.table_cols:
                tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
                tbl.autofit = True
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
                col_tbl.autofit = True
                for c in range(num_cols):
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
            box_tbl.rows[0].height = Inches(b.box_height_inches or 2.0)
            set_table_borders(box_tbl, "888888")
            
            if b.diagram_description:
                cell_p = box_tbl.rows[0].cells[0].paragraphs[0]
                cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cell_p.add_run(f"[Insert Diagram Here: {b.diagram_description}]")
                run.font.color.rgb = RGBColor(150, 150, 150)
                
            doc.add_paragraph()

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. Web UI Body ---
st.markdown('<p class="main-title">Question Paper Studio</p>', unsafe_allow_html=True)

# Clean SVG Explainer Graphic
st.markdown("""
<div style="display: flex; justify-content: center; align-items: center; padding: 10px 0 10px 0; gap: 25px; opacity: 0.8;">
   <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="#8E8E93" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
      <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
      <circle cx="8.5" cy="8.5" r="1.5"></circle>
      <polyline points="21 15 16 10 5 21"></polyline>
   </svg>
   <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#007AFF" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <line x1="5" y1="12" x2="19" y2="12"></line>
      <polyline points="12 5 19 12 12 19"></polyline>
   </svg>
   <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#007AFF" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
      <path d="M9.5 2A2.5 2.5 0 0 1 12 4.5v15a2.5 2.5 0 0 1-4.96.44 2.5 2.5 0 0 1-2.96-3.08 3 3 0 0 1-.34-5.58 2.5 2.5 0 0 1 1.32-4.24 2.5 2.5 0 0 1 1.98-3A2.5 2.5 0 0 1 9.5 2Z"></path>
      <path d="M14.5 2A2.5 2.5 0 0 0 12 4.5v15a2.5 2.5 0 0 0 4.96.44 2.5 2.5 0 0 0 2.96-3.08 3 3 0 0 0 .34-5.58 2.5 2.5 0 0 0-1.32-4.24 2.5 2.5 0 0 0-1.98-3A2.5 2.5 0 0 0 14.5 2Z"></path>
   </svg>
   <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#007AFF" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <line x1="5" y1="12" x2="19" y2="12"></line>
      <polyline points="12 5 19 12 12 19"></polyline>
   </svg>
   <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="#8E8E93" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
      <polyline points="14 2 14 8 20 8"></polyline>
      <line x1="16" y1="13" x2="8" y2="13"></line>
      <line x1="16" y1="17" x2="8" y2="17"></line>
      <polyline points="10 9 9 9 8 9"></polyline>
   </svg>
</div>
""", unsafe_allow_html=True)

# Short, Stylized Explainer
st.markdown("""
<div class="hero-explainer">
    Transform messy handwritten and printed exams into <strong>perfectly structured, completely editable Word documents</strong> in seconds. Smart layout, multi-language support, zero formatting headaches.
</div>
""", unsafe_allow_html=True)


# Sidebar Settings
with st.sidebar:
    st.markdown('<p class="section-header">System Settings</p>', unsafe_allow_html=True)
    
    api_key = st.text_input("Gemini Pro API Key", type="password")
    if not api_key:
        st.warning("API key required to process documents.")
        
    exam_language = st.selectbox("Target Language", ["Bengali", "English", "Hindi"])
    
    ai_engine = st.radio("Processing Engine", ["Speed (Gemini Flash)", "High Accuracy (Gemini Pro)"])
    
    st.markdown("---")
    st.markdown('<p class="section-header">Document Styling</p>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        custom_font_size = st.number_input("Font Size (Pt)", min_value=8, max_value=18, value=11)
    with col_b:
        custom_margin = st.number_input("Margins (In)", min_value=0.5, max_value=2.0, value=0.75, step=0.1)

    st.markdown("---")
    st.info("Tip: Adjust font and margins anytime after compiling. The document will update instantly.")

# Main Interactive Workspace
col1, col_space, col2 = st.columns([1.2, 0.1, 1])

with col1:
    st.markdown('<p class="section-header">1. Upload Files</p>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        f"Drop {exam_language} image files here", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
    if uploaded_files and st.session_state.exam_data and len(uploaded_files) != getattr(st.session_state, 'last_upload_count', 0):
        st.session_state.exam_data = None
        st.session_state.raw_json = None
    st.session_state.last_upload_count = len(uploaded_files)

with col2:
    st.markdown('<p class="section-header">2. Process Images</p>', unsafe_allow_html=True)
    if not uploaded_files:
        st.info("Please drop your image files in the dropzone to continue.")
    else:
        st.success(f"{len(uploaded_files)} Pages Ready for Processing")
        
        if st.button("Compile Document"):
            if not api_key:
                st.error("Please provide your Gemini API Key in the sidebar.")
            else:
                with st.status(f"Parsing structure with AI...", expanded=True) as status:
                    try:
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        
                        st.write("Optimizing images for processing...")
                        img_list = [optimize_image(Image.open(f)) for f in sorted_files]
                        
                        client = genai.Client(api_key=api_key)
                        
                        system_instruction = (
                            f"You are an expert, highly precise document OCR parser specialized in {exam_language} exam papers. "
                            f"RULES: "
                            f"1. Extract text exactly as written. Do not solve the questions. "
                            f"2. Handle messy handwriting accurately. If completely illegible, write '[ILLEGIBLE]'. "
                            f"3. For images, diagrams, or graphs, use 'drawing_box_block'. You MUST write a brief description of what the image shows in the 'diagram_description' field. "
                            f"4. Map visual formats strictly to the 'blocks' array: 'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block', or 'drawing_box_block'."
                        )
                        
                        prompt = f"Analyze all pages in order. Extract layout structure and text completely faithfully in {exam_language}."
                        contents = [prompt] + img_list
                        
                        st.write("Running high-fidelity OCR scanning...")
                        
                        # Route based on user selection
                        if "Pro" in ai_engine:
                            fallback_models = ['gemini-1.5-pro', 'gemini-2.5-pro', 'gemini-3.5-pro']
                        else:
                            fallback_models = ['gemini-3.5-flash', 'gemini-1.5-flash', 'gemini-2.5-flash']
                            
                        response = None
                        last_error = None

                        for model_name in fallback_models:
                            try:
                                response = client.models.generate_content(
                                    model=model_name,
                                    contents=contents,
                                    config=types.GenerateContentConfig(
                                        system_instruction=system_instruction,
                                        response_mime_type="application/json",
                                        response_schema=UniversalExamPaper,
                                        temperature=0.1
                                    )
                                )
                                break
                            except Exception as e:
                                error_msg = str(e)
                                if "503" in error_msg or "UNAVAILABLE" in error_msg or "429" in error_msg:
                                    st.warning(f"{model_name} is currently busy. Rerouting to backup server...")
                                    last_error = error_msg
                                    continue
                                else:
                                    raise e
                                    
                        if not response:
                            raise Exception(f"All backup servers are currently busy. Please try again in a few minutes. (Last Error: {last_error})")
                        
                        raw_json = json.loads(response.text)
                        st.session_state.exam_data = UniversalExamPaper(**raw_json)
                        st.session_state.raw_json = raw_json
                        
                        status.update(label="Process Complete", state="complete", expanded=False)
                        
                    except Exception as e:
                        status.update(label="Compile Interrupted", state="error")
                        st.error(f"Diagnostics: {e}")

        if st.session_state.exam_data:
            st.markdown("---")
            st.success("Document Ready. Adjust Font Size and Margins in the sidebar; the download will update automatically.")
            
            word_bytes = create_docx(
                st.session_state.exam_data, 
                exam_language, 
                int(custom_font_size), 
                float(custom_margin)
            )
            
            st.download_button(
                label="Download Stylized Document",
                data=word_bytes,
                file_name=f"Digitized_{exam_language}_Exam.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("View Live JSON Data"):
                st.json(st.session_state.raw_json)
