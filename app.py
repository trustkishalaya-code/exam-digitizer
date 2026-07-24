import io
import json
from typing import List, Optional
from pydantic import BaseModel, Field
from PIL import Image, ImageEnhance
from google import genai
from google.genai import types
import streamlit as st

# --- PDF ENGINE IMPORTS ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Academic Studio // Brutalism Edition",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- THE BRUTALIST CSS W/ FOCUS UPGRADES ---
st.markdown(
    """
<style>
    /* Global Reset & Base */
    .stApp {
        background-color: #000000;
        color: #FFFFFF;
        font-family: 'Courier New', Courier, monospace;
    }
    
    /* Hide default elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Brutalist Container Box */
    .brut-box {
        background: #000000;
        border: 3px solid #FFEA00;
        padding: 20px;
        border-radius: 0px;
        box-shadow: 6px 6px 0px #FFEA00;
        margin-bottom: 20px;
    }

    /* Header Banner */
    .brut-header {
        background: #FFEA00;
        color: #000000;
        padding: 24px;
        border: 3px solid #FFEA00;
        font-weight: 900;
        margin-bottom: 24px;
        box-shadow: 6px 6px 0px #FFFFFF;
    }

    /* Form Labels */
    .brut-label {
        font-weight: bold;
        font-size: 9.5pt;
        color: #FFEA00;
        display: block;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }

    /* Input Fields Style */
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stNumberInput input {
        border: 2px solid #FFFFFF !important;
        background-color: #000000 !important;
        color: #FFFFFF !important;
        font-family: 'Courier New', Courier, monospace !important;
        font-weight: bold !important;
        border-radius: 0px !important;
    }

    /* Focus States for Inputs */
    .stTextInput input:focus, .stSelectbox div[data-baseweb="select"]:focus-within, .stNumberInput input:focus {
        border-color: #FFEA00 !important;
        box-shadow: 4px 4px 0px #FFFFFF !important;
        outline: none !important;
    }

    /* Action Buttons */
    .stButton>button, .stDownloadButton>button {
        display: block;
        width: 100%;
        padding: 12px;
        border: 3px solid #FFEA00 !important;
        background: #FFEA00 !important;
        color: #000000 !important;
        font-family: 'Courier New', Courier, monospace !important;
        font-weight: 900 !important;
        font-size: 10pt;
        text-transform: uppercase;
        border-radius: 0px !important;
        box-shadow: 4px 4px 0px #FFFFFF !important;
        transition: transform 0.1s ease, box-shadow 0.1s ease !important;
    }
    
    .stButton>button:hover, .stDownloadButton>button:hover {
        transform: translate(-2px, -2px) !important;
        box-shadow: 6px 6px 0px #FFFFFF !important;
    }

    /* Override Streamlit Dropzone to match Brutal Dropzone */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #FFEA00 !important;
        padding: 25px !important;
        text-align: center !important;
        background: #000000 !important;
        border-radius: 0px !important;
        margin-bottom: 16px !important;
    }
    
    /* Text rendering */
    p, li, span, label { font-family: 'Courier New', Courier, monospace !important; color: #FFFFFF; }
</style>
""",
    unsafe_allow_html=True,
)

# --- PYDANTIC SCHEMAS ---
class LayoutBlock(BaseModel):
    block_type: str = Field(description="Can be: 'text_paragraph', 'list_block', 'grid_table_block', 'drawing_box_block'")
    text_content: Optional[str] = Field(default=None)
    list_items: Optional[List[str]] = Field(default=None)
    table_data: Optional[List[List[str]]] = Field(description="2D array matrix for grid_table_block", default=None)
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

# --- UTILITIES ---
def optimize_image(img, max_width=2500):
    if img.width > max_width:
        ratio = max_width / img.width
        new_size = (max_width, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    return img.convert("RGB")

# --- PDF ENGINE ---
def create_pdf(data: UniversalExamPaper, grade_tier: str, base_font_size: int):
    bio = io.BytesIO()
    
    is_early_childhood = grade_tier in ["Nursery", "PP / LKG / UKG"]
    page_size = A4 if is_early_childhood else landscape(A4)
    
    doc = SimpleDocTemplate(
        bio, 
        pagesize=page_size,
        rightMargin=0.5*inch, leftMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch
    )
    
    styles = getSampleStyleSheet()
    elements = []
    
    # Titles
    styles.add(ParagraphStyle(name='CenterTitle', alignment=1, fontSize=base_font_size + 4, fontName="Helvetica-Bold", spaceAfter=6))
    elements.append(Paragraph(data.school_name, styles['CenterTitle']))
    
    styles.add(ParagraphStyle(name='CenterSubtitle', alignment=1, fontSize=base_font_size + 2, fontName="Helvetica-Bold", spaceAfter=12))
    elements.append(Paragraph(data.exam_title, styles['CenterSubtitle']))
    
    # Metadata Table
    meta_data = [
        [f"Class: {data.class_name}", f"Full Marks: {data.full_marks}"],
        [f"Subject: {data.subject}", f"Time: {data.time}"]
    ]
    
    t = Table(meta_data, colWidths=[doc.width/2.0]*2)
    t.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), base_font_size),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 0.2*inch))
    
    # Student Info & Divider
    elements.append(Paragraph(data.student_info_line, styles['Normal']))
    elements.append(Spacer(1, 0.1*inch))
    elements.append(Paragraph("-" * 80, styles['CenterTitle']))
    elements.append(Spacer(1, 0.2*inch))
    
    # Content Blocks
    for b in data.blocks:
        if b.block_type == "text_paragraph" and b.text_content:
            elements.append(Paragraph(b.text_content, styles['Normal']))
            elements.append(Spacer(1, 0.1*inch))
            
        elif b.block_type == "list_block" and b.list_items:
            if b.text_content:
                elements.append(Paragraph(b.text_content, styles['Normal']))
            for item in b.list_items:
                elements.append(Paragraph(f"• {item}", styles['Normal']))
            elements.append(Spacer(1, 0.1*inch))
            
        elif b.block_type == "grid_table_block" and b.table_data:
            if b.text_content:
                elements.append(Paragraph(b.text_content, styles['Normal']))
            grid_table = Table(b.table_data)
            grid_table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,0), (-1,-1), base_font_size),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            elements.append(grid_table)
            elements.append(Spacer(1, 0.15*inch))
            
        elif b.block_type == "drawing_box_block":
            if b.text_content:
                elements.append(Paragraph(b.text_content, styles['Normal']))
            box = Table([[""]], colWidths=[doc.width], rowHeights=[(b.box_height_inches or 2.0) * inch])
            box.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black)]))
            elements.append(box)
            elements.append(Spacer(1, 0.15*inch))

    doc.build(elements)
    return bio.getvalue()

# --- STATE MANAGEMENT ---
if "parsed_data" not in st.session_state:
    st.session_state.parsed_data = None
if "original_filename" not in st.session_state:
    st.session_state.original_filename = "Exam_Output.pdf"

# --- UI RENDER ---
st.markdown(
    """
<div class="brut-header">
    <div style="font-size: 8pt; font-weight: bold; letter-spacing: 2px; margin-bottom: 4px;">STYLE 3: HIGH-CONTRAST BRUTALISM (ELECTRIC YELLOW)</div>
    <div style="font-size: 24pt; font-weight: bold; letter-spacing: -1px;">ACADEMIC_STUDIO//</div>
    <div style="font-size: 9.5pt; font-weight: bold; margin-top: 4px;">SUPERCHARGED DIGITIZER & LAYOUT COMPILER ⚡</div>
</div>
""",
    unsafe_allow_html=True,
)

# Sidebar
with st.sidebar:
    st.markdown('<div class="brut-box"><h3>[ CONTROLS ]</h3>', unsafe_allow_html=True)

    st.markdown('<label class="brut-label">API KEY</label>', unsafe_allow_html=True)
    api_key = st.text_input("API KEY", type="password", label_visibility="collapsed")
    
    st.markdown('<label class="brut-label">TARGET GRADE TIER</label>', unsafe_allow_html=True)
    grade_tier = st.selectbox("TARGET GRADE TIER", ["Nursery", "PP / LKG / UKG", "Classes 1 to 4"], label_visibility="collapsed")
    
    st.markdown('<label class="brut-label">BASE FONT SIZE (PT)</label>', unsafe_allow_html=True)
    custom_font_size = st.number_input("BASE FONT SIZE", min_value=9, max_value=24, value=12, label_visibility="collapsed")

    st.write("")
    if st.button("[ CLEAR WORKSPACE ]"):
        st.session_state.parsed_data = None
        st.session_state.original_filename = "Exam_Output.pdf"
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# Main Workspace Columns
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.markdown('<div class="brut-box"><h3>[ 01. INTAKE STREAM ]</h3>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Upload Exam Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    
    if uploaded_files:
        st.session_state.original_filename = uploaded_files[0].name.rsplit(".", 1)[0] + ".pdf"
        
        if st.button("[ RUN HIGH-PRECISION EXTRACTION ]"):
            if not api_key:
                st.error("⚠️ Access Key required to launch.")
            else:
                with st.spinner("⚡ Extracting data via Gemini..."):
                    try:
                        sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                        img_list = []
                        for f in sorted_files:
                            raw_img = Image.open(f)
                            enhancer = ImageEnhance.Contrast(raw_img)
                            img_list.append(optimize_image(enhancer.enhance(1.5)))

                        client = genai.Client(api_key=api_key)
                        prompt = f"Extract the exam paper structure strictly according to the schema for {grade_tier}."
                        
                        response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=[prompt] + img_list,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                response_schema=UniversalExamPaper,
                                temperature=0.0,
                            ),
                        )
                        
                        raw_json = json.loads(response.text)
                        st.session_state.parsed_data = UniversalExamPaper(**raw_json)
                        st.rerun()

                    except Exception as e:
                        st.error(f"Extraction failed: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

with col_right:
    st.markdown('<div class="brut-box"><h3>[ 02. STUDIO REVIEW & EXPORT ]</h3>', unsafe_allow_html=True)

    if st.session_state.parsed_data is None:
        st.info("Upload source files and run extraction to unlock export protocols.")
    else:
        data = st.session_state.parsed_data
        
        with st.form("exam_editor_form"):
            st.markdown('<label class="brut-label">DOCUMENT METADATA</label>', unsafe_allow_html=True)
            data.school_name = st.text_input("School Name", value=data.school_name)
            data.exam_title = st.text_input("Exam Title", value=data.exam_title)
            
            c1, c2 = st.columns(2)
            with c1:
                data.class_name = st.text_input("Class", value=data.class_name)
                data.subject = st.text_input("Subject", value=data.subject)
            with c2:
                data.full_marks = st.text_input("Full Marks", value=data.full_marks)
                data.time = st.text_input("Time", value=data.time)
                
            data.student_info_line = st.text_input("Student Info Line", value=data.student_info_line)
            
            if st.form_submit_button("[ 💾 SAVE METADATA ]"):
                st.success("Metadata locked in.")
        
        st.divider()
        st.markdown('<label class="brut-label">COMPILE TO PDF</label>', unsafe_allow_html=True)
        
        pdf_bytes = create_pdf(data, grade_tier, int(custom_font_size))
        
        st.download_button(
            label=f"[ ⬇️ DOWNLOAD {st.session_state.original_filename.upper()} ]",
            data=pdf_bytes,
            file_name=st.session_state.original_filename,
            mime="application/pdf"
        )

    st.markdown("</div>", unsafe_allow_html=True)
