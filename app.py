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

# --- Page Setup ---
st.set_page_config(
    page_title="AI Multi-Lingual Exam Digitizer",
    page_icon="📝",
    layout="centered"
)

# --- 1. Define the Universal Layout Schema ---
class LayoutBlock(BaseModel):
    block_type: str = Field(
        description="Type of visual block. Can be: 'text_paragraph', 'list_block', 'grid_table_block', 'column_layout_block', 'drawing_box_block'"
    )
    text_content: Optional[str] = Field(default=None, description="Main text content, header, or question instruction.")
    list_items: Optional[List[str]] = Field(default=None, description="List of items/questions under this block.")
    table_rows: Optional[int] = Field(default=None, description="Number of rows required for the grid table.")
    table_cols: Optional[int] = Field(default=None, description="Number of columns required for the grid table.")
    table_data: Optional[List[List[str]]] = Field(default=None, description="Optional text content inside the table cells.")
    columns_data: Optional[List[List[str]]] = Field(default=None, description="Data list for each column. Can be 2, 3, or 4 columns wide.")
    box_height_inches: Optional[float] = Field(default=1.5, description="Height in inches of blank writing/drawing box.")

class UniversalExamPaper(BaseModel):
    school_name: str
    exam_title: str
    class_name: str
    subject: str
    full_marks: str
    time: str
    student_info_line: str = Field(description="Details like: নাম _________ রোল নং _________ or Name _________ Roll No _________ or नाम _________ अनुक्रमांक _________")
    blocks: List[LayoutBlock] = Field(description="Chronological order of all layout blocks extracted from the exam pages.")

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
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    # Font Mapping per Language
    font_mapping = {
        "Bengali": "Kalpurush",
        "Hindi": "Mangal",
        "English": "Calibri"
    }
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
    
    # Metadata Grid Label Settings
    class_label = "Class" if language == "English" else ("শ্রেণী" if language == "Bengali" else "कक्षा")
    marks_label = "Full Marks" if language == "English" else ("পূর্ণমান" if language == "Bengali" else "पूर्णांक")
    subject_label = "Subject" if language == "English" else ("বিষয়" if language == "Bengali" else "विषय")
    time_label = "Time" if language == "English" else ("সময়" if language == "Bengali" else "समय")

    def format_meta(label, val):
        if not val:
            return ""
        # Avoid prefixing label if Gemini already extracted it with a label
        if label.lower() in val.lower() or "শ্রেণী" in val or "कक्षा" in val or "পূর্ণমান" in val or "विषय" in val or "সময়" in val or "समय" in val or "पूर्णांक" in val:
            return val
        return f"{label} — {val}"

    # Setup Metadata Table
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
    
    # Divider Line
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_div = p_div.add_run("—" * 60)
    run_div.font.name = selected_font
    p_div.paragraph_format.space_after = Pt(12)
    
    # Render Dynamic Layout Blocks
    for b in data.blocks:
        if b.block_type == 'text_paragraph':
            if b.text_content:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(b.text_content)
                run.font.name = selected_font
                
                # Check for standard paragraph starters across Ben, Eng, and Hin to bold list headers
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

    # Save to Bytes buffer for online downloading
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. Web UI Design ---
st.title("📝 Universal Multilingual Exam Digitizer")
st.write("Convert any handwritten or printed exam papers in **Bengali, Hindi, or English** into a structured `.docx` file!")

# Sidebar for Settings
st.sidebar.header("🔑 Credentials")
api_key = st.sidebar.text_input("Gemini API Key", type="password", help="Get your free key at aistudio.google.com")

# Language Selection
st.sidebar.subheader("🌐 Exam Settings")
exam_language = st.sidebar.selectbox(
    "Select Exam Language",
    ["Bengali", "English", "Hindi"],
    index=0
)

st.info("💡 **Pro-Tip:** Make sure your files are sorted chronologically (e.g., `page1.jpg`, `page2.jpg`) before uploading so the layout remains in order!")

uploaded_files = st.file_uploader(
    f"Upload {exam_language} Exam Images (Choose multiple if needed)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"📂 Loaded {len(uploaded_files)} pages successfully!")
    
    if st.button("🚀 Process & Generate Word Document"):
        if not api_key:
            st.error("Please enter your Gemini API Key in the sidebar!")
        else:
            with st.spinner(f"🤖 Gemini is parsing images, matching {exam_language} formatting & structure..."):
                # Sort uploaded files by name
                sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                
                img_list = []
                for f in sorted_files:
                    img = Image.open(f)
                    img_list.append(img)
                
                client = genai.Client(api_key=api_key)
                
                # Dynamic instructions tuned to the chosen language
                system_instruction = (
                    f"You are an expert layout-agnostic document OCR parser specialized in parsing {exam_language} exam papers. "
                    f"Your task is to scan the uploaded exam papers, read the header information from Page 1 in {exam_language}, "
                    f"and sequentialize all questions across all pages into standard structured layout blocks. "
                    f"Ensure you preserve original {exam_language} spellings, characters, and formatting. "
                    f"Translate whatever visual format you see into our 'blocks' array: "
                    f"- Simple paragraphs or headings go to 'text_paragraph'. "
                    f"- Continuous sub-questions or standard lists of sentences go to 'list_block'. "
                    f"- Grid structures or blank writing cells go to 'grid_table_block'. "
                    f"- Side-by-side MCQs or split matching items (2, 3, or 4 columns wide) go to 'column_layout_block'. "
                    f"- Empty blank areas where students must draw something go to 'drawing_box_block'."
                )
                
                prompt = (
                    f"Analyze all these pages in order. Deconstruct the layout into a sequential list of blocks. "
                    f"The text extracted must be completely faithful to the {exam_language} words and letters present in the images."
                )
                
                contents = [prompt] + img_list
                
                try:
                    # Request structure from Gemini
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
                    exam_data = UniversalExamPaper(**raw_json)
                    
                    # Generate Word bytes with dynamic language support
                    word_bytes = create_docx(exam_data, exam_language)
                    
                    st.success(f"🎉 Word Document generated perfectly in {exam_language}!")
                    st.download_button(
                        label="📥 Download Microsoft Word File",
                        data=word_bytes,
                        file_name=f"Formatted_{exam_language}_Exam.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"Something went wrong: {e}")
