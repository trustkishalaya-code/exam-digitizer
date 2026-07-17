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
    page_title="AI Exam Paper Digitizer",
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
    student_info_line: str = Field(description="Details like: নাম _________ বিভাগ ______ রোল নং _________")
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

def create_docx(data: UniversalExamPaper):
    doc = Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    style = doc.styles['Normal']
    style.font.name = 'Kalpurush'
    style.font.size = Pt(11)
    
    # Header
    p_school = doc.add_paragraph()
    p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_school.add_run(data.school_name).bold = True
    p_school.runs[0].font.size = Pt(15)
    p_school.paragraph_format.space_after = Pt(2)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run(data.exam_title).bold = True
    p_title.runs[0].font.size = Pt(12)
    p_title.paragraph_format.space_after = Pt(8)
    
    # Metadata Grid
    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.autofit = False
    meta_table.columns[0].width = Inches(3.5)
    meta_table.columns[1].width = Inches(3.5)
    
    meta_table.rows[0].cells[0].paragraphs[0].text = f"শ্রেণী — {data.class_name}"
    meta_table.rows[0].cells[1].paragraphs[0].text = f"পূর্ণমান — {data.full_marks}"
    meta_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    meta_table.rows[1].cells[0].paragraphs[0].text = f"বিষয় — {data.subject}"
    meta_table.rows[1].cells[1].paragraphs[0].text = f"সময় — {data.time}"
    meta_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for row in meta_table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(11)
                    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    # Student Details
    p_info = doc.add_paragraph()
    p_info.add_run(data.student_info_line).font.size = Pt(11)
    p_info.paragraph_format.space_after = Pt(8)
    
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("—" * 60)
    p_div.paragraph_format.space_after = Pt(12)
    
    # Render Blocks
    for b in data.blocks:
        if b.block_type == 'text_paragraph':
            if b.text_content:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(b.text_content)
                if "।" in b.text_content or ":" in b.text_content or b.text_content.strip().startswith(('১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯', '০')):
                    run.bold = True
                    
        elif b.block_type == 'list_block':
            if b.text_content:
                p = doc.add_paragraph()
                p.add_run(b.text_content).bold = True
                p.paragraph_format.space_after = Pt(4)
            if b.list_items:
                for item in b.list_items:
                    lp = doc.add_paragraph()
                    lp.paragraph_format.left_indent = Inches(0.4)
                    lp.paragraph_format.space_after = Pt(4)
                    lp.add_run(item)
                    
        elif b.block_type == 'grid_table_block':
            if b.text_content:
                p = doc.add_paragraph()
                p.add_run(b.text_content).bold = True
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
                                    tbl.rows[r_idx].cells[c_idx].paragraphs[0].text = val
                doc.add_paragraph().paragraph_format.space_after = Pt(8)
                
        elif b.block_type == 'column_layout_block':
            if b.text_content:
                p = doc.add_paragraph()
                p.add_run(b.text_content).bold = True
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
                doc.add_paragraph().paragraph_format.space_after = Pt(8)
                
        elif b.block_type == 'drawing_box_block':
            if b.text_content:
                p = doc.add_paragraph()
                p.add_run(b.text_content).bold = True
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
st.title("📝 Universal Bengali Exam Digitizer")
st.write("Convert any handwritten exam question papers into a perfectly formatted `.docx` file!")

# Sidebar for Settings
st.sidebar.header("🔑 Credentials")
api_key = st.sidebar.text_input("Gemini API Key", type="password", help="Get your free key at aistudio.google.com")

st.info("💡 **Pro-Tip:** Make sure your files are sorted chronologically (e.g., `page1.jpg`, `page2.jpg`) before uploading so the layout remains in order!")

uploaded_files = st.file_uploader(
    "Upload Exam Images (Choose multiple if needed)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"📂 Loaded {len(uploaded_files)} pages successfully!")
    
    if st.button("🚀 Process & Generate Word Document"):
        if not api_key:
            st.error("Please enter your Gemini API Key in the sidebar!")
        else:
            with st.spinner("🤖 Gemini is parsing images, matching formatting and generating layouts..."):
                # Sort uploaded files by name
                sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                
                img_list = []
                for f in sorted_files:
                    img = Image.open(f)
                    img_list.append(img)
                
                client = genai.Client(api_key=api_key)
                
                system_instruction = (
                    "You are an expert layout-agnostic document OCR parser. "
                    "Your task is to scan the uploaded exam papers, read the header information from Page 1, "
                    "and sequentialize all questions across all pages into standard structured layout blocks. "
                    "Translate whatever visual format you see into our 'blocks' array: "
                    "- Simple paragraphs or headings go to 'text_paragraph'. "
                    "- Continuous sub-questions or standard lists of sentences go to 'list_block'. "
                    "- Grid structures or blank writing cells go to 'grid_table_block'. "
                    "- Side-by-side MCQs or split matching items (2, 3, or 4 columns wide) go to 'column_layout_block'. "
                    "- Empty blank areas where students must draw something go to 'drawing_box_block'."
                )
                
                prompt = (
                    "Analyze all these pages in order. Deconstruct the layout into a sequential list of blocks. "
                    "Be extremely faithful to the original wording and spellings in the images."
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
                    
                    # Generate Word bytes
                    word_bytes = create_docx(exam_data)
                    
                    st.success("🎉 Word Document generated perfectly!")
                    st.download_button(
                        label="📥 Download Microsoft Word File",
                        data=word_bytes,
                        file_name="Formatted_Bengali_Exam.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"Something went wrong: {e}")
