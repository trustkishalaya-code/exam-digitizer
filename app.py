import io
import json
from typing import List, Optional
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from google import genai
from google.genai import types
from PIL import Image, ImageEnhance
import streamlit as st

# --- Page Setup ---
st.set_page_config(
    page_title="Academic Studio // Brutalism Edition",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- ⚡ STYLE 3: HIGH-CONTRAST ELECTRIC YELLOW BRUTALISM SYSTEM ---
st.markdown(
    """
<style>
    /* Global Reset & Base */
    .stApp {
        background-color: #000000 !important;
        color: #FFFFFF !important;
        font-family: 'Courier New', Courier, monospace !important;
    }
    
    /* Hide default Streamlit elements for a clean custom look */
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

    /* Bold Sidebar */
    [data-testid="stSidebar"] {
        background-color: #000000 !important;
        border-right: 3px solid #FFEA00 !important;
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

    /* Action Buttons */
    .stButton>button {
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
        box-shadow: 4px 4px 0px #FFFFFF;
        transition: transform 0.1s ease;
    }
    
    .stButton>button:hover {
        transform: translate(-2px, -2px);
        box-shadow: 6px 6px 0px #FFFFFF;
    }
    
    /* Success / Download Button */
    .stDownloadButton>button {
        display: block;
        width: 100%;
        padding: 12px;
        border: 3px solid #FFFFFF !important;
        background: #FFFFFF !important;
        color: #000000 !important;
        font-family: 'Courier New', Courier, monospace !important;
        font-weight: 900 !important;
        font-size: 10pt;
        text-transform: uppercase;
        border-radius: 0px !important;
        box-shadow: 4px 4px 0px #FFEA00;
        transition: transform 0.1s ease;
    }
    
    .stDownloadButton>button:hover {
        transform: translate(-2px, -2px);
        box-shadow: 6px 6px 0px #FFEA00;
    }

    /* Chunky Upload Zone */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #FFEA00 !important;
        padding: 25px;
        text-align: center;
        background: #000000 !important;
        border-radius: 0px !important;
    }
    
    p, li { font-family: 'Courier New', Courier, monospace !important; color: #FFFFFF; }
</style>
""",
    unsafe_allow_html=True,
)


# --- 1. Universal Structural Schemas ---
class LayoutBlock(BaseModel):
  block_type: str = Field(
      description=(
          "Can be: 'text_paragraph', 'list_block', 'grid_table_block',"
          " 'column_layout_block', 'drawing_box_block'"
      )
  )
  text_content: Optional[str] = Field(default=None)
  list_items: Optional[List[str]] = Field(default=None)
  table_rows: Optional[int] = Field(default=None)
  table_cols: Optional[int] = Field(default=None)
  table_data: Optional[List[List[str]]] = Field(
      description="2D array matrix for grid_table_block", default=None
  )
  columns_data: Optional[List[List[str]]] = Field(
      description="Column text lists for column_layout_block", default=None
  )
  box_height_inches: Optional[float] = Field(default=2.0)


class UniversalExamPaper(BaseModel):
  school_name: str
  exam_title: str
  class_name: str
  subject: str
  full_marks: str
  time: str
  student_info_line: str = Field(
      description="Student details line placeholder (Name, Roll, etc.)"
  )
  blocks: List[LayoutBlock] = Field(
      description="Chronological sequence of structured objects found"
  )


# --- 2. Advanced Typography & Word Engine ---
def optimize_image(img, max_width=2500):
  if img.width > max_width:
    ratio = max_width / img.width
    new_size = (max_width, int(img.height * ratio))
    img = img.resize(new_size, Image.Resampling.LANCZOS)
  return img.convert("RGB")


def set_table_borders(table, color="000000", sz="8"):
  tblPr = table._tbl.tblPr
  tblBorders = OxmlElement("w:tblBorders")
  for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
    border = OxmlElement(f"w:{border_name}")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), sz)
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), color)
    tblBorders.append(border)
  tblPr.append(tblBorders)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
  tcPr = cell._tc.get_or_add_tcPr()
  tcMar = OxmlElement("w:tcMar")
  for m_name, m_val in [
      ("top", top),
      ("bottom", bottom),
      ("left", left),
      ("right", right),
  ]:
    node = OxmlElement(f"w:{m_name}")
    node.set(qn("w:w"), str(m_val))
    node.set(qn("w:type"), "dxa")
    tcMar.append(node)
  tcPr.append(tcMar)


def set_section_columns(section, num_cols):
  sectPr = section._sectPr
  cols = sectPr.xpath("./w:cols")
  if cols:
    cols[0].set(qn("w:num"), str(num_cols))
  else:
    new_cols = OxmlElement("w:cols")
    new_cols.set(qn("w:num"), str(num_cols))
    new_cols.set(qn("w:space"), "720")
    sectPr.append(new_cols)


def create_docx(
    data: UniversalExamPaper, language: str, grade_tier: str, base_font_size: int
):
  doc = Document()
  section = doc.sections[0]

  is_early_childhood = grade_tier in ["Nursery", "PP / LKG / UKG"]

  if is_early_childhood:
    headline_fs = 18
    question_fs = 22
    header_title_fs = 24
    meta_fs = 18

    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.orientation = WD_ORIENT.PORTRAIT
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
  else:
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

  font_mapping = {
      "Bengali": "Kalpurush",
      "Hindi": "Mangal",
      "English": "Courier New",
  }
  selected_font = font_mapping.get(language, "Courier New")

  style = doc.styles["Normal"]
  style.font.name = selected_font
  style.font.size = Pt(question_fs)

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

  class_label = (
      "Class"
      if language == "English"
      else ("শ্রেণী" if language == "Bengali" else "कक्षा")
  )
  marks_label = (
      "Full Marks"
      if language == "English"
      else ("পূর্ণমান" if language == "Bengali" else "पूर्णांक")
  )
  subject_label = (
      "Subject"
      if language == "English"
      else ("বিষয়" if language == "Bengali" else "विषय")
  )
  time_label = (
      "Time"
      if language == "English"
      else ("সময়" if language == "Bengali" else "समय")
  )

  def format_meta(label, val):
    if not val:
      return ""
    if (
        label.lower() in val.lower()
        or "শ্রেণী" in val
        or "कक्षा" in val
        or "পূর্ণমান" in val
        or "বিষয়" in val
        or "সময়" in val
        or "पूर्णांक" in val
    ):
      return val
    return f"{label} — {val}"

  meta_table = doc.add_table(rows=2, cols=2)
  meta_table.autofit = False
  col_w = Inches(3.5) if is_early_childhood else Inches(4.5)
  meta_table.columns[0].width = col_w
  meta_table.columns[1].width = col_w

  meta_table.rows[0].cells[0].paragraphs[0].text = format_meta(
      class_label, data.class_name
  )
  meta_table.rows[0].cells[1].paragraphs[0].text = format_meta(
      marks_label, data.full_marks
  )
  meta_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

  meta_table.rows[1].cells[0].paragraphs[0].text = format_meta(
      subject_label, data.subject
  )
  meta_table.rows[1].cells[1].paragraphs[0].text = format_meta(
      time_label, data.time
  )
  meta_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

  for row in meta_table.rows:
    for cell in row.cells:
      for p in cell.paragraphs:
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
          run.bold = True
          run.font.size = Pt(meta_fs)

  doc.add_paragraph().paragraph_format.space_after = Pt(4)

  p_info = doc.add_paragraph()
  run_info = p_info.add_run(data.student_info_line)
  run_info.font.size = Pt(meta_fs)
  run_info.bold = True
  p_info.paragraph_format.space_after = Pt(6)

  p_div = doc.add_paragraph()
  p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_div.add_run("—" * 50 if not is_early_childhood else "—" * 60)
  p_div.paragraph_format.space_after = Pt(12)

  for b in data.blocks:
    space_after_val = 14 if is_early_childhood else 6

    if b.block_type == "text_paragraph" and b.text_content:
      p = doc.add_paragraph()
      p.paragraph_format.space_after = Pt(space_after_val)

      cleaned = b.text_content.strip()
      is_headline = cleaned.endswith((":", "।")) and not any(
          char.isdigit() for char in cleaned[:3]
      )

      run = p.add_run(b.text_content)

      if is_headline:
        run.font.size = Pt(headline_fs)
        run.bold = True
      else:
        run.font.size = Pt(question_fs)
        if cleaned.startswith((
            "১",
            "২",
            "৩",
            "৪",
            "৫",
            "৬",
            "৭",
            "৮",
            "৯",
            "০",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "0",
            "१",
            "२",
            "३",
            "४",
            "५",
            "६",
            "७",
            "८",
            "९",
            "०",
        )):
          run.bold = True

    elif b.block_type == "list_block":
      if b.text_content:
        lp_head = doc.add_paragraph()
        run_h = lp_head.add_run(b.text_content)
        run_h.bold = True
        run_h.font.size = Pt(headline_fs)
        lp_head.paragraph_format.space_after = Pt(6)
      if b.list_items:
        for item in b.list_items:
          lp = doc.add_paragraph()
          lp.paragraph_format.left_indent = Inches(0.3)
          lp.paragraph_format.space_after = Pt(8 if is_early_childhood else 3)
          run_item = lp.add_run(item)
          run_item.font.size = Pt(question_fs)

    elif b.block_type == "grid_table_block":
      if b.text_content:
        tp = doc.add_paragraph()
        run_t = tp.add_run(b.text_content)
        run_t.bold = True
        run_t.font.size = Pt(headline_fs)
      if b.table_rows and b.table_cols and b.table_data:
        tbl = doc.add_table(rows=b.table_rows, cols=b.table_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, color="000000", sz="8")
        for r_idx, row_items in enumerate(b.table_data):
          if r_idx < b.table_rows:
            for c_idx, val in enumerate(row_items):
              if c_idx < b.table_cols:
                cell = tbl.rows[r_idx].cells[c_idx]
                cell.paragraphs[0].text = val
                if is_early_childhood:
                  set_cell_margins(
                      cell, top=220, bottom=220, left=200, right=200
                  )
                for r in cell.paragraphs[0].runs:
                  r.font.size = Pt(question_fs)
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    elif b.block_type == "column_layout_block":
      if b.text_content:
        cp = doc.add_paragraph()
        run_c = cp.add_run(b.text_content)
        run_c.bold = True
        run_c.font.size = Pt(headline_fs)
      if b.columns_data:
        num_cols = len(b.columns_data)
        max_rows = max(len(col) for col in b.columns_data)
        col_tbl = doc.add_table(rows=max_rows, cols=num_cols)
        set_table_borders(col_tbl, color="000000", sz="8")
        col_width = (
            Inches(3.2 / num_cols)
            if is_early_childhood
            else Inches(4.0 / num_cols)
        )
        for c in range(num_cols):
          col_tbl.columns[c].width = col_width
          col_items = b.columns_data[c]
          for r in range(max_rows):
            if r < len(col_items):
              cell = col_tbl.rows[r].cells[c]
              cell.paragraphs[0].text = col_items[r]
              for rn in cell.paragraphs[0].runs:
                rn.font.size = Pt(question_fs)
        doc.add_paragraph()

    elif b.block_type == "drawing_box_block":
      if b.text_content:
        dp = doc.add_paragraph()
        run_d = dp.add_run(b.text_content)
        run_d.bold = True
        run_d.font.size = Pt(headline_fs)
      box_tbl = doc.add_table(rows=1, cols=1)
      box_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
      default_h = (b.box_height_inches or 2.0) * (
          1.5 if is_early_childhood else 1.0
      )
      box_tbl.rows[0].height = Inches(default_h)
      set_table_borders(box_tbl, color="000000", sz="8")
      doc.add_paragraph().paragraph_format.space_after = Pt(12)

  bio = io.BytesIO()
  doc.save(bio)
  return bio.getvalue()


# --- 3. UI Header & Sidebar Controls ---
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

with st.sidebar:
  st.markdown(
      '<div class="brut-box"><h3 style="font-size: 11pt; color: #FFEA00; margin-bottom: 14px; border-bottom: 2px solid #FFEA00; padding-bottom: 6px;">[ CONTROLS ]</h3>',
      unsafe_allow_html=True,
  )

  st.markdown('<label class="brut-label">API KEY</label>', unsafe_allow_html=True)
  api_key = st.text_input("API KEY", type="password", label_visibility="collapsed")
  if not api_key:
    st.warning("⚠️ Access Key required to launch.")

  st.markdown(
      '<label class="brut-label">TARGET GRADE TIER</label>',
      unsafe_allow_html=True,
  )
  grade_tier = st.selectbox(
      "TARGET GRADE TIER",
      ["Nursery", "PP / LKG / UKG", "Classes 1 to 4"],
      label_visibility="collapsed",
      help=(
          "Nursery/PP automatically configures 18pt headlines & 22pt questions"
          " with generous answer padding."
      ),
  )

  st.markdown(
      '<label class="brut-label">DOCUMENT LANGUAGE</label>',
      unsafe_allow_html=True,
  )
  exam_language = st.selectbox(
      "DOCUMENT LANGUAGE",
      ["Bengali", "English", "Hindi"],
      label_visibility="collapsed",
  )

  st.markdown(
      '<label class="brut-label">TYPOGRAPHY CONFIG</label>',
      unsafe_allow_html=True,
  )
  if grade_tier in ["Nursery", "PP / LKG / UKG"]:
    st.info(
        "📏 **Early Childhood Active:**\n- **Headlines:** `18pt` (Bold)\n-"
        " **Questions:** `22pt` (Spacious)\n- **Borders:** `#000000`"
    )
    custom_font_size = 22
  else:
    st.info(
        "📑 **Standard Exam Active:**\n- **Layout:** Landscape A4 (2"
        " Columns)\n- **Borders:** `#000000`"
    )
    custom_font_size = st.number_input(
        "BASE FONT SIZE (PT)",
        min_value=9,
        max_value=16,
        value=11,
        label_visibility="collapsed",
    )

  st.write("")
  if st.button("[ CLEAR WORKSPACE ]"):
    st.session_state.parsed_data = None
    st.session_state.original_filename = "Exam_Output.docx"
    st.rerun()

  st.markdown("</div>", unsafe_allow_html=True)

# --- Session State Initialization ---
if "parsed_data" not in st.session_state:
  st.session_state.parsed_data = None
if "original_filename" not in st.session_state:
  st.session_state.original_filename = "Exam_Output.docx"

# --- Main Workspace Layout ---
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
  st.markdown(
      '<div class="brut-box"><h3 style="font-size: 11pt; color: #FFEA00; margin-bottom: 14px; border-bottom: 2px solid #FFEA00; padding-bottom: 6px;">[ 01. INTAKE STREAM ]</h3>',
      unsafe_allow_html=True,
  )

  uploaded_files = st.file_uploader(
      f"Upload {exam_language} exam sheets",
      type=["jpg", "jpeg", "png"],
      accept_multiple_files=True,
  )

  if uploaded_files:
    st.session_state.original_filename = (
        uploaded_files[0].name.rsplit(".", 1)[0] + ".docx"
    )
    st.success(
        f"🔥 Loaded {len(uploaded_files)} source files ready for processing."
    )

    st.write("")
    if st.button("[ RUN HIGH-PRECISION EXTRACTION ]"):
      if not api_key:
        st.error("Missing API Key.")
      else:
        with st.status(
            "⚡ Reading & processing files...", expanded=True
        ) as status:
          try:
            sorted_files = sorted(uploaded_files, key=lambda x: x.name)
            img_list = []

            st.write("Preparing raw images...")
            for f in sorted_files:
              raw_img = Image.open(f)

              enhancer = ImageEnhance.Contrast(raw_img)
              img_cont = enhancer.enhance(1.7)
              enhancer2 = ImageEnhance.Sharpness(img_cont)
              enhanced_img = enhancer2.enhance(2.1)

              img_list.append(optimize_image(enhanced_img))

            client = genai.Client(api_key=api_key)

            system_instruction = (
                f"You are a meticulous, zero-error academic document"
                f" transcriber specializing in {exam_language} primary papers"
                f" for {grade_tier}. Your goal is absolute fidelity. Do not"
                " hallucinate or skip any words, numbers, punctuation marks,"
                " or fill-in lines (......). Carefully evaluate layout"
                " sequences item by item. Ensure all tabular data matrices are"
                " fully represented element by element. Translate icon or"
                " drawing illustrations into fitting contextual emojis (e.g."
                " ☀️, 🍎, 🌳)."
            )

            prompt = (
                f"Perform a comprehensive structural extraction of all"
                f" provided pages in sequence for {grade_tier}."
                " Double-check every character to ensure complete accuracy"
                " matching the visual source material."
            )
            contents = [prompt] + img_list

            st.write(
                f"Running high-fidelity OCR on {len(img_list)} total page(s)..."
            )

            fallback_models = ["gemini-3.5-flash", "gemini-3.1-flash-lite"]
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
                        temperature=0.0,
                    ),
                )
                break
              except Exception as e:
                error_msg = str(e)
                if any(
                    code in error_msg
                    for code in [
                        "503",
                        "UNAVAILABLE",
                        "429",
                        "404",
                        "NOT_FOUND",
                    ]
                ):
                  st.warning(
                      f"Model {model_name} unavailable or busy. Rerouting..."
                  )
                  last_error = error_msg
                  continue
                else:
                  raise e

            if not response:
              raise Exception(
                  "All backup servers are currently busy or unavailable. Please"
                  f" try again in a few minutes. (Last Error: {last_error})"
              )

            st.write("Parsing data...")
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
              raw_text = raw_text[7:-3].strip()
            elif raw_text.startswith("```"):
              raw_text = raw_text[3:-3].strip()

            raw_json = json.loads(raw_text)
            st.session_state.parsed_data = UniversalExamPaper(**raw_json)
            status.update(
                label="Extraction complete!", state="complete", expanded=False
            )
            st.rerun()

          except json.JSONDecodeError:
            status.update(label="Extraction failed", state="error")
            st.error(
                "Failed to parse API output properly. Please try clicking"
                " extract again."
            )
          except Exception as e:
            status.update(label="Extraction failed", state="error")
            st.error(f"Error: {e}")

  st.markdown("</div>", unsafe_allow_html=True)

with col_right:
  st.markdown(
      '<div class="brut-box"><h3 style="font-size: 11pt; color: #FFEA00; margin-bottom: 14px; border-bottom: 2px solid #FFEA00; padding-bottom: 6px;">[ 02. STUDIO REVIEW & EXPORT ]</h3>',
      unsafe_allow_html=True,
  )

  if st.session_state.parsed_data is None:
    st.info(
        "Upload source files and run extraction to preview and export"
        " documents here."
    )
  else:
    data = st.session_state.parsed_data

    with st.form("exam_editor_form"):
      st.markdown(
          '<label class="brut-label">DOCUMENT METADATA</label>',
          unsafe_allow_html=True,
      )
      data.school_name = st.text_input("School Name", value=data.school_name)
      data.exam_title = st.text_input("Exam Title", value=data.exam_title)

      c1, c2 = st.columns(2)
      with c1:
        data.class_name = st.text_input("Class", value=data.class_name)
        data.subject = st.text_input("Subject", value=data.subject)
      with c2:
        data.full_marks = st.text_input("Full Marks", value=data.full_marks)
        data.time = st.text_input("Time", value=data.time)

      data.student_info_line = st.text_input(
          "Student Info Placeholder", value=data.student_info_line
      )

      update_submitted = st.form_submit_button("💾 SAVE METADATA EDITS")

      if update_submitted:
        st.success("Metadata updated successfully! Ready for download.")

    st.divider()
    st.markdown(
        '<label class="brut-label">DOWNLOAD OUTPUT DOCUMENT</label>',
        unsafe_allow_html=True,
    )

    word_bytes = create_docx(
        st.session_state.parsed_data,
        exam_language,
        grade_tier,
        int(custom_font_size),
    )

    st.download_button(
        label=f"⬇️ DOWNLOAD {st.session_state.original_filename.upper()}",
        data=word_bytes,
        file_name=st.session_state.original_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )

  st.markdown("</div>", unsafe_allow_html=True)
