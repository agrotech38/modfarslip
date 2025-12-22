import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os

# ---------------- PAGE SETUP ----------------
def setup_landscape(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

# ---------------- BORDER ----------------
def set_table_border(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '18')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    tblPr.append(borders)

# ---------------- SLIP PAGE ----------------
def create_slip(doc, doc_type, batch_id, num):
    section = doc.sections[0]

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False

    table.columns[0].width = section.page_width - section.left_margin - section.right_margin
    table.rows[0].height = section.page_height - section.top_margin - section.bottom_margin

    set_table_border(table)
    cell = table.cell(0, 0)
    cell.paragraphs[0].clear()

    def line(text, size, bold=False):
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold

    # BIGGER TEXT
    if doc_type == "FAR":
        line("FARINA GUAR 200 MESH 5000 T/C", 34, True)
        line("")
    else:
        line("F074025-000000", 30)
        line("GUAR GUM POWDER", 34, True)
        line("MODIFIED", 32)
        line("")

    line("NET WEIGHT: 900 KG", 30)
    line("GROSS WEIGHT: 903 KG", 30)
    line("(Without Pallet)", 26)
    line("")
    line(f"BATCH NO.: {batch_id} ({num})", 38, True)

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Batch Slip Generator")
st.title("Batch Slip Generator")

doc_type = st.selectbox("Select Type", ["FAR", "MOD"])
batch_count = st.number_input("Number of Batches", min_value=1, step=1)

batches = []
for i in range(batch_count):
    st.subheader(f"Batch {i+1}")
    bid = st.text_input("Batch ID", key=f"id{i}")
    c1, c2 = st.columns(2)
    with c1:
        s = st.number_input("From", min_value=1, key=f"s{i}")
    with c2:
        e = st.number_input("To", min_value=s, key=f"e{i}")
    batches.append((bid, s, e))

# ---------------- GENERATE ----------------
if st.button("Generate Word File"):
    doc = Document()
    setup_landscape(doc)

    first_page = True

    for bid, s, e in batches:
        for num in range(s, e + 1):
            for _ in range(2):  # 🔥 EXACTLY TWO PAGES
                if not first_page:
                    doc.add_page_break()  # ONLY before new page
                create_slip(doc, doc_type, bid, num)
                first_page = False

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download batch_slips.docx",
                f,
                file_name="batch_slips.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    os.remove(tmp.name)
