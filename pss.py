import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os

def setup_landscape(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

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

def create_slip(doc, doc_type, batch_id, num):
    section = doc.sections[0]
    page_width = section.page_width - section.left_margin - section.right_margin
    page_height = section.page_height - section.top_margin - section.bottom_margin

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = page_width
    set_table_border(table)

    row = table.rows[0]
    row.height = page_height
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:trHeight')
    h.set(qn('w:val'), str(int(page_height)))
    h.set(qn('w:hRule'), 'exact')
    trPr.append(h)

    cell = row.cells[0]
    cell.paragraphs[0].clear()

    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side in ('top', 'left', 'bottom', 'right'):
        mar = OxmlElement(f'w:{side}')
        mar.set(qn('w:w'), '500')
        mar.set(qn('w:type'), 'dxa')
        tcMar.append(mar)
    tcPr.append(tcMar)

    def line(text, size=34, bold=False):
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold

    if doc_type == "FAR":
        line("FARINA GUAR 200 MESH 5000 T/C", 42, True)
        line("")
    else:
        line("F074025-000000", 36)
        line("GUAR GUM POWDER", 42, True)
        line("MODIFIED", 40)
        line("")

    line("NET WEIGHT: 900 KG", 36)
    line("GROSS WEIGHT: 903 KG", 36)
    line("(Without Pallet)", 30)
    line("")
    line(f"BATCH NO.: {batch_id} ({num})", 46, True)

st.set_page_config(page_title="Batch Slip Generator", layout="centered")
st.title("Batch Slip Generator")

doc_type = st.selectbox("Select Type", ["FAR", "MOD"])
batch_count = st.number_input("Number of Batches", min_value=1, step=1)

batches = []
for i in range(batch_count):
    st.subheader(f"Batch {i + 1}")
    batch_id = st.text_input("Batch ID", key=f"id{i}")
    col1, col2 = st.columns(2)
    with col1:
        start = st.number_input("From", min_value=1, key=f"s{i}")
    with col2:
        end = st.number_input("To", min_value=start, key=f"e{i}")
    batches.append((batch_id, start, end))

if st.button("Generate Word File"):
    doc = Document()
    setup_landscape(doc)
    first_page = True
    for batch_id, start, end in batches:
        for num in range(start, end + 1):
            for _ in range(2):
                if not first_page:
                    doc.add_page_break()
                create_slip(doc, doc_type, batch_id, num)
                first_page = False
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download batch_slips.docx",
                f,
                file_name="batch_slips.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    os.remove(tmp.name)
