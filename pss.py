import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os

def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')

    for edge in ('top', 'left', 'bottom', 'right'):
        elem = OxmlElement(f'w:{edge}')
        elem.set(qn('w:val'), 'single')
        elem.set(qn('w:sz'), '12')
        elem.set(qn('w:space'), '0')
        elem.set(qn('w:color'), '000000')
        borders.append(elem)

    tblPr.append(borders)

def create_slip(doc, doc_type, batch_id, number):
    table = doc.add_table(rows=1, cols=1)
    set_table_border(table)

    cell = table.cell(0, 0)
    cell.paragraphs[0].clear()

    def add_line(text, bold=False, size=12):
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)

    if doc_type == "FAR":
        add_line("FARINA GUAR 200 MESH 5000 T/C", size=14)
    else:
        add_line("F074025-000000")
        add_line("GUAR GUM POWDER")
        add_line("MODIFIED")

    add_line("")
    add_line("NET WEIGHT: 900 KG")
    add_line("GROSS WEIGHT: 903 KG")
    add_line("(Without Pallet)")
    add_line("")
    add_line(f"BATCH NO.: {batch_id} ({number})", bold=True, size=14)

st.set_page_config(page_title="Batch Slip Generator")
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

    first = True
    for batch_id, start, end in batches:
        for num in range(start, end + 1):
            if not first:
                doc.add_page_break()
            create_slip(doc, doc_type, batch_id, num)
            first = False

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download batch_slips.docx",
                data=f,
                file_name="batch_slips.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    os.remove(tmp.name)
