import streamlit as st

# 🔐 Safety: always show UI even if docx fails
st.set_page_config(page_title="Batch Slip Generator")
st.title("Batch Slip Generator")

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception as e:
    st.error("python-docx failed to load. Check runtime.txt and requirements.txt")
    st.exception(e)
    st.stop()

import tempfile
import os

# ---------------- LANDSCAPE ----------------
def set_landscape(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

# ---------------- BORDER ----------------
def set_table_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')

    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '16')
        el.set(qn('w:color'), '000000')
        borders.append(el)

    tblPr.append(borders)

# ---------------- SLIP ----------------
def create_slip(doc, doc_type, batch_id, number):
    section = doc.sections[0]

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False

    table.columns[0].width = section.page_width - section.left_margin - section.right_margin
    table.rows[0].height = section.page_height - section.top_margin - section.bottom_margin

    set_table_border(table)
    cell = table.cell(0, 0)
    cell.paragraphs[0].clear()

    def add(text, size=26, bold=False):
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold

    if doc_type == "MOD":
        add("F074025-000000", 26)
        add("GUAR GUM POWDER", 30)
        add("MODIFIED", 30)
        add("")
    else:
        add("FARINA GUAR 200 MESH 5000 T/C", 32)
        add("")

    add("NET WEIGHT: 900 KG", 26)
    add("GROSS WEIGHT: 903 KG", 26)
    add("(Without Pallet)", 22)
    add("")
    add(f"BATCH NO.: {batch_id} ({number})", 34, True)

# ---------------- UI ----------------
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

if st.button("Generate Word File"):
    doc = Document()
    set_landscape(doc)

    first = True
    for bid, s, e in batches:
        for n in range(s, e + 1):
            for _ in range(2):  # 🔥 TWO PAGES PER NUMBER
                if not first:
                    doc.add_page_break()
                create_slip(doc, doc_type, bid, n)
                first = False

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download batch_slips.docx",
                f,
                file_name="batch_slips.docx"
            )
    os.remove(tmp.name)
