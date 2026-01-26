import streamlit as st
from docx import Document
import copy
import tempfile

TEMPLATE_PATH = "template.docx"   # auto load from repo


def replace_placeholders(doc, b1, b2):
    b2_text = f"[{b2}]"

    def replace_para(p):
        for run in p.runs:
            run.text = run.text.replace("{{B1}}", str(b1))
            run.text = run.text.replace("{{B2}}", b2_text)

    for p in doc.paragraphs:
        replace_para(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_para(p)


def append_template(master, b1, b2):
    temp = Document(TEMPLATE_PATH)
    replace_placeholders(temp, b1, b2)

    for el in temp.element.body:
        master.element.body.append(copy.deepcopy(el))


def build_file(batches):
    master = Document()
    master.element.body.clear()

    for batch in batches:
        b1 = batch["b1"]
        start = batch["start_b2"]
        pages = batch["pages"]

        for i in range(pages):
            b2 = start + (i // 2)
            append_template(master, b1, b2)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    master.save(tmp.name)
    return tmp.name


# ---------------- UI ----------------

st.title("📄 Batch Label Generator")

num_batches = st.number_input("Number of batches", 1, 10, 1)

batches = []

for i in range(num_batches):
    st.subheader(f"Batch {i+1}")
    c1, c2, c3 = st.columns(3)

    b1 = c1.text_input("B1", key=f"b1{i}")
    start_b2 = c2.number_input("Start B2", value=1, key=f"s{i}")
    pages = c3.number_input("Pages", value=10, key=f"p{i}")

    batches.append({
        "b1": b1,
        "start_b2": start_b2,
        "pages": pages
    })


if st.button("Generate File"):
    file_path = build_file(batches)

    with open(file_path, "rb") as f:
        st.download_button("⬇ Download DOCX", f, "labels_output.docx")
