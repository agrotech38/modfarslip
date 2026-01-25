import streamlit as st
from docx import Document
import copy
import tempfile
import os


# -----------------------------
# Replace placeholders
# -----------------------------
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


# -----------------------------
# TRUE page cloning (NO shift)
# -----------------------------
def append_template_page(master, template_path, b1, b2):
    temp_doc = Document(template_path)
    replace_placeholders(temp_doc, b1, b2)

    for element in temp_doc.element.body:
        master.element.body.append(copy.deepcopy(element))


def build_file(template_path, batches):
    master = Document()

    # remove blank first paragraph
    master.element.body.clear()

    for batch in batches:
        b1 = batch["b1"]
        start_b2 = batch["start_b2"]
        pages = batch["pages"]

        for i in range(pages):
            b2_val = start_b2 + (i // 2)
            append_template_page(master, template_path, b1, b2_val)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    master.save(tmp.name)
    return tmp.name


# =============================
# UI
# =============================
st.title("📄 Batch Label Generator (Perfect Alignment)")

uploaded_template = st.file_uploader(
    "Upload template (.docx with {{B1}} {{B2}})",
    type=["docx"]
)

if uploaded_template:
    template_path = os.path.join(tempfile.gettempdir(), "template.docx")
    with open(template_path, "wb") as f:
        f.write(uploaded_template.read())

    num_batches = st.number_input("Number of batches", min_value=1, value=1)

    batches = []

    for i in range(num_batches):
        st.subheader(f"Batch {i+1}")
        c1, c2, c3 = st.columns(3)

        with c1:
            b1 = st.text_input("B1", key=f"b1{i}")

        with c2:
            start_b2 = st.number_input("Start B2", value=1, key=f"s{i}")

        with c3:
            pages = st.number_input("Pages", value=10, key=f"p{i}")

        batches.append({
            "b1": b1,
            "start_b2": start_b2,
            "pages": pages
        })

    if st.button("Generate File"):
        out_file = build_file(template_path, batches)

        with open(out_file, "rb") as f:
            st.download_button(
                "⬇ Download DOCX",
                f,
                file_name="labels_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
