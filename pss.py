import streamlit as st
from docx import Document
from docxcompose.composer import Composer
import tempfile
import os


# -----------------------------
# Replace placeholders in doc
# -----------------------------
def replace_text(doc, b1, b2):
    for p in doc.paragraphs:
        if "{{B1}}" in p.text or "{{B2}}" in p.text:
            for run in p.runs:
                run.text = run.text.replace("{{B1}}", str(b1))
                run.text = run.text.replace("{{B2}}", str(b2))

    # also replace inside tables if any
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.text = run.text.replace("{{B1}}", str(b1))
                        run.text = run.text.replace("{{B2}}", str(b2))


# -----------------------------
# Build final doc from template
# -----------------------------
def build_file(template_path, batches):
    master = Document(template_path)
    composer = Composer(master)

    first = True

    for batch in batches:
        b1 = batch["b1"]
        start_b2 = batch["start_b2"]
        pages = batch["pages"]

        for i in range(pages):
            b2_val = start_b2 + (i // 2)

            doc = Document(template_path)
            replace_text(doc, b1, b2_val)

            if first:
                master = doc
                composer = Composer(master)
                first = False
            else:
                composer.append(doc)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    composer.save(tmp.name)
    return tmp.name


# =============================
# UI
# =============================
st.title("📄 Batch Label Generator (Template Safe)")

st.success("Your template formatting will remain EXACT — nothing shifts 👍")

uploaded_template = st.file_uploader(
    "Upload your template .docx (with {{B1}} and {{B2}})",
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
                "⬇ Download Final DOCX",
                f,
                file_name="labels_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
