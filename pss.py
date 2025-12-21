import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import tempfile
import os

# ---------- Templates ----------
FAR_TEMPLATE = """
FARINA GUAR 200 MESH 5000 T/C

NET WEIGHT: 900KG
GROSS WEIGHT: 903KG
(Without Pallet)

BATCH NO.: {{B2}} {{B22}}
"""

MOD_TEMPLATE = """
F074025-000000

GUAR GUM POWDER
MODIFIED

NET WEIGHT: 900 KG
GROSS WEIGHT: 903 KG
(Without Pallet)

BATCH NO.: {{B1}} {{B11}}
"""

# ---------- Helper Function ----------
def add_two_pages(doc, text):
    for _ in range(2):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(12)
        doc.add_page_break()

# ---------- Streamlit UI ----------
st.title("Batch Document Generator")

doc_type = st.selectbox("Select Document Type", ["FAR", "MOD"])
num_batches = st.number_input("Number of batches", min_value=1, step=1)

batch_data = []

st.subheader("Batch Details")

for i in range(num_batches):
    st.markdown(f"### Batch {i+1}")
    batch_id = st.text_input(
        f"Batch ID ({'B2' if doc_type == 'FAR' else 'B1'})",
        key=f"batch_id_{i}"
    )

    col1, col2 = st.columns(2)
    with col1:
        start_num = st.number_input(
            "From",
            min_value=1,
            step=1,
            key=f"start_{i}"
        )
    with col2:
        end_num = st.number_input(
            "To",
            min_value=start_num,
            step=1,
            key=f"end_{i}"
        )

    batch_data.append((batch_id, start_num, end_num))

# ---------- Generate Document ----------
if st.button("Generate Word Document"):
    doc = Document()

    for batch_id, start, end in batch_data:
        for attached_number in range(start, end + 1):

            if doc_type == "FAR":
                content = FAR_TEMPLATE.replace("{{B2}}", batch_id)\
                                      .replace("{{B22}}", str(attached_number))
            else:
                content = MOD_TEMPLATE.replace("{{B1}}", batch_id)\
                                      .replace("{{B11}}", str(attached_number))

            add_two_pages(doc, content)

    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        file_path = tmp.name

    with open(file_path, "rb") as f:
        st.download_button(
            label="Download Word Document",
            data=f,
            file_name="batch_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    os.remove(file_path)
