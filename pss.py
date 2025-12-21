import streamlit as st
from docx import Document
import tempfile
import os

# ------------------ Helper: Replace placeholders ------------------
def replace_placeholders(doc, replacements):
    # paragraphs
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, value)

    # tables / boxes
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, replacements)

# ------------------ Streamlit UI ------------------
st.set_page_config(page_title="Batch Slip Generator", layout="centered")
st.title("Batch Slip Generator")

doc_type = st.selectbox("Select Type", ["FAR", "MOD"])
num_batches = st.number_input("Number of Batches", min_value=1, step=1)

batches = []

for i in range(num_batches):
    st.subheader(f"Batch {i + 1}")

    batch_id = st.text_input(
        "Batch ID",
        key=f"batch_id_{i}",
        placeholder="e.g. B/25/10001"
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

    batches.append((batch_id, start_num, end_num))

# ------------------ Generate Document ------------------
if st.button("Generate Word Document"):

    final_doc = Document()

    template_path = "far_template.docx" if doc_type == "FAR" else "mod_template.docx"

    for batch_id, start, end in batches:
        for num in range(start, end + 1):

            temp_doc = Document(template_path)

            if doc_type == "FAR":
                replace_placeholders(
                    temp_doc,
                    {
                        "{{B2}}": batch_id,
                        "{{B22}}": str(num)
                    }
                )
            else:
                replace_placeholders(
                    temp_doc,
                    {
                        "{{B1}}": batch_id,
                        "{{B11}}": str(num)
                    }
                )

            # Each number = 2 pages
            for _ in range(2):
                for element in temp_doc.element.body:
                    final_doc.element.body.append(element)
                final_doc.add_page_break()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        final_doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                label="Download Word File",
                data=f,
                file_name="batch_slips.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    os.remove(tmp.name)
