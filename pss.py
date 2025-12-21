import streamlit as st
from docx import Document
from copy import deepcopy
import tempfile
import os

# -------------------------------------------------
# Replace placeholders (handles split runs)
# -------------------------------------------------
def replace_text(doc, replacements):
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        for k, v in replacements.items():
            full_text = full_text.replace(k, v)

        if p.runs:
            p.runs[0].text = full_text
            for r in p.runs[1:]:
                r.text = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text(cell, replacements)

# -------------------------------------------------
# Copy template content safely
# -------------------------------------------------
def copy_body(src, dst):
    for el in src.element.body:
        dst.element.body.append(deepcopy(el))

# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------
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

# -------------------------------------------------
# Generate Document
# -------------------------------------------------
if st.button("Generate Word File"):

    final_doc = Document()

    # 🔥 remove default empty paragraph (CRITICAL)
    final_doc.element.body.clear()

    template_file = "far_template.docx" if doc_type == "FAR" else "mod_template.docx"
    first_page = True

    for batch_id, start, end in batches:
        for num in range(start, end + 1):

            temp_doc = Document(template_file)

            if doc_type == "FAR":
                replace_text(
                    temp_doc,
                    {
                        "{{B2}}": batch_id,
                        "{{B22}}": f"({num})"
                    }
                )
            else:
                replace_text(
                    temp_doc,
                    {
                        "{{B1}}": batch_id,
                        "{{B11}}": f"({num})"
                    }
                )

            # 2 pages per number
            for _ in range(2):
                if not first_page:
                    final_doc.add_page_break()
                copy_body(temp_doc, final_doc)
                first_page = False

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        final_doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download batch_slips.docx",
                data=f,
                file_name="batch_slips.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    os.remove(tmp.name)
