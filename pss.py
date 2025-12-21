import streamlit as st
from docx import Document
from copy import deepcopy
import tempfile
import os

# -------------------------------------------------
# Replace placeholders safely (split runs handled)
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
# Append full document as-is (NO layout break)
# -------------------------------------------------
def append_document(src, dst):
    for element in src.element.body:
        dst.element.body.append(deepcopy(element))

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
# Generate Document (CORRECT WAY)
# -------------------------------------------------
if st.button("Generate Word File"):

    final_doc = Document()
    final_doc.element.body.clear()

    template_file = "far_template.docx" if doc_type == "FAR" else "mod_template.docx"
    first = True

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

            # Add page break ONLY between slips
            if not first:
                final_doc.add_page_break()

            append_document(temp_doc, final_doc)
            first = False

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
