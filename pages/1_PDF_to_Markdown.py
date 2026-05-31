import io
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="PDF to Markdown",
    page_icon="📄",
    layout="centered",
)

st.title("PDF to Markdown")
st.markdown("Upload a PDF and get the extracted text as Markdown.")

uploaded_file = st.file_uploader("Upload a PDF", type=["pdf"])

if not uploaded_file:
    st.stop()

with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
    pages_text = []
    for i, page in enumerate(pdf.pages, start=1):
        text = page.extract_text() or ""
        pages_text.append(f"## Page {i}\n\n{text}")

markdown_output = "\n\n---\n\n".join(pages_text)

st.divider()
st.subheader("Extracted Text")
st.text_area("Markdown output", value=markdown_output, height=400, label_visibility="collapsed")

st.download_button(
    label="Download as .md",
    data=markdown_output.encode("utf-8"),
    file_name=uploaded_file.name.replace(".pdf", ".md"),
    mime="text/markdown",
)
