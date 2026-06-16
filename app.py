import streamlit as st
import zipfile
import io
from extract_tables import extract_tables

st.set_page_config(page_title="PDF Table Extractor", page_icon="📄")

st.title("📄 PDF Table Extractor")
st.write("Upload a Mengenkalkulation PDF and download all 5 tables as CSV files.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file:
    st.success(f"Uploaded: {uploaded_file.name}")

    if st.button("Extract Tables"):
        with st.spinner("Extracting tables..."):
            # save uploaded PDF temporarily
            tmp_pdf = "/tmp/input.pdf"
            with open(tmp_pdf, "wb") as f:
                f.write(uploaded_file.read())

            # extract tables to /tmp
            tables = extract_tables(tmp_pdf, "/tmp")

            # bundle all CSVs into a zip
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for name in tables.keys():
                    zf.write(f"/tmp/{name}.csv", f"{name}.csv")
            zip_buffer.seek(0)

        st.success("Done! All 5 tables extracted.")

        # individual download buttons
        for name, rows in tables.items():
            csv_content = "\n".join([",".join(str(c) for c in row) for row in rows])
            st.download_button(
                label=f"⬇️ Download {name}.csv ({len(rows)-1} rows)",
                data=csv_content,
                file_name=f"{name}.csv",
                mime="text/csv"
            )

        # zip download
        st.download_button(
            label="⬇️ Download all as ZIP",
            data=zip_buffer,
            file_name="tables.zip",
            mime="application/zip"
        )
