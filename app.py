import streamlit as st
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from extract_tables import extract_tables

st.set_page_config(page_title="PDF Table Extractor", page_icon="📄")

st.title("📄 PDF Table Extractor")
st.write("Upload a Mengenkalkulation PDF and download all 5 tables as a single Excel file.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file:
    st.success(f"Uploaded: {uploaded_file.name}")
    filename_stem = os.path.splitext(uploaded_file.name)[0]

    if st.button("Extract Tables"):
        with st.spinner("Extracting tables..."):
            tmp_pdf = "/tmp/input.pdf"
            with open(tmp_pdf, "wb") as f:
                f.write(uploaded_file.read())

            tables = extract_tables(tmp_pdf, "/tmp")

            # build xlsx in memory
            wb = Workbook()
            wb.remove(wb.active)  # remove default sheet

            header_font = Font(bold=True, name="Arial")
            header_fill = PatternFill("solid", start_color="D9E1F2")

            for sheet_name, rows in tables.items():
                ws = wb.create_sheet(title=sheet_name[:31])
                for r_idx, row in enumerate(rows, start=1):
                    # prepend filename to every data row (not header)
                    if r_idx == 1:
                        full_row = ["Source File"] + row
                    else:
                        full_row = [filename_stem] + row

                    for c_idx, value in enumerate(full_row, start=1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=value)
                        if r_idx == 1:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center")

                # auto-width columns
                for col in ws.columns:
                    max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

            xlsx_buffer = io.BytesIO()
            wb.save(xlsx_buffer)
            xlsx_buffer.seek(0)

        st.success("Done! All 5 tables extracted.")
        st.download_button(
            label="⬇️ Download Excel file",
            data=xlsx_buffer,
            file_name=f"{filename_stem}_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
