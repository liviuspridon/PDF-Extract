import streamlit as st
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from extract_tables import extract_tables
from fill_fisa import fill_fisa

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "_FIsa_Prototip.xlsx")

st.set_page_config(page_title="PDF Table Extractor", page_icon="📄")
st.title("📄 Mengenkalkulation Extractor")

# ── Câmpuri antet ─────────────────────────────────────────────
st.subheader("Date proiect")

col1, col2 = st.columns(2)

with col1:
    client      = st.text_input("Client")
    titlu       = st.text_input("Titlu proiect")
    nr_proiect  = st.text_input("Nr. proiect")
    data        = st.date_input("Data")

with col2:
    tip_proiect   = st.selectbox("Tip proiect", ["", "Mobilier în Roomdesigner", "Debitare căntuire"])
    tip_solicitare = st.selectbox("Tip solicitare", ["", "Ofertă", "Comandă (cu work preparation)"])
    preluat_de    = st.selectbox("Preluat de", ["", "Dan Slotea", "David Constantin", "Iulian Necula",
                                                 "Liliana Chiriță", "Adrian Mărgărit", "Denisa Manea"])
    proiectat_de  = st.selectbox("Proiectat de", ["", "Client", "Plan M"])

observatii = st.text_area("Observații")

antet = {
    "B3": client,
    "B4": titlu,
    "B5": nr_proiect,
    "B6": str(data),
    "I3": tip_proiect,
    "I4": tip_solicitare,
    "I5": preluat_de,
    "I6": proiectat_de,
    "B50": observatii,
}

st.divider()

# ── Upload PDF ────────────────────────────────────────────────
uploaded_file = st.file_uploader("Încarcă PDF-ul", type="pdf")

if uploaded_file:
    st.success(f"Fișier încărcat: {uploaded_file.name}")
    filename_stem = os.path.splitext(uploaded_file.name)[0]

    if st.button("Extrage"):
        with st.spinner("Se extrag datele..."):
            tmp_pdf = "/tmp/input.pdf"
            with open(tmp_pdf, "wb") as f:
                f.write(uploaded_file.read())
            tables = extract_tables(tmp_pdf, "/tmp")

        st.success("Extragere finalizată!")

        # ── Fișa completată ───────────────────────────────────
        if os.path.exists(TEMPLATE_PATH):
            with st.spinner("Se completează fișa..."):
                out_fisa = f"/tmp/{filename_stem}_fisa.xlsx"
                fill_fisa(tmp_pdf, out_fisa, TEMPLATE_PATH, antet=antet)
            with open(out_fisa, "rb") as f:
                st.download_button(
                    label="⬇️ Descarcă Fișa Completată",
                    data=f.read(),
                    file_name=f"{filename_stem}_fisa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Template _FIsa_Prototip.xlsx nu a fost găsit în repository.")

        # ── Excel cu toate tabelele ───────────────────────────
        with st.spinner("Se generează Excel..."):
            wb = Workbook()
            wb.remove(wb.active)
            header_font = Font(bold=True, name="Arial")
            header_fill = PatternFill("solid", start_color="D9E1F2")

            for sheet_name, rows in tables.items():
                ws = wb.create_sheet(title=sheet_name[:31])
                for r_idx, row in enumerate(rows, start=1):
                    full_row = ["Source File"] + row if r_idx == 1 else [filename_stem] + row
                    for c_idx, value in enumerate(full_row, start=1):
                        if r_idx > 1:
                            try:
                                num = float(value)
                                value = int(num) if num == int(num) else num
                            except (ValueError, TypeError):
                                pass
                        cell = ws.cell(row=r_idx, column=c_idx, value=value)
                        if r_idx == 1:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center")
                for col in ws.columns:
                    max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

            xlsx_buffer = io.BytesIO()
            wb.save(xlsx_buffer)
            xlsx_buffer.seek(0)

        st.download_button(
            label="⬇️ Descarcă Tabele Excel (5 sheet-uri)",
            data=xlsx_buffer,
            file_name=f"{filename_stem}_tabele.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
