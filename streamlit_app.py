"""
Gerador de Folhas de Ponto — mala direta em massa a partir de planilha Excel
+ template Word, com seleção de período e download em ZIP.

Como rodar localmente:
    pip install streamlit pandas openpyxl
    streamlit run streamlit_app.py
"""
import io
import re
import zipfile
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import docx_engine as eng

st.set_page_config(page_title="Gerar Folha de Ponto", layout="centered")

REQUIRED_COLUMNS = [
    "Nome (Empresa)", "DESCRIÇÃO DO LOCAL", "ENDEREÇO", "NUMERO", "CIDADE",
    "BAIRRO", "MATRICULA", "NOME", "ADMISSAO", "FUNÇÃO", "HORARIO",
]

COLUMN_TO_FIELD = {
    "Nome (Empresa)": "NOME_EMPRESA",
    "DESCRIÇÃO DO LOCAL": "DESCRIÇÃO_DO_LOCAL",
    "ENDEREÇO": "ENDEREÇO",
    "NUMERO": "NUMERO",
    "CIDADE": "CIDADE",
    "BAIRRO": "BAIRRO",
    "MATRICULA": "MATRICULA",
    "NOME": "NOME",
    "ADMISSAO": "ADMISSAO",
    "FUNÇÃO": "FUNÇÃO",
    "HORARIO": "HORARIO",
}


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '', str(name)).strip()
    return name[:150]


def try_convert_pdf(docx_bytes: bytes) -> bytes | None:
    """Tenta converter para PDF via LibreOffice, se disponível no ambiente."""
    import subprocess
    import tempfile
    import os
    try:
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = os.path.join(tmp, "doc.docx")
            with open(docx_path, "wb") as f:
                f.write(docx_bytes)
            result = subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
                capture_output=True, timeout=60,
            )
            pdf_path = os.path.join(tmp, "doc.pdf")
            if result.returncode == 0 and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return f.read()
    except Exception:
        return None
    return None


st.title("Gerar Folha de Ponto")

st.subheader("1. Planilha de colaboradores (.xlsx)")
xlsx_file = st.file_uploader("Envie a planilha base", type=["xlsx"])

st.subheader("2. Base Word (.docx)")
docx_file = st.file_uploader(
    "Envie a planilha base",
    type=["docx"],
)

st.subheader("3. Período trabalhado")
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data inicial", value=date.today().replace(day=16))
with col2:
    default_end = (start_date.replace(day=1) + timedelta(days=32)).replace(day=15)
    end_date = st.date_input("Data final", value=default_end)


gerar = st.button("Gerar folhas de ponto", type="primary", use_container_width=True)

if gerar:
    if not xlsx_file or not docx_file:
        st.error("Envie a planilha e o template antes de gerar.")
        st.stop()
    if start_date > end_date:
        st.error("A data inicial não pode ser depois da data final.")
        st.stop()

    df = pd.read_excel(xlsx_file)
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        st.error(f"Colunas faltando na planilha: {', '.join(missing_cols)}")
        st.stop()

    template_bytes = docx_file.read()
    try:
        base_content, others = eng.load_template_xml(template_bytes)
    except Exception as e:
        st.error(f"Não consegui ler o template: {e}")
        st.stop()

    zip_buffer = io.BytesIO()
    progress = st.progress(0, text="Gerando documentos...")
    status_area = st.empty()

    faltando_dados = []
    gerados = 0
    total = len(df)

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, row in df.iterrows():
            values = {}
            campos_vazios = []
            for col, field in COLUMN_TO_FIELD.items():
                raw = row.get(col, "")
                if pd.isna(raw) or str(raw).strip() == "":
                    campos_vazios.append(col)
                    raw = ""
                if field == "ADMISSAO" and raw != "":
                    try:
                        raw = pd.to_datetime(raw).strftime("%d/%m/%Y")
                    except Exception:
                        raw = str(raw)
                values[field] = str(raw)
            MESES = {
            1: "JAN",
            2: "FEV",
            3: "MAR",
            4: "ABR",
            5: "MAI",
            6: "JUN",
            7: "JUL",
            8: "AGO",
            9: "SET",
            10: "OUT",
            11: "NOV",
            12: "DEZ"}
    
            values["PERIODO"] = f"{MESES[start_date.month]}/{MESES[end_date.month]}"
            values["ANO"] = str(end_date.year) 

            nome = values.get("NOME", f"colaborador_{i}").strip() or f"colaborador_{i}"
            matricula = values.get("MATRICULA", "").strip()

            if campos_vazios:
                faltando_dados.append((nome or f"linha {i+2}", campos_vazios))

            try:
                docx_bytes = eng.build_docx(base_content, others, values, start_date, end_date)
            except Exception as e:
                faltando_dados.append((nome, [f"ERRO ao gerar: {e}"]))
                continue

            filename_base = sanitize_filename(f"{matricula} - {nome}") if matricula else sanitize_filename(nome)
            zf.writestr(f"{filename_base}.docx", docx_bytes)

            if gerar_pdf:
                pdf_bytes = try_convert_pdf(docx_bytes)
                if pdf_bytes:
                    zf.writestr(f"{filename_base}.pdf", pdf_bytes)

            gerados += 1
            progress.progress((i + 1) / total, text=f"Gerando documentos... ({i+1}/{total})")

    progress.empty()
    status_area.success(f"✅ {gerados} de {total} documentos gerados com sucesso.")

    if faltando_dados:
        with st.expander(f"⚠️ {len(faltando_dados)} colaborador(es) com dados faltando ou erro"):
            for nome, campos in faltando_dados:
                st.write(f"**{nome}**: {', '.join(campos)}")

    zip_buffer.seek(0)
    st.download_button(
        "⬇️ Baixar ZIP com todas as folhas de ponto",
        data=zip_buffer.getvalue(),
        file_name=f"folhas_de_ponto_{start_date.strftime('%d-%m-%Y')}_a_{end_date.strftime('%d-%m-%Y')}.zip",
        mime="application/zip",
        use_container_width=True,
    )
