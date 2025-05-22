import streamlit as st
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
import zipfile
import time
import datetime
import sys
from io import BytesIO
import warnings

warnings.filterwarnings(
    "ignore",
    message="File contains an invalid specification for 0. This will be removed"
)

st.set_page_config(
    page_title="SPDO Resumo de Fichas Técnicas",
    page_icon="logo_fgv.png",
    layout="wide"
)
st.logo('logo_ibre.png')

# Funções utilitárias (igual ao seu código)
def safe_cell(ws, cell_ref, default="-"):
    try:
        if ws is None:
            return default
        val = ws[cell_ref].value
        if val is None or (isinstance(val, str) and not val.strip()):
            return default
        return val
    except Exception:
        return default

def format_date(val, default="-"):
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%d/%m/%Y")
    return default

st.title("Resumo de Fichas Técnicas")
st.markdown("Faça upload da pasta **Fichas Tecnicas** compactada em .zip e clique em Iniciar.")

# 1) Upload do zip
uploaded_zip = st.file_uploader(
    "Upload da pasta Fichas Tecnicas (zip)", type="zip", accept_multiple_files=False
)

if uploaded_zip and st.button("Iniciar Processamento"):
    # Descompacta em memória
    with zipfile.ZipFile(uploaded_zip) as z:
        # Lista só os .xlsx dentro de Fichas Tecnicas/...
        xlsx_paths = [
            name for name in z.namelist()
            if name.startswith("Fichas Tecnicas/") and name.lower().endswith(".xlsx")
        ]
        total_files = len(xlsx_paths)
        if total_files == 0:
            st.error("Nenhum arquivo .xlsx encontrado em Fichas Tecnicas/")
            st.stop()

        progress_bar = st.progress(0)
        status_text = st.empty()

        data_by_mother = {}
        start = time.perf_counter()
        for idx, rel_path in enumerate(xlsx_paths, 1):
            progress_bar.progress(idx / total_files)
            status_text.text(f"Processando {idx}/{total_files}: {rel_path}")

            # extrai nome das pastas mãe/filha e o nome do arquivo
            parts = Path(rel_path).parts  
            _, mother_name, child_name, filename = parts
            tipo = (
                "Materiais" if "Materiais" in child_name else
                "Equipamentos" if "Equipamentos" in child_name else
                child_name
            )

            # abre o XLSX em memória
            with z.open(rel_path) as f:
                wb = load_workbook(BytesIO(f.read()), read_only=True, data_only=True)

            nome_mother = mother_name.upper()
            nome_arquivo = Path(filename).stem

            # ————————————————
            # 1) lógica de seleção da worksheet
            if nome_mother == "ECON-DNIT" and tipo == "Equipamentos":
                ws = wb["Ficha de insumo"] if "Ficha de insumo" in wb.sheetnames else wb.active
            elif len(wb.sheetnames) > 1:
                ws = wb.worksheets[0]
            elif len(wb.sheetnames) == 1:
                ws = wb.active
            else:
                ws = None

            # ————————————————
            # 2) lógica de definição de células por cliente
            if nome_mother == "CAGECE":
                criac_ref, atual_ref, ext_ref = "D5", "H5", "K9"
            elif nome_mother == "DER-MG":
                if nome_arquivo == "MATRO-1794":
                    criac_ref, atual_ref, ext_ref = "D3", "I3", "N9"
                elif nome_arquivo in ["EQRO-1508", "EQRO-5651", "EQRO-5652"]:
                    criac_ref, atual_ref, ext_ref = "D4", "H4", "N9"
                else:
                    criac_ref, atual_ref, ext_ref = "D5", "H5", "K9"
            elif nome_mother == "ECON-DNIT":
                if tipo == "Equipamentos":
                    criac_ref, atual_ref, ext_ref = "D5", "H5", "K7"
                elif nome_arquivo in ["E8888", "E8351", "E8306"]:
                    criac_ref, atual_ref, ext_ref = "D4", "H4", "N9"
                else:
                    if nome_arquivo in ["M7062", "M7099", "M7228", "M7618", "M7642"]:
                        criac_ref, atual_ref, ext_ref = "D3", "I3", "N9"
                    else:
                        criac_ref, atual_ref, ext_ref = "D5", "H5", "K7"
            elif nome_mother == "SANEAGO":
                criac_ref, atual_ref, ext_ref = "D3", "I3", "N9"
            elif nome_mother == "SICRO":
                    if nome_arquivo in [
                        "M0291", "M0292", "M0293", "M0294", "M0295", "M0296", "M0297",
                        "M0375", "M0713", "M0714", "M0715", "M0716", "M1728", "M1729",
                        "M1730", "M1894", "M1895", "M1899", "M1920", "M1923", "M2017",
                        "M2019", "M2020", "M2029", "M2030", "M2031", "M2032", "M2033",
                        "M2035", "M2089", "M2090", "M2091", "M2094", "M2095", "M2096",
                        "M2101", "M2325", "M2326", "M2327", "M2328", "M2329", "M2330",
                        "M2331", "M2332", "M2333", "M2585", "M2586", "M2587", "M2588",
                        "M2589", "M2590", "M2591", "M3091", "M3094", "M3523", "M3825",
                        "M3829", "M3843", "M3853"
                    ]:
                        criac_ref, atual_ref, ext_ref = "D3", "I3", "N9"
                    elif nome_arquivo in [
                        "A9316", "A9344", "A9345", "E9066", "E9145", "E9259", "E9260",
                        "E9261", "E9262", "E9263", "E9264", "E9265", "E9267", "E9543",
                        "E9575", "E9579", "E9667"
                    ]:
                        criac_ref, atual_ref, ext_ref = "D4", "H4", "N9"
                    else:
                        criac_ref, atual_ref, ext_ref = "D5", "H5", "K9"
                
            
            else:
                criac_ref, atual_ref, ext_ref = "D5", "H5", "K9"
            # ————————————————

            # agora extrai as células usando safe_cell()
            raw_criacao     = safe_cell(ws, criac_ref)
            raw_atualizacao = safe_cell(ws, atual_ref)
            raw_externo     = safe_cell(ws, ext_ref)

            criacao_fmt     = format_date(raw_criacao)
            atualizacao_fmt = format_date(raw_atualizacao)

            data_by_mother.setdefault(mother_name, []).append({
                "Arquivo":        filename,
                "Tipo":           tipo,
                "Criação":        criacao_fmt,
                "Atualização":    atualizacao_fmt,
                "Código Externo": raw_externo,
            })



        elapsed = time.perf_counter() - start
        st.success(f"Processamento concluído em {elapsed:.2f} segundos!")

        # Gera o Excel para download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for sheet, recs in data_by_mother.items():
                pd.DataFrame(recs).to_excel(writer, sheet_name=sheet[:31], index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Baixar Resumo como Excel",
            data=buffer,
            file_name="Resumo_Fichas_Tecnicas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
