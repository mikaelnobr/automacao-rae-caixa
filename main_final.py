import streamlit as st
import sys
import os
import re
import json
import time
import tempfile
from io import BytesIO

# --- CONFIGURAÇÃO INICIAL (OBRIGATORIAMENTE O PRIMEIRO COMANDO ST) ---
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# --- PATCH DE METADADOS ULTRA-ROBUSTO ---
try:
    import importlib.metadata as metadata
except ImportError:
    import importlib_metadata as metadata

_original_version = metadata.version
def patched_version(package_name):
    try:
        return _original_version(package_name)
    except Exception:
        versions = {
            'docling': '2.15.0',
            'docling-core': '2.9.0',
            'docling-parse': '2.4.0',
            'docling-ibm-models': '1.1.0',
            'pypdfium2': '4.30.0',
            'openpyxl': '3.1.5',
            'transformers': '4.40.0',
            'torch': '2.2.0',
            'torchvision': '0.17.0',
            'timm': '0.9.16',
            'optree': '0.11.0'
        }
        return versions.get(package_name, "1.0.0")
metadata.version = patched_version

# --- IMPORTAÇÃO DAS DEPENDÊNCIAS ---
try:
    import pandas as pd
    from openpyxl import load_workbook
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.datamodel.base_models import InputFormat
    import google.generativeai as genai
    import onnxruntime
    import transformers
    import timm
    import optree 
    DEPENDENCIAS_OK = True
    ERRO_IMPORT = ""
except ImportError as e:
    DEPENDENCIAS_OK = False
    ERRO_IMPORT = str(e)

# --- ESTILIZAÇÃO ---
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    .stButton>button {
        width: 100%; border-radius: 8px; height: 3.5em;
        background-color: #4f46e5; color: white; font-weight: bold; border: none;
    }
    .stDownloadButton>button {
        width: 100%; border-radius: 8px;
        background-color: #059669; color: white; border: none;
    }
    </style>
    """, unsafe_allow_html=True)

# Cache para o conversor (Evita estourar a RAM do servidor)
@st.cache_resource
def get_converter():
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = True 
    pipeline_options.table_structure_options.do_cell_matching = True
    return DocumentConverter(
        allowed_formats=[InputFormat.PDF],
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
        }
    )

def call_gemini(api_key, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    for attempt in range(3):
        try:
            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            return json.loads(response.text)
        except Exception as e:
            if attempt == 2: raise e
            time.sleep(2)

def main():
    st.title("🏛️ Automação RAE CAIXA")
    st.markdown("##### Inteligência Artificial para Engenharia")

    if not DEPENDENCIAS_OK:
        st.error(f"❌ Erro de Dependências: {ERRO_IMPORT}")
        st.info("💡 Para corrigir, adicione 'optree', 'timm', 'torch', 'torchvision' e 'transformers' ao requirements.txt.")
        return

    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")
        st.divider()
        st.caption("v3.4 - Fixed Page Config & Cache")

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. Enviar Laudo (PDF)", type=["pdf"])
    with col2:
        excel_file = st.file_uploader("2. Enviar Modelo (.xlsm)", type=["xlsm"])

    if st.button("🚀 INICIAR PROCESSAMENTO"):
        if not api_key or not pdf_file or not excel_file:
            st.warning("Preencha a chave e carregue os ficheiros.")
            return

        try:
            with st.status("A processar laudo técnico...", expanded=True) as status:
                # Ficheiro temporário
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(pdf_file.getbuffer())
                    tmp_path = tmp.name

                try:
                    st.write("📖 Lendo estrutura do PDF com Docling...")
                    converter = get_converter()
                    res = converter.convert(tmp_path)
                    md_content = re.sub(r'\n\s*\n', '\n', res.document.export_to_markdown())
                finally:
                    if os.path.exists(tmp_path): os.remove(tmp_path)

                st.write("🧠 IA: Extraindo dados técnicos...")
                prompt = f"""
                Atue como engenheiro revisor da CAIXA. Extraia os dados para JSON:
                - CAMPOS: proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel
                - OFICIO: Número após a matrícula em DOCUMENTOS (ex: 12345 / 3 / CE, ofício é 3).
                - COORDENADAS: GMS puro (ex: 06°24'08.8"). SEM letras no final.
                - TABELAS: 'incidencias' (20 números PESO %), 'acumulado' (percentuais % ACUMULADO).
                DOCUMENTO: {md_content}
                """
                dados = call_gemini(api_key, prompt)

                st.write("📊 Gravando na planilha Excel...")
                wb = load_workbook(BytesIO(excel_file.read()), keep_vba=True)
                wb.calculation.fullCalcOnLoad = True

                def to_f(v):
                    if isinstance(v, (int, float)): return v
                    try: return float(str(v).replace(',', '.').replace('%', '').strip())
                    except: return 0

                if "Início Vistoria" in wb.sheetnames:
                    ws = wb["Início Vistoria"]
                    mapping = {
                        "G43": "proponente", "AJ43": "cpf_cnpj", "AP43": "ddd", "AR43": "telefone",
                        "G49": "endereco", "AD49": "lat_s", "AH49": "long_w", "AL49": "complemento",
                        "G51": "bairro", "V51": "cep", "AA51": "municipio", "AS51": "uf_vistoria",
                        "AS53": "uf_registro", "G53": "valor_terreno", "Q53": "matricula",
                        "AA53": "oficio", "AJ53": "comarca"
                    }
                    for cell, key in mapping.items():
                        val = dados.get(key, "")
                        ws[cell] = to_f(val) if key == "valor_terreno" else str(val).upper()
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                if "RAE" in wb.sheetnames:
                    ws_rae = wb["RAE"]
                    ws_rae.sheet_state = 'visible'
                    ws_rae["AH66"] = to_f(dados.get("valor_imovel", 0))
                    incs, acus = dados.get("incidencias", []), dados.get("acumulado", [])
                    for i in range(20):
                        ws_rae[f"S{69+i}"] = to_f(incs[i]) if i < len(incs) else 0
                    for i in range(len(acus)):
                        if i < 37: ws_rae[f"AE{72+i}"] = to_f(acus[i])

                output = BytesIO()
                wb.save(output)
                processed_data = output.getvalue()
                
                proponente = dados.get("proponente", "").strip()
                primeiro_nome = proponente.split(' ')[0].upper() if proponente else "FINAL"
                nome_arq = f"RAE_{primeiro_nome}.xlsm"

                status.update(label="✅ Mapeamento concluído!", state="complete", expanded=False)

            st.balloons()
            st.download_button(
                label=f"📥 BAIXAR RAE - {primeiro_nome}",
                data=processed_data,
                file_name=nome_arq,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

        except Exception as e:
            if "libGL.so.1" in str(e):
                st.error("❌ Erro de Dependência do Sistema (libGL.so.1)")
                st.markdown("""
                **Como resolver:**
                Crie o arquivo **`packages.txt`** no GitHub com:
                ```text
                libgl1
                libglib2.0-0
                ```
                """)
            else:
                st.error(f"Erro no processamento: {e}")

if __name__ == "__main__":
    main()
