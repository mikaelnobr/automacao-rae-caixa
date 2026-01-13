import streamlit as st
import sys
import os
import re
import json
import time
from io import BytesIO

# --- PATCH DE METADADOS PARA O DOCLING ---
# Evita que o app trave ao procurar versões de pacotes no ambiente Linux do Streamlit
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
            'openpyxl': '3.1.5'
        }
        return versions.get(package_name, "1.0.0")
metadata.version = patched_version

# --- IMPORTAÇÃO DAS DEPENDÊNCIAS COM TRATAMENTO DE ERRO ---
try:
    import pandas as pd
    from openpyxl import load_workbook
    from docling.document_converter import DocumentConverter
    import google.generativeai as genai
    import onnxruntime
    DEPENDENCIAS_OK = True
except ImportError as e:
    DEPENDENCIAS_OK = False
    ERRO_IMPORT = str(e)

# Configuração da página
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# CSS para interface
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3.5em;
        background-color: #4f46e5;
        color: white;
        font-weight: bold;
        border: none;
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 8px;
        background-color: #059669;
        color: white;
        border: none;
    }
    </style>
    """, unsafe_allow_html=True)

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
    st.markdown("##### Inteligência Artificial para Laudos de Engenharia")

    if not DEPENDENCIAS_OK:
        st.error(f"❌ Erro de Dependências: {ERRO_IMPORT}")
        st.warning("O Streamlit Cloud não instalou as bibliotecas do arquivo 'requirements.txt'.")
        
        with st.expander("🛠️ Como resolver este erro agora", expanded=True):
            st.markdown("""
            1. Vá ao painel do **Streamlit Cloud** (share.streamlit.io).
            2. Localize seu app e clique nos **três pontos (...)** no canto direito.
            3. Selecione **'Reboot App'**. 
            4. Se não funcionar, clique em **'Delete'** e suba o app novamente apontando para o repositório. Isso força a limpeza do cache de instalação.
            
            **Seu arquivo 'requirements.txt' no GitHub deve ser EXATAMENTE assim:**
            ```text
            streamlit
            pandas
            openpyxl
            docling
            google-generativeai
            onnxruntime
            ```
            """)
        return

    st.info("Carregue o laudo em PDF e a planilha modelo para iniciar.")

    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")
        st.divider()
        st.caption("v2.5 - Streamlit Cloud Edition")

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. Laudo PDF", type=["pdf"])
    with col2:
        excel_file = st.file_uploader("2. Planilha Modelo (.xlsm)", type=["xlsm"])

    if st.button("🚀 PROCESSAR DOCUMENTOS"):
        if not api_key or not pdf_file or not excel_file:
            st.warning("Preencha todos os campos e envie os arquivos.")
            return

        try:
            with st.status("Trabalhando no laudo...", expanded=True) as status:
                st.write("📖 Extraindo dados do PDF com Docling...")
                with open("temp_file.pdf", "wb") as f:
                    f.write(pdf_file.getbuffer())
                
                converter = DocumentConverter()
                res = converter.convert("temp_file.pdf")
                md_content = re.sub(r'\n\s*\n', '\n', res.document.export_to_markdown())

                st.write("🧠 Analisando informações com Gemini 2.5...")
                prompt = f"""
                Atue como engenheiro revisor da CAIXA. Extraia os dados para JSON:
                - CAMPOS: proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel
                - OFICIO: Número após a matrícula em DOCUMENTOS (ex: 12345 / 3 / CE, ofício é 3).
                - COORDENADAS: GMS puro (ex: 06°24'08.8"). SEM letras S, N, W ou E.
                - TABELAS: 'incidencias' (20 números PESO %), 'acumulado' (percentuais % ACUMULADO).
                DOCUMENTO: {md_content}
                """
                dados = call_gemini(api_key, prompt)

                st.write("📊 Preenchendo planilha Excel...")
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

                status.update(label="✅ Concluído!", state="complete", expanded=False)

            st.balloons()
            st.download_button(
                label="📥 BAIXAR RAE PREENCHIDA",
                data=processed_data,
                file_name=nome_arq,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
        finally:
            if os.path.exists("temp_file.pdf"): os.remove("temp_file.pdf")

if __name__ == "__main__":
    main()
