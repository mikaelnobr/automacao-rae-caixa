import streamlit as st
import sys
import os
import re
import json
import time
from io import BytesIO

# --- PATCH PARA AMBIENTES CLOUD (Evita erro de metadados no Docling) ---
try:
    import importlib.metadata as metadata
except ImportError:
    import importlib_metadata as metadata

_original_version = metadata.version
def patched_version(package_name):
    try:
        return _original_version(package_name)
    except metadata.PackageNotFoundError:
        # Versões seguras para as dependências do Docling
        versions = {
            'docling': '2.15.0',
            'docling-core': '2.8.0',
            'docling-parse': '2.4.0',
            'docling-ibm-models': '1.1.0'
        }
        return versions.get(package_name, "1.0.0")
metadata.version = patched_version

# --- IMPORTAÇÃO DAS DEPENDÊNCIAS ---
import pandas as pd
from openpyxl import load_workbook
from docling.document_converter import DocumentConverter
import google.generativeai as genai

# Configuração da página Web
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# Estilização Premium via CSS
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 3.5em;
        background-color: #4f46e5;
        color: white;
        font-weight: bold;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #4338ca;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 10px;
        background-color: #10b981;
        color: white;
        border: none;
    }
    div.stStatus { border-radius: 15px; }
    </style>
    """, unsafe_allow_html=True)

def call_gemini(api_key, prompt):
    genai.configure(api_key=api_key)
    # Modelo 2.5 Flash conforme validado nos logs de sucesso
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
    st.markdown("### Processamento Inteligente via Docling & Gemini AI")
    st.info("Utilize esta ferramenta para preencher automaticamente planilhas RAE a partir de laudos em PDF.")

    # 1. Configuração de Acesso
    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password", help="Gere sua chave em aistudio.google.com")
        st.divider()
        st.markdown("Developed by Mikael Engineering AI")

    # 2. Upload de Arquivos
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. Envie o Laudo (PDF)", type=["pdf"])
    with col2:
        excel_file = st.file_uploader("2. Envie a Planilha Modelo (.xlsm)", type=["xlsm"])

    if st.button("🚀 INICIAR PROCESSAMENTO"):
        if not api_key:
            st.warning("Insira sua Gemini API Key na barra lateral.")
            return
        if not pdf_file or not excel_file:
            st.warning("Por favor, faça o upload de ambos os arquivos.")
            return

        try:
            with st.status("Executando inteligência artificial...", expanded=True) as status:
                
                # Passo 1: Converter PDF para Markdown via Docling
                st.write("📖 Lendo e estruturando laudo com Docling...")
                # Salvamento temporário em memória (Streamlit Cloud exige isso)
                with open("temp_doc.pdf", "wb") as f:
                    f.write(pdf_file.getbuffer())
                
                converter = DocumentConverter()
                res = converter.convert("temp_doc.pdf")
                md_content = res.document.export_to_markdown()
                md_content = re.sub(r'\n\s*\n', '\n', md_content)

                # Passo 2: Extração Inteligente
                st.write("🧠 Consultando Gemini 2.5 Flash para análise técnica...")
                prompt = f"""
                Atue como engenheiro revisor da CAIXA. Extraia os dados do documento abaixo para este formato JSON:
                
                CAMPOS CADASTRAIS:
                - proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel
                
                LÓGICA DO OFÍCIO:
                - oficio: Número após a matrícula no item DOCUMENTOS (ex: 12345 / 3 / CE, o ofício é 3).
                
                COORDENADAS (Obrigatório GMS puro):
                - lat_s: Formato Graus, Minutos e Segundos (ex: 06°24'08.8"). NÃO use letras S/N.
                - long_w: Formate como Graus, Minutos e Segundos (ex: 39°18'21.5"). NÃO use letras W/E.
                
                TABELAS:
                - incidencias: Lista de 20 números decimais (coluna PESO % da pág 4).
                - acumulado: Lista de percentuais da coluna % ACUMULADO (Cronograma pág 6). Traga apenas os meses que possuem dados preenchidos.

                REGRAS: JSON puro, ponto decimal.
                DOCUMENTO:
                {md_content}
                """
                
                dados = call_gemini(api_key, prompt)

                # Passo 3: Preencher Excel
                st.write("📊 Gravando dados na planilha e preservando Macros...")
                # Lendo o arquivo enviado da memória
                wb = load_workbook(BytesIO(excel_file.read()), keep_vba=True)
                wb.calculation.fullCalcOnLoad = True

                def to_f(v):
                    if isinstance(v, (int, float)): return v
                    try: return float(str(v).replace(',', '.').replace('%', '').strip())
                    except: return 0

                # Aba Início Vistoria
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
                        if key in ["valor_terreno"]:
                            ws[cell] = to_f(val)
                        else:
                            ws[cell] = str(val).upper() if val else ""
                    # Força seleção da finalidade para habilitar a RAE
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                # Aba RAE
                if "RAE" in wb.sheetnames:
                    ws_rae = wb["RAE"]
                    ws_rae.sheet_state = 'visible' # Garante que a aba esteja visível
                    ws_rae["AH66"] = to_f(dados.get("valor_imovel", 0))
                    
                    incs, acus = dados.get("incidencias", []), dados.get("acumulado", [])
                    for i in range(20):
                        ws_rae[f"S{69+i}"] = to_f(incs[i]) if i < len(incs) else 0
                    for i in range(len(acus)):
                        if i < 37: ws_rae[f"AE{72+i}"] = to_f(acus[i])

                # Preparar Download
                output = BytesIO()
                wb.save(output)
                final_data = output.getvalue()
                
                proponente = dados.get("proponente", "").strip()
                primeiro_nome = proponente.split(' ')[0].upper() if proponente else "PROCESSADA"
                nome_final = f"RAE_{primeiro_nome}.xlsm"

                status.update(label="✅ Tudo pronto!", state="complete", expanded=False)

            st.success(f"Laudo de {primeiro_nome} processado com sucesso!")
            st.download_button(
                label="📥 BAIXAR RAE PREENCHIDA",
                data=final_data,
                file_name=nome_final,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

        except Exception as e:
            st.error(f"Erro Crítico: {e}")
        finally:
            if os.path.exists("temp_doc.pdf"):
                os.remove("temp_doc.pdf")

if __name__ == "__main__":
    main()