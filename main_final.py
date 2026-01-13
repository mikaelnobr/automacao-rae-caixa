import streamlit as st
import sys
import os
import re
import json
import time
from io import BytesIO

# --- PATCH DE METADADOS PARA O DOCLING (Ambiente Cloud/Nuitka) ---
# Este bloco evita o erro 'PackageNotFoundError' ao carregar o Docling na nuvem
try:
    import importlib.metadata as metadata
except ImportError:
    import importlib_metadata as metadata

_original_version = metadata.version
def patched_version(package_name):
    try:
        return _original_version(package_name)
    except metadata.PackageNotFoundError:
        # Versões de segurança para os pacotes do ecossistema Docling
        versions = {
            'docling': '2.15.0',
            'docling-core': '2.9.0',
            'docling-parse': '2.4.0',
            'docling-ibm-models': '1.1.0',
            'pypdfium2': '4.30.0'
        }
        return versions.get(package_name, "1.0.0")
metadata.version = patched_version

# --- IMPORTAÇÃO DAS DEPENDÊNCIAS ---
import pandas as pd
from openpyxl import load_workbook
from docling.document_converter import DocumentConverter
import google.generativeai as genai

# Configuração da página do Streamlit
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# Estilização da Interface (CSS)
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
    div.stStatus { border-radius: 12px; }
    </style>
    """, unsafe_allow_html=True)

def call_gemini(api_key, prompt):
    """Efectua a chamada à API do Gemini com gestão de tentativas."""
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
    st.markdown("##### Processamento Inteligente de Laudos Técnicos")
    st.info("Carregue o laudo em PDF e a planilha modelo para gerar o preenchimento automático via IA.")

    # 1. Barra Lateral de Configuração
    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password", help="Obtenha a sua chave em aistudio.google.com")
        st.divider()
        st.markdown("**Versão Cloud 2.5**")
        st.caption("Desenvolvido para Engenharia CAIXA")

    # 2. Área de Upload de Ficheiros
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. Laudo Técnico (PDF)", type=["pdf"])
    with col2:
        excel_file = st.file_uploader("2. Planilha Modelo (.xlsm)", type=["xlsm"])

    # 3. Botão de Execução
    if st.button("🚀 INICIAR MAPEAMENTO INTELIGENTE"):
        if not api_key:
            st.warning("Por favor, insira a sua Gemini API Key na barra lateral.")
            return
        if not pdf_file or not excel_file:
            st.warning("É necessário carregar ambos os ficheiros (PDF e XLSM).")
            return

        try:
            with st.status("A processar laudo técnico...", expanded=True) as status:
                
                # Passo 1: Conversão do PDF com Docling
                st.write("📖 A extrair estrutura do PDF com Docling...")
                with open("temp_render.pdf", "wb") as f:
                    f.write(pdf_file.getbuffer())
                
                converter = DocumentConverter()
                res = converter.convert("temp_render.pdf")
                md_content = res.document.export_to_markdown()
                # Limpeza simples de quebras de linha duplas
                md_content = re.sub(r'\n\s*\n', '\n', md_content)

                # Passo 2: Análise via Gemini 2.5 Flash
                st.write("🧠 A analisar dados técnicos com Gemini 2.5...")
                prompt = f"""
                Atue como engenheiro revisor da CAIXA. Extraia os dados do laudo abaixo para este formato JSON:
                
                CAMPOS CADASTRAIS:
                - proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel
                
                LÓGICA DO OFÍCIO:
                - oficio: Número após a matrícula no item DOCUMENTOS (ex: 12345 / 3 / CE, o ofício é 3).
                
                COORDENADAS (Formato GMS Limpo):
                - lat_s: Graus, Minutos e Segundos (ex: 06°24'08.8"). NÃO inclua as letras S ou N.
                - long_w: Graus, Minutos e Segundos (ex: 39°18'21.5"). NÃO inclua as letras W ou E.
                
                TABELAS:
                - incidencias: Lista de exatamente 20 números da coluna PESO % (Pág 4).
                - acumulado: Lista de percentuais da coluna % ACUMULADO (Pág 6). Traga apenas meses existentes.

                REGRAS: Retorne apenas JSON puro, use ponto para decimais.
                DOCUMENTO:
                {md_content}
                """
                
                dados = call_gemini(api_key, prompt)

                # Passo 3: Escrita na Planilha Excel
                st.write("📊 A gravar dados e a preservar Macros VBA...")
                # Carregar o ficheiro Excel em memória
                wb = load_workbook(BytesIO(excel_file.read()), keep_vba=True)
                wb.calculation.fullCalcOnLoad = True

                def to_f(v):
                    if isinstance(v, (int, float)): return v
                    try: return float(str(v).replace(',', '.').replace('%', '').strip())
                    except: return 0

                # Preenchimento da Aba Início Vistoria
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
                        if key == "valor_terreno":
                            ws[cell] = to_f(val)
                        else:
                            ws[cell] = str(val).upper() if val else ""
                    # Configura a finalidade para activar as abras dependentes
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                # Preenchimento da Aba RAE
                if "RAE" in wb.sheetnames:
                    ws_rae = wb["RAE"]
                    ws_rae.sheet_state = 'visible' # Força a visibilidade da aba
                    
                    ws_rae["AH66"] = to_f(dados.get("valor_imovel", 0))
                    
                    incs, acus = dados.get("incidencias", []), dados.get("acumulado", [])
                    # Gravação das incidências (20 linhas)
                    for i in range(20):
                        val = incs[i] if i < len(incs) else 0
                        ws_rae[f"S{69+i}"] = to_f(val)
                    # Gravação do acumulado proposto
                    for i in range(len(acus)):
                        if i < 37: # Limite da coluna na planilha
                            ws_rae[f"AE{72+i}"] = to_f(acus[i])

                # Preparação do Ficheiro de Saída
                output = BytesIO()
                wb.save(output)
                processed_data = output.getvalue()
                
                # Gerar nome sugerido: RAE_PRIMEIRONOME.xlsm
                proponente = dados.get("proponente", "").strip()
                primeiro_nome = proponente.split(' ')[0].upper() if proponente else "FINAL"
                nome_ficheiro = f"RAE_{primeiro_nome}.xlsm"

                status.update(label="✅ Processamento terminado com sucesso!", state="complete", expanded=False)

            st.balloons()
            st.success(f"O laudo do proponente {primeiro_nome} foi processado!")
            
            st.download_button(
                label="📥 DESCARREGAR PLANILHA PREENCHIDA",
                data=processed_data,
                file_name=nome_ficheiro,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

        except Exception as e:
            st.error(f"Erro Crítico durante o processamento: {e}")
        finally:
            # Limpeza de ficheiro temporário
            if os.path.exists("temp_render.pdf"):
                os.remove("temp_render.pdf")

if __name__ == "__main__":
    main()
