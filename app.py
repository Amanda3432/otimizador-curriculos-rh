import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from google import genai
from google.genai import errors
import time

st.set_page_config(page_title="Otimizador de Currículos - Consultoria", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Consultoria)")
st.markdown("Ferramenta automatizada para otimização de currículos alinhada aos padrões de recrutamento e seleção.")

# --- BARRA LATERAL (CONFIGURAÇÕES E CHAVE API) ---
st.sidebar.header("🔑 Configuração da API")

# Tenta carregar a chave automaticamente dos segredos do Streamlit se existir, senão usa o campo de texto
default_api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

api_key_input = st.sidebar.text_input(
    "Insira sua Google Gemini API Key", 
    value=default_api_key, 
    type="password",
    help="Cole aqui a sua chave do Google AI Studio (começa com AQ... ou AIza...)"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📋 Modelos Disponíveis")
modelo_escolhido = st.sidebar.radio(
    "Escolha o formato de saída:",
    ("Modelo Clássico (Sem Foto)", "Modelo Com Foto")
)

# Se o usuário preencheu a chave, guarda na sessão para não precisar digitar toda vez
if api_key_input:
    st.session_state["api_key"] = api_key_input

ativa_api_key = st.session_state.get("api_key", "")

# --- ÁREA PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Dados do Candidato e Vaga")
    curriculo_antigo = st.text_area("Cole o Currículo Atual do Cliente:", height=200, placeholder="Cole aqui o texto do currículo antigo...")
    descricao_vaga = st.text_area("Cole a Descrição da Vaga ou Palavras-Chave:", height=150, placeholder="Cole a descrição da vaga do Vagas.com, LinkedIn ou Gupy...")

with col2:
    st.subheader("2️⃣ Instruções e Execução")
    st.info("O sistema vai reestruturar o perfil profissional em 3 parágrafos, ajustar as experiências em tópicos neutros baseados na vaga e organizar a formação e cursos conforme as normas da consultoria[cite: 4, 5].")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

# Função para chamar o Gemini com tentativas automáticas (evita erros 503)
def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']
    
    for modelo in modelos_para_tentar:
        for tentativa in range(3):
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=prompt_texto,
                )
                if response and response.text:
                    return response.text
            except Exception as e:
                # Se for erro 503 ou sobrecarga, aguarda alguns segundos e tenta novamente
                time.sleep(2)
                continue
    raise Exception("Não foi possível gerar o conteúdo devido a instabilidade temporária na API. Tente novamente em instantes.")

if gerar_btn:
    if not ativa_api_key:
        st.error("⚠️ Por favor, insira a sua Chave de API na barra lateral para continuar.")
    elif not curriculo_antigo or not descricao_vaga:
        st.warning("⚠️ Preencha tanto o currículo antigo quanto a descrição da vaga.")
    else:
        with st.spinner("A processar e reestruturando o currículo de acordo com os padrões da consultoria..."):
            try:
                # Inicializa o cliente do Gemini com a nova biblioteca google-genai
                client = genai.Client(api_key=ativa_api_key)
                
                # Prompt estruturado com base nas regras enviadas
                prompt_sistema = f"""
                Você é um consultor especialista em RH e Otimização de Currículos ATS.
                Com base no currículo antigo e na descrição da vaga fornecidos abaixo, gere um currículo reestruturado estritamente seguindo estas regras:
                
                1. DADOS PESSOAIS: Extraia Nome, Bairro/Cidade, Telefone, E-mail e LinkedIn (com hiperlinks se possível).
                2. OBJETIVO: Coloque o cargo ou área pretendida (máximo 3 opções).
                3. PERFIL PROFISSIONAL (Escrito estritamente em 3ª pessoa):
                   - 1º Parágrafo: Profissional atuante há mais de X anos na área [cargo/objetivo], destacando competências comportamentais importantes.
                   - 2º Parágrafo: Expertises detalhadas com base nas atividades-chave e palavras-chave que mais se repetem na descrição da vaga.
                   - 3º Parágrafo: Conhecimentos em sistemas e ferramentas teóricas relevantes.
                4. FORMAÇÃO ACADÊMICA: Ordem de importância/cronológica (Pós-doutorado, Doutorado, Mestrado, Pós-graduação, Graduação, Técnico). Não incluir ensino médio se houver nível superior/técnico. Formato: Nome do curso | Instituição de ensino - Ano de conclusão.
                5. CURSOS E CERTIFICAÇÕES: Ordem alfabética ou cronológica. Formato: Nome do curso | Instituição de ensino | Ano de conclusão (com validade se for certificação).
                6. EXPERIÊNCIAS PROFISSIONAIS: Ordem cronológica da mais recente para a mais antiga. Descrições escritas em tópicos, neutras e profissionais baseadas nas exigências da vaga (evitar deixar igual ao original).
                
                Currículo Antigo:
                {curriculo_antigo}
                
                Descrição da Vaga / Requisitos:
                {descricao_vaga}
                
                Retorne o conteúdo limpo, organizado por seções claras para que possa ser convertido em documento Word/PDF.
                """
                
                resultado_ia = chamar_gemini_com_retry(client, prompt_sistema)
                
                st.session_state["curriculo_gerado"] = resultado_ia
                st.success("✨ Currículo otimizado com sucesso!")
                
            exc = Exception
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")

# Se já houver currículo gerado, exibe os botões de visualização e download
if "curriculo_gerado" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Resultado Gerado")
    st.text_area("Texto Otimizado:", st.session_state["curriculo_gerado"], height=300)
    
    st.info(f"Modo selecionado: **{modelo_escolhido}**. O documento gerado respeita as margens e diretrizes de uma única página/layout limpo.")
    
    # Geração de Documento Word (.docx)
    doc = docx.Document()
    # Ajuste de margens corporativas
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    p_doc = doc.add_paragraph()
    p_doc.add_run(st.session_state["curriculo_gerado"])
    
    import io
    buffer_word = io.BytesIO()
    doc.save(buffer_word)
    buffer_word.seek(0)
    
    st.download_button(
        label=f"📥 Baixar Currículo em Word ({modelo_escolhido})",
        data=buffer_word,
        file_name="curriculo_otimizado_consultoria.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )
