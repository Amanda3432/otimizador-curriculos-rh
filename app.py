import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.graphics.shapes import Drawing, Rect, Polygon, Circle, Path
from google import genai
import time
import io
import os
import re

st.set_page_config(page_title="Otimizador de Currículos - Padrão Sala de Emprego", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Sala de Emprego)")
st.markdown("Ferramenta automatizada ajustada estritamente à fonte Times New Roman/Roman, cores pretas sólidas e ícones da imagem de referência.")

# --- BARRA LATERAL (CONFIGURAÇÕES E CHAVE API) ---
st.sidebar.header("🔑 Configuração da API")

default_api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if "api_key" in st.session_state and st.session_state["api_key"]:
    default_api_key = st.session_state["api_key"]

api_key_input = st.sidebar.text_input(
    "Insira sua Google Gemini API Key", 
    value=default_api_key, 
    type="password",
    help="Cole aqui a sua chave do Google AI Studio"
)

if api_key_input:
    st.session_state["api_key"] = api_key_input

st.sidebar.markdown("---")
st.sidebar.markdown("### 📋 Modelos Disponíveis")
modelo_escolhido = st.sidebar.radio(
    "Escolha o formato de saída:",
    ("Modelo Clássico (Sem Foto)", "Modelo Com Foto")
)

foto_arquivo = None
if modelo_escolhido == "Modelo Com Foto":
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🖼️ Foto do Candidato")
    foto_arquivo = st.sidebar.file_uploader("Carregue a foto (JPG ou PNG)", type=["jpg", "jpeg", "png"])

ativa_api_key = st.session_state.get("api_key", "")

# --- ÁREA PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Dados do Candidato e Vaga")
    curriculo_antigo = st.text_area("Cole o Currículo Atual do Cliente:", height=200, placeholder="Cole aqui o texto do currículo antigo...")
    descricao_vaga = st.text_area("Cole a Descrição da Vaga ou Palavras-Chave:", height=150, placeholder="Cole a descrição da vaga...")

with col2:
    st.subheader("2️⃣ Instruções e Execução")
    st.info("O sistema gerará o currículo com fonte Times, cores 100% pretas e os ícones gráficos idênticos aos da referência.")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = ['gemini-3.8-flash', 'gemini-3.7-flash', 'gemini-3.1-pro']
    
    erros_acumulados = []
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
                erros_acumulados.append(f"[{modelo}] {str(e)}")
                time.sleep(2)
                continue
                
    raise Exception(f"Todos os modelos atuais estão ocupados no momento. Detalhes: {erros_acumulados[-1]}")

if gerar_btn:
    if not ativa_api_key:
        st.error("⚠️ Por favor, insira a sua Chave de API na barra lateral para continuar.")
    elif not curriculo_antigo or not descricao_vaga:
        st.warning("⚠️ Preencha tanto o currículo antigo quanto a descrição da vaga.")
    else:
        with st.spinner("A processar e reestruturando o currículo de acordo com o padrão visual oficial..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH seguindo rigorosamente o tutorial de elaboração de currículos da Sala de Emprego.
                Com base no currículo antigo e na descrição da vaga fornecidos, gere o conteúdo limpo do currículo seguindo estritamente este formato de blocos com tags em maiúsculas:
                
                [NOME]
                Nome Completo do Candidato
                
                [CARGO]
                Cargo ou Objetivo Profissional (Máximo de 3 opções separadas por barra)
                
                [CONTATOS]
                Bairro, Cidade | Telefone | E-mail | LinkedIn
                
                [PERFIL PROFISSIONAL]
                (Escrito estritamente em 3ª pessoa e dividido em exatamente 3 parágrafos, sem quebras extras dentro do mesmo parágrafo):
                1º Parágrafo: Profissional atuante há mais de X anos na área [cargo/objetivo], destacando-se pela sua capacidade [duas ou três competências comportamentais importantes para o cargo], garantindo que [benefício dessas competências para a função].
                2º Parágrafo: Expertises em [atividades chaves, competências e palavras-chave que mais se repetem na descrição da vaga].
                3º Parágrafo: Conhecimentos em [sistemas, ferramentas e fundamentos teóricos exigidos].
                
                [FORMAÇÃO ACADÊMICA]
                Nome do curso | Instituição de ensino - Ano de conclusão
                
                [CURSOS E CERTIFICAÇÕES]
                Nome do curso | Instituição de ensino - Ano de conclusão
                
                [HABILIDADES E COMPETÊNCIAS]
                - Competência técnica ou comportamental 1
                - Competência técnica ou comportamental 2
                
                [EXPERIÊNCIA PROFISSIONAL]
                Empresa | Mês/Ano de entrada - Mês/Ano de saída (ou Atual)
                Cargo
                - Atividade neutra e profissional 1 baseada em anúncios de vagas
                - Atividade neutra e profissional 2 baseada em anúncios de vagas
                
                Currículo Antigo:
                {curriculo_antigo}
                
                Descrição da Vaga / Requisitos:
                {descricao_vaga}
                """
                
                resultado_ia = chamar_gemini_com_retry(client, prompt_sistema)
                st.session_state["curriculo_gerado"] = resultado_ia
                st.success("✨ Currículo otimizado com sucesso!")
                
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")

if "curriculo_gerado" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Resultado Gerado")
    st.text_area("Texto Otimizado:", st.session_state["curriculo_gerado"], height=300)
    
    st.info(f"Modo selecionado: **{modelo_escolhido}**. Escolha abaixo o formato de download desejado.")

    texto_gerado = st.session_state["curriculo_gerado"]

    # --- GERAÇÃO DE WORD (.DOCX) ---
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0, 0, 0)

    p_corpo = doc.add_paragraph()
    p_corpo.add_run(texto_gerado)

    buffer_word = io.BytesIO()
    doc.save(buffer_word)
    buffer_word.seek(0)

    # --- GERAÇÃO DE PDF PERSONALIZADO (PADRÃO EXATO COM FONTES TIMES E PRETO SÓLIDO) ---
    buffer_pdf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    cor_total = colors.HexColor("#000000")

    estilo_nome = ParagraphStyle('NomeEstilo', parent=styles['Heading1'], fontSize=15, leading=17, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)
    estilo_cargo = ParagraphStyle('CargoEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Italic", spaceAfter=2)
    estilo_contato_dir = ParagraphStyle('ContatoDirEstilo', parent=styles['Normal'], fontSize=8.5, leading=11.5, textColor=cor_total, fontName="Times-Roman", alignment=2) 
    
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=9.5, leading=12, textColor=cor_total, spaceBefore=6, spaceAfter=2, fontName="Times-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=8.5, leading=11.5, textColor=cor_total, fontName="Times-Roman", spaceAfter=2.5, alignment=4)
    estilo_exp_empresa = ParagraphStyle('EmpresaExpEstilo', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=cor_total, fontName="Times-Italic", spaceAfter=1)
    estilo_exp_cargo = ParagraphStyle('CargoExpEstilo', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)

    story = []

    def extrair_tag(tag, texto):
        match = re.search(rf'\[{tag}\](.*?)(?=\[|$)', texto, re.DOTALL)
        return match.group(1).strip() if match else ""

    nome_txt = extrair_tag("NOME", texto_gerado) or "NOME DO CANDIDATO"
    cargo_txt = extrair_tag("CARGO", texto_gerado) or "Cargo Profissional"
    contato_txt = extrair_tag("CONTATOS", texto_gerado) or "Bairro, Cidade | Telefone | E-mail | LinkedIn"
    perfil_txt = extrair_tag("PERFIL PROFISSIONAL", texto_gerado)
    formacao_txt = extrair_tag("FORMAÇÃO ACADÊMICA", texto_gerado)
    cursos_txt = extrair_tag("CURSOS E CERTIFICAÇÕES", texto_gerado)
    habilidades_txt = extrair_tag("HABILIDADES E COMPETÊNCIAS", texto_gerado)
    experiencia_txt = extrair_tag("EXPERIÊNCIA PROFISSIONAL", texto_gerado)

    partes_contato = [p.strip() for p in contato_txt.replace('|', ',').split(',') if p.strip()]
    cidade_bairro = partes_contato[0] if len(partes_contato) > 0 else "Bairro, Cidade"
    telefone = partes_contato[1] if len(partes_contato) > 1 else "Telefone"
    email = partes_contato[2] if len(partes_contato) > 2 else "E-mail"
    linkedin = partes_contato[3] if len(partes_contato) > 3 else "https://www.linkedin.com/in/"

    def criar_icone_casa():
        d = Drawing(12, 11)
        d.add(Polygon([1, 4, 6, 0, 11, 4], fillColor=colors.black, strokeColor=colors.black))
        d.add(Rect(3, 0, 6, 5, fillColor=colors.black, strokeColor=colors.black))
        return d

    def criar_icone_telefone():
        d = Drawing(12, 11)
        p = Path(fillColor=colors.black, strokeColor=colors.black)
        p.moveTo(1, 9)
        p.curveTo(1, 11, 3, 11, 5, 9)
        p.lineTo(7, 7)
        p.curveTo(8, 6, 8, 4, 6, 2)
        p.lineTo(4, 4)
        p.lineTo(6, 6)
        p.lineTo(5, 7)
        d.add(p)
        return d

    def criar_icone_email():
        d = Drawing(13, 11)
        d.add(Rect(0, 1, 13, 9, fillColor=colors.black, strokeColor=colors.black))
        p = Path(fillColor=colors.white, strokeColor=colors.white, strokeWidth=1)
        p.moveTo(1, 9)
        p.lineTo(6.5, 5)
        p.lineTo(12, 9)
        d.
