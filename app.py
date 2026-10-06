import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from google import genai
import time
import io
import os

st.set_page_config(page_title="Otimizador de Currículos - Consultoria", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Consultoria)")
st.markdown("Ferramenta automatizada para otimização de currículos com layout personalizado e exportação em Word e PDF.")

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
    st.info("O sistema vai reestruturar o perfil em 3 parágrafos, ajustar as experiências e gerar o documento estruturado nos formatos Word e PDF com o layout exato da consultoria.")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = ['gemini-2.5-flash', 'gemini-flash-latest']
    erros_acumulados = []
    for modelo in modelos_para_tentar:
        for tentativa in range(2):
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=prompt_texto,
                )
                if response and response.text:
                    return response.text
            except Exception as e:
                erros_acumulados.append(str(e))
                time.sleep(1)
                continue
    raise Exception(f"Erro ao conectar com a API do Gemini. Detalhes: {erros_acumulados[-1] if erros_acumulados else 'Desconhecido'}")

if gerar_btn:
    if not ativa_api_key:
        st.error("⚠️ Por favor, insira a sua Chave de API na barra lateral para continuar.")
    elif not curriculo_antigo or not descricao_vaga:
        st.warning("⚠️ Preencha tanto o currículo antigo quanto a descrição da vaga.")
    else:
        with st.spinner("A processar e reestruturando o currículo de acordo com o padrão visual da consultoria..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH e Otimização de Currículos ATS.
                Com base no currículo antigo e na descrição da vaga fornecidos, gere o conteúdo estruturado rigorosamente com os seguintes campos e seções exatas:
                
                [NOME]
                [CARGO/OBJETIVO] (Máximo 3 opções)
                [CONTATOS] (Bairro, Cidade | Telefone | E-mail | LinkedIn)
                
                [PERFIL PROFISSIONAL]
                (Escrito estritamente em 3ª pessoa: 1º Parágrafo com tempo de atuação e competências comportamentais; 2º Parágrafo com expertises e palavras-chave da vaga; 3º Parágrafo com conhecimentos em sistemas e ferramentas).
                
                [FORMAÇÃO ACADÊMICA]
                (Ordem de importância/cronológica. Formato: Nome do curso | Instituição de ensino - Ano de conclusão).
                
                [CURSOS E CERTIFICAÇÕES]
                (Ordem alfabética ou cronológica. Formato: Nome do curso | Instituição de ensino | Ano de conclusão).
                
                [HABILIDADES E COMPETÊNCIAS]
                (Listar tópicos com as competências técnicas e comportamentais extraídas da vaga).
                
                [EXPERIÊNCIA PROFISSIONAL]
                (Ordem cronológica da mais recente para a mais antiga. Empresa | Período e Cargo | Tópicos neutros de atividades).
                
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
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(51, 51, 51)

    p_corpo = doc.add_paragraph()
    p_corpo.add_run(texto_gerado)

    buffer_word = io.BytesIO()
    doc.save(buffer_word)
    buffer_word.seek(0)

    # --- GERAÇÃO DE PDF PERSONALIZADO ---
    buffer_pdf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    estilo_nome = ParagraphStyle('NomeEstilo', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=2)
    estilo_cargo = ParagraphStyle('CargoEstilo', parent=styles['Normal'], fontSize=12, leading=16, textColor=colors.HexColor("#4A5568"), alignment=1, spaceAfter=8)
    estilo_contato = ParagraphStyle('ContatoEstilo', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#718096"), alignment=1, spaceAfter=15)
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=11, leading=15, textColor=colors.HexColor("#1A365D"), spaceBefore=10, spaceAfter=4, fontName="Helvetica-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=10, leading=14, textColor=colors.HexColor("#2D3748"), spaceAfter=6)

    story = []
    
    if modelo_escolhido == "Modelo Com Foto" and foto_arquivo is not None:
        temp_foto_path = "temp_foto.png"
        with open(temp_foto_path, "wb") as f:
            f.write(foto_arquivo.getbuffer())
        
        img = RLImage(temp_foto_path, width=70, height=70)
        header_text = Paragraph("<b>NOME DO CANDIDATO</b><br/><font size=10 color='#4A5568'>Cargo Pretendido</font><br/><font size=8 color='#718096'>Mogi das Cruzes - SP | (11) 99999-9999</font>", estilo_texto)
        t_header = Table([[img, header_text]], colWidths=[80, 400])
        t_header.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
        ]))
        story.append(t_header)
        story.append(Spacer(1, 10))
        if os.path.exists(temp_foto_path):
            os.remove(temp_foto_path)
    else:
        story.append(Paragraph("<b>NOME DO CANDIDATO</b>", estilo_nome))
        story.append(Paragraph("Cargo / Objetivo Profissional", estilo_cargo))
        story.append(Paragraph("Bairro, Cidade | Telefone | E-mail | LinkedIn", estilo_contato))

    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>PERFIL PROFISSIONAL</b>", estilo_titulo_secao))
    story.append(Paragraph(texto_gerado.replace('\n', '<br/>'), estilo_texto))

    pdf_doc.build(story)
    buffer_pdf.seek(0)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label=f"📥 Baixar em Word ({modelo_escolhido})",
            data=buffer_word,
            file_name="curriculo_otimizado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    with col_d2:
        st.download_button(
            label=f"📥 Baixar em PDF ({modelo_escolhido})",
            data=buffer_pdf,
            file_name="curriculo_otimizado.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
