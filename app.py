import streamlit as st
from google import genai
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import reportlab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Otimizador de Currículos ATS",
    page_icon="📄",
    layout="wide"
)

st.markdown("""
    <style>
    @media print {
        header, .stSidebar, .stButton, div[data-baseweb="input"], textarea {
            display: none !important;
        }
        .main {
            background-color: white !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

st.title("📄 Otimizador de Currículos Focado em ATS")
st.markdown("Esta ferramenta analisa o seu currículo atual e gera um documento limpo, profissional e otimizado.")

# Barra lateral
with st.sidebar:
    st.header("⚙️ Configurações")
    api_key = st.text_input("Insira sua Google Gemini API Key:", type="password")
    st.markdown("---")
    st.markdown("### 💡 Como usar:")
    st.markdown("1. Insira sua chave da API.")
    st.markdown("2. Cole seu currículo atual e a vaga.")
    st.markdown("3. Clique em Otimizar e descarregue formatado.")

# Layout principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 Seu Currículo Atual")
    curriculum_text = st.text_area("Cole o texto do seu currículo aqui:", height=300)

with col2:
    st.subheader("🎯 Descrição da Vaga")
    job_description = st.text_area("Cole a descrição da vaga desejada aqui:", height=300)

def extract_pure_cv(text):
    filtered_text = text
    if "---" in text:
        parts = text.split("---")
        for part in parts:
            if "RESUMO" in part.upper() or "EXPERIÊNCIA" in part.upper() or "PERFIL" in part.upper():
                filtered_text = part
                break
    return filtered_text.strip()

# Função para criar um Word (.docx) estruturado e elegante
def create_docx(text_content):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    for line in text_content.split('\n'):
        line_stripped = line.strip()
        if not line_stripped:
            continue
            
        if line_stripped.isupper() and len(line_stripped) < 40 and not any(c in line_stripped for c in [':', '-', '|']):
            p = doc.add_paragraph()
            run = p.add_run(line_stripped)
            run.font.name = 'Arial'
            run.font.size = Pt(15)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
        elif line_stripped.startswith('## ') or line_stripped.startswith('### ') or (line_stripped.isupper() and len(line_stripped) < 30):
            clean_sec = line_stripped.replace('## ', '').replace('### ', '')
            p = doc.add_paragraph()
            run = p.add_run(clean_sec)
            run.font.name = 'Arial'
            run.font.size = Pt(11.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
        elif line_stripped.startswith('- ') or line_stripped.startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            run = p.add_run(line_stripped[2:].replace('**', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            p.paragraph_format.space_after = Pt(2)
        else:
            p = doc.add_paragraph()
            run = p.add_run(line_stripped.replace('**', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(51, 51, 51)
            p.paragraph_format.space_after = Pt(3)
            if "@" in line_stripped or "|" in line_stripped:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run.font.color.rgb = RGBColor(100, 100, 100)
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Função para criar um PDF elegante e corporativo
def create_pdf(text_content):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=54, leftMargin=54, topMargin=54, bottomMargin=54)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CVTitle', parent=styles['Normal'], fontSize=14, leading=16,
        textColor=reportlab.lib.colors.HexColor('#003366'), alignment=1, spaceAfter=4
    )
    contact_style = ParagraphStyle(
        'CVContact', parent=styles['Normal'], fontSize=9, leading=11,
        textColor=reportlab.lib.colors.HexColor('#666666'), alignment=1, spaceAfter=8
    )
    heading_style = ParagraphStyle(
        'CVHeading', parent=styles['Normal'], fontSize=11, leading=13,
        textColor=reportlab.lib.colors.HexColor('#003366'), spaceBefore=8, spaceAfter=3
    )
    normal_style = ParagraphStyle(
        'CVNormal', parent=styles['Normal'], fontSize=9.5, leading=13,
        textColor=reportlab.lib.colors.HexColor('#333333'), spaceAfter=3
    )

    story = []
    for line in text_content.split('\n'):
        line_stripped = line.strip()
        if not line_stripped:
            story.append(Spacer(1, 3))
            continue
            
        clean_line = line_stripped.replace('**', '').replace('### ', '').replace('## ', '').replace('# ', '')
        
        if line_stripped.isupper() and len(line_stripped) < 40 and not any(c in line_stripped for c in [':', '-', '|']):
            story.append(Paragraph(f"<b>{clean_line}</b>", title_style))
        elif "@" in line_stripped or "|" in line_stripped:
            story.append(Paragraph(clean_line, contact_style))
            story.append(HRFlowable(width="100%", thickness=0.5, color=reportlab.lib.colors.HexColor('#CCCCCC'), spaceBefore=1, spaceAfter=6))
        elif line_stripped.startswith('#') or (line_stripped.isupper() and len(line_stripped) < 30):
            story.append(Paragraph(f"<b>{clean_line}</b>", heading_style))
        elif line_stripped.startswith('- ') or line_stripped.startswith('* '):
            story.append(Paragraph(f"• {clean_line[2:]}", normal_style))
        else:
            story.append(Paragraph(clean_line, normal_style))
            
    doc.build(story)
    bio.seek(0)
    return bio

# Botão de otimização
if st.button("🚀 Otimizar Currículo para esta Vaga", type="primary"):
    if not api_key:
        st.error("Por favor, insira a sua Google Gemini API Key na barra lateral.")
    elif not curriculum_text or not job_description:
        st.warning("Por favor, preencha tanto o currículo quanto a descrição da vaga.")
    else:
        try:
            client = genai.Client(api_key=api_key)
            
            with st.spinner("A analisar o perfil e a otimizar para o ATS..."):
                system_instruction = (
                   "Você é um especialista em recrutamento e seleção. Sua tarefa é reescrever o currículo do candidato para a vaga fornecida. "
                    "ATENÇÃO: Você DEVE gerar o currículo COMPLETO seguindo estritamente esta estrutura abaixo, preenchendo todas as seções com as informações reais do candidato:\n\n"
                    "[NOME COMPLETO]\n"
                    "[Cidade - UF | Telefone | E-mail]\n\n"
                    "## OBJETIVO\n"
                    "[Cargo desejado]\n\n"
                    "## RESUMO PROFISSIONAL\n"
                    "[Um parágrafo coeso, elegante e fluído destacando a trajetória e competências, sem usar listas]\n\n"
                    "## EXPERIÊNCIA PROFISSIONAL\n"
                    "[Empresa] - [Cargo] | [Período]\n"
                    "- [Atribuição 1 com verbo de ação]\n"
                    "- [Atribuição 2 com verbo de ação]\n\n"
                    "## FORMAÇÃO ACADEMICA\n"
                    "[Curso e Instituição | Previsão]\n\n"
                    "## CURSOS E QUALIFICAÇÕES\n"
                    "- [Curso 1]\n"
                    "- [Curso 2]\n\n"
                    "## HABILIDADES E COMPETÊNCIAS\n"
                    "- [Competência 1]\n"
                    "- [Competência 2]\n\n"
                    "REGRAS: Não invente dados e não adicione conversas ou introduções."
                
                )
                
                user_message = f"CURRÍCULUM ATUAL:\n{curriculum_text}\n\nDESCRIÇÃO DA VAGA:\n{job_description}"

                response = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=user_message,
                    config={
                        'system_instruction': system_instruction,
                        'temperature': 0.2,
                    }
                )
                
                pure_cv = extract_pure_cv(response.text)
                st.session_state['optimized_cv'] = pure_cv
                
        except Exception as e:
            st.error(f"Ocorreu um erro durante o processamento: {e}")

# Mostrar resultado e botões de download
if 'optimized_cv' in st.session_state:
    st.success("Currículo otimizado com sucesso!")
    st.markdown("### 📊 Currículo Pronto para Envio:")
    st.markdown(st.session_state['optimized_cv'])
    
    st.markdown("---")
    
    col_dl1, col_dl2 = st.columns(2)
    
    docx_file = create_docx(st.session_state['optimized_cv'])
    pdf_file = create_pdf(st.session_state['optimized_cv'])
    
    with col_dl1:
        st.download_button(
            label="📥 Descarregar em Word (.docx)",
            data=docx_file,
            file_name="Curriculo_Otimizado_ATS.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
    with col_dl2:
        st.download_button(
            label="📥 Descarregar em PDF (.pdf)",
            data=pdf_file,
            file_name="Curriculo_Otimizado_ATS.pdf",
            mime="application/pdf",
            type="primary"
        )