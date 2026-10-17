import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.graphics.shapes import Drawing, Rect, Polygon, String
from google import genai
import time
import io
import os
import re
import pypdf

st.set_page_config(page_title="Otimizador de Currículos - Padrão Sala de Emprego", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Sala de Emprego)")
st.markdown("Ferramenta automatizada com suporte a upload, links de vagas, limite de 2 páginas, fontes otimizadas e seções de Idiomas, Projetos e Voluntariado.")

# --- CONFIGURAÇÃO AUTOMÁTICA DA API ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        ativa_api_key = st.secrets["GEMINI_API_KEY"]
    else:
        ativa_api_key = ""
except Exception:
    ativa_api_key = ""

# --- BARRA LATERAL (MODELOS E DOCUMENTAÇÃO) ---
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

# --- CARREGAMENTO DO MANUAL DO COLABORADOR OFICIAL (PDF) ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 📘 Documentação")

pdf_manual_path = "Manual_Otimizador_de_Curriculos_2.pdf"
if os.path.exists(pdf_manual_path):
    with open(pdf_manual_path, "rb") as f:
        manual_bytes = f.read()
    
    st.sidebar.download_button(
        label="📥 Baixar Manual do Colaborador (PDF)",
        data=manual_bytes,
        file_name="Manual_Otimizador_de_Curriculos_2.pdf",
        mime="application/pdf",
        use_container_width=True
    )
else:
    st.sidebar.warning("⚠️ Arquivo 'Manual_Otimizador_de_Curriculos_2.pdf' não encontrado na pasta do projeto.")

# --- ÁREA PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Currículo do Cliente")
    tipo_entrada_cv = st.radio("Como deseja inserir o currículo?", ("Digitar / Colar Texto", "Carregar Arquivo (PDF ou Word)"), horizontal=True)
    
    curriculo_antigo = ""
    if tipo_entrada_cv == "Digitar / Colar Texto":
        curriculo_antigo = st.text_area("Cole o texto do currículo antigo:", height=180, placeholder="Cole aqui...")
    else:
        arquivo_cv = st.file_uploader("Carregar Currículo (PDF ou DOCX)", type=["pdf", "docx"])
        if arquivo_cv is not None:
            if arquivo_cv.name.endswith(".pdf"):
                try:
                    leitor_pdf = pypdf.PdfReader(arquivo_cv)
                    for pagina in leitor_pdf.pages:
                        texto_extraido = pagina.extract_text()
                        if texto_extraido:
                            curriculo_antigo += texto_extraido + "\n"
                    st.success("✅ PDF do currículo carregado com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao ler PDF: {e}")
            elif arquivo_cv.name.endswith(".docx"):
                try:
                    doc_leitor = docx.Document(arquivo_cv)
                    for para in doc_leitor.paragraphs:
                        curriculo_antigo += para.text + "\n"
                    st.success("✅ Word do currículo carregado com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao ler DOCX: {e}")

    st.markdown("---")
    st.subheader("2️⃣ Descrição da Vaga")
    tipo_entrada_vaga = st.radio("Como deseja inserir a vaga?", ("Digitar / Colar Descrição", "Inserir Link da Vaga (Gupy, LinkedIn, etc.)"), horizontal=True)
    
    descricao_vaga = ""
    if tipo_entrada_vaga == "Digitar / Colar Descrição":
        descricao_vaga = st.text_area("Cole os requisitos da vaga:", height=150, placeholder="Cole a descrição da vaga aqui...")
    else:
        link_vaga = st.text_input("Cole o link da vaga (ex: Gupy, LinkedIn):", placeholder="https://www.linkedin.com/jobs/... ou https://empresa.gupy.io/...")
        info_adicional_vaga = st.text_area("Resumo ou Requisitos principais da vaga (opcional para auxiliar a IA):", height=80, placeholder="Ex: Vaga de Assistente de DP, exige Excel avançado...")
        if link_vaga:
            descricao_vaga = f"Link da Vaga: {link_vaga}\nInformações Complementares: {info_adicional_vaga}"

with col2:
    st.subheader("3️⃣ Instruções e Execução")
    st.info("O sistema gerará o currículo com fonte Times otimizada (10pt), cores 100% pretas, ícones perfeitamente sincronizados e máximo de 2 páginas.")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = [
        'gemini-2.0-flash',
        'gemini-3.8-flash',
        'gemini-3.6-flash'
    ]
    
    erros_acumulados = []
    
    for modelo in modelos_para_tentar:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=prompt_texto,
            )
            if response and response.text:
                return response.text
        except Exception as e:
            erro_str = str(e)
            erros_acumulados.append(f"[{modelo}] {erro_str}")
            time.sleep(2)
            continue
                
    raise Exception(f"Todos os modelos testados falharam devido a alta procura ou limites de cota. Detalhes: {erros_acumulados}")

if gerar_btn:
    if not curriculo_antigo.strip() or not descricao_vaga.strip():
        st.warning("⚠️ Por favor, forneça o currículo do cliente e os dados/link da vaga.")
    else:
        with st.spinner("A processar e reestruturando o currículo..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH. A sua tarefa é OTIMIZAR o currículo antigo do candidato com base estritamente nos dados verídicos dele, direcionando-o para a vaga desejada.
                RESTRIÇÃO ABSOLUTA: NÃO invente dados, experiências, cursos ou informações que não constem no currículo original. Mantenha os factos reais e adapte apenas a linguagem para ficar profissional e alinhada à vaga.
                
                FORMATO EXIGIDO PARA O PERFIL PROFISSIONAL (Dividido em 3 parágrafos fluidos):
                - Parágrafo 1: Apresentação profissional (Ex: Profissional graduado/cursando em... com sólida trajetória em...) destacando a experiência geral.
                - Parágrafo 2: Destaques de atuação, hard skills e competências técnicas alinhadas à vaga.
                - Parágrafo 3: Conhecimentos em sistemas, ferramentas e termos/palavras-chave relevantes.
                
                Gere o conteúdo final preenchendo rigorosamente estes blocos com tags em maiúsculas:
                
                [NOME]
                [Nome real do candidato extraído do currículo antigo]
                
                [CARGO]
                [Cargo ou até 3 opções baseadas na vaga e na experiência real, separadas por barra]
                
                [CONTATOS]
                [Bairro, Cidade | Telefone | E-mail | LinkedIn reais extraídos do currículo]
                
                [PERFIL PROFISSIONAL]
                [Parágrafo 1 de introdução e trajetória]
                
                [Parágrafo 2 de competências e destaques]
                
                [Parágrafo 3 de sistemas, ferramentas e palavras-chave]
                
                [FORMAÇÃO ACADÊMICA]
                [Curso real | Instituição real - Ano real]
                
                [CURSOS E CERTIFICAÇÕES]
                [Curso real | Instituição real - Ano real] (Apenas se houver no original)
                
                [IDIOMAS]
                [Idioma - Nível real] (Apenas se houver no original)
                
                [PROJETOS E RESULTADOS]
                [Nome/Descrição de projeto ou resultado real] (Apenas se houver no original)
                
                [TRABALHO VOLUNTÁRIO]
                [Atividade ou projeto voluntário real] (Apenas se houver no original)
                
                [HABILIDADES E COMPETÊNCIAS]
                - [Competência real 1]
                - [Competência real 2]
                
                [EXPERIÊNCIA PROFISSIONAL]
                [Empresa real] | [Período real]
                [Cargo real]
                - [Atividade real reescrita de forma neutra e profissional]
                - [Atividade real reescrita de forma neutra e profissional]
                
                Currículo Antigo do Cliente:
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

    # --- GERAÇÃO DE PDF PERSONALIZADO ---
    buffer_pdf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    cor_total = colors.HexColor("#000000")

    estilo_nome = ParagraphStyle('NomeEstilo', parent=styles['Heading1'], fontSize=15, leading=17, textColor=cor_total, fontName="Times-Bold", spaceAfter=2)
    estilo_cargo = ParagraphStyle('CargoEstilo', parent=styles['Normal'], fontSize=10.5, leading=12.5, textColor=cor_total, fontName="Times-Italic", spaceAfter=2)
    estilo_contato_dir = ParagraphStyle('ContatoDirEstilo', parent=styles['Normal'], fontSize=8.5, leading=10.5, textColor=cor_total, fontName="Times-Roman", alignment=2) 
    
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=10.5, leading=12.5, textColor=cor_total, spaceBefore=6, spaceAfter=3, fontName="Times-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=2.5, alignment=4)
    estilo_perfil_texto = ParagraphStyle('PerfilTextoEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=4, alignment=4)
    estilo_exp_empresa = ParagraphStyle('EmpresaExpEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=1)
    estilo_exp_cargo = ParagraphStyle('CargoExpEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)

    story = []

    def extrair_tag(tag, texto):
        match = re.search(rf'\[{tag}\]\s*(.*?)(?=\[|$)', texto, re.DOTALL)
        return match.group(1).strip() if match else ""

    nome_txt = extrair_tag("NOME", texto_gerado) or "NOME DO CANDIDATO"
    cargo_txt = extrair_tag("CARGO", texto_gerado) or "Cargo Profissional"
    contato_txt = extrair_tag("CONTATOS", texto_gerado) or "Bairro, Cidade | Telefone | E-mail | LinkedIn"
    perfil_txt = extrair_tag("PERFIL PROFISSIONAL", texto_gerado)
    formacao_txt = extrair_tag("FORMAÇÃO ACADÊMICA", texto_gerado)
    cursos_txt = extrair_tag("CURSOS E CERTIFICAÇÕES", texto_gerado)
    idiomas_txt = extrair_tag("IDIOMAS", texto_gerado)
    projetos_txt = extrair_tag("PROJETOS E RESULTADOS", texto_gerado)
    voluntariado_txt = extrair_tag("TRABALHO VOLUNTÁRIO", texto_gerado)
    habilidades_txt = extrair_tag("HABILIDADES E COMPETÊNCIAS", texto_gerado)
    experiencia_txt = extrair_tag("EXPERIÊNCIA PROFISSIONAL", texto_gerado)

    # --- IDENTIFICAÇÃO RIGOROSA DOS CONTATOS VIA REGEX ---
    partes_contato = [p.strip() for p in contato_txt.replace('|', '\n').split('\n') if p.strip()]
    
    email_val = "E-mail"
    telefone_val = "Telefone"
    linkedin_val = "https://www.linkedin.com/in/"
    endereco_val = "Bairro, Cidade"

    for p in partes_contato:
        if "@" in p:
            email_val = p
        elif "linkedin" in p.lower() or "in/" in p.lower():
            linkedin_val = p
        elif re.search(r'\d{8,}', re.sub(r'\D', '', p)):
            telefone_val = p
        else:
            endereco_val = p

    endereco_val_limpo = endereco_val.replace('\n', ' • ')
    
    # Encurta a exibição visual do link do LinkedIn para evitar quebras longas na tabela
    linkedin_exibicao = linkedin_val
    if "linkedin.com/in/" in linkedin_val:
        linkedin_exibicao = linkedin_val.replace("https://", "").replace("www.", "")
        if len(linkedin_exibicao) > 32:
            linkedin_exibicao = linkedin_exibicao[:30] + "..."

    # Funções de Ícones
    def criar_icone_casa():
        d = Drawing(12, 12)
        d.add(Rect(0, 0, 12, 12, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Polygon([2.5, 6.5, 6, 3, 9.5, 6.5], fillColor=colors.white, strokeColor=colors.white))
        d.add(Rect(4, 1.5, 4, 4.5, fillColor=colors.white, strokeColor=colors.white))
        d.add(Rect(5, 1.5, 2, 2.5, fillColor=colors.black, strokeColor=colors.black))
        return d

    def criar_icone_telefone():
        d = Drawing(12, 12)
        d.add(Rect(0, 0, 12, 12, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Polygon([3, 8.5, 4.5, 10, 6, 8.5, 5, 7.5, 7.5, 5, 8.5, 6, 10, 4.5, 8.5, 3, 7.5, 4, 5, 6.5], fillColor=colors.white, strokeColor=colors.white))
        return d

    def criar_icone_email():
        d = Drawing(12, 12)
        d.add(Rect(0, 0, 12, 12, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Rect(1.5, 3, 9, 6, rx=1, ry=1, fillColor=colors.white, strokeColor=colors.white))
        d.add(Polygon([1.8, 8.8, 6, 6, 10.2, 8.8], fillColor=colors.black, strokeColor=colors.black))
        return d

    def criar_icone_linkedin():
        d = Drawing(12, 12)
        d.add(Rect(0, 0, 12, 12, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(String(2, 2.5, "in", fontName="Helvetica-Bold", fontSize=8, fillColor=colors.white))
        return d

    if modelo_escolhido == "Modelo Com Foto" and foto_arquivo is not None:
        temp_foto_path = "temp_foto.png"
        with open(temp_foto_path, "wb") as f:
            f.write(foto_arquivo.getbuffer())
        img = RLImage(temp_foto_path, width=48, height=48)
        
        infos_com_foto = [
            Paragraph(f"<b>{nome_txt}</b>", estilo_nome),
            Paragraph(cargo_txt, estilo_cargo)
        ]
        
        tabela_esquerda_foto = Table([
            [img, infos_com_foto]
        ], colWidths=[55, 275])
        tabela_esquerda_foto.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'LEFT'),
            ('LEFTPADDING', (1,0), (1,0), 5),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        coluna_esquerda = tabela_esquerda_foto
    else:
        coluna_esquerda = [
            Paragraph(f"<b>{nome_txt}</b>", estilo_nome),
            Paragraph(cargo_txt, estilo_cargo)
        ]

    # --- TABELA DE CONTATOS COM LARGURA ESTÁVEL PARA O LINKEDIN ---
    tabela_contatos_direita = Table([
        [Paragraph(endereco_val_limpo, estilo_contato_dir), criar_icone_casa()],
        [Paragraph(telefone_val, estilo_contato_dir), criar_icone_telefone()],
        [Paragraph(email_val, estilo_contato_dir), criar_icone_email()],
        [Paragraph(linkedin_exibicao, estilo_contato_dir), criar_icone_linkedin()]
    ], colWidths=[205, 18])
    
    tabela_contatos_direita.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,-1), 'RIGHT'),
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 0.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.5),
        ('LEFTPADDING', (1,0), (1,-1), 2),
        ('RIGHTPADDING', (0,0), (0,-1), 0),
    ]))

    t_header = Table([ [coluna_esquerda, tabela_contatos_direita] ], colWidths=[317, 223])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(t_header)
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.8, color=cor_total, spaceAfter=4, spaceBefore=0))

    if perfil_txt:
        story.append(Paragraph("<b>Perfil Profissional</b>", estilo_titulo_secao))
        paragrafos_perfil = [p.strip() for p in perfil_txt.split('\n') if p.strip()]
        for p_paragrafo in paragrafos_perfil:
            story.append(Paragraph(p_paragrafo, estilo_perfil_texto))

    if formacao_txt:
        story.append(Paragraph("<b>Formação Acadêmica</b>", estilo_titulo_secao))
        story.append(Paragraph(formacao_txt.replace('\n', '<br/>'), estilo_texto))

    if cursos_txt and "[CURSOS E CERTIFICAÇÕES]" not in cursos_txt:
        story.append(Paragraph("<b>Cursos e Certificações</b>", estilo_titulo_secao))
        story.append(Paragraph(cursos_txt.replace('\n', '<br/>'), estilo_texto))

    if idiomas_txt and "[IDIOMAS]" not in idiomas_txt:
        story.append(Paragraph("<b>Idiomas</b>", estilo_titulo_secao))
        story.append(Paragraph(idiomas_txt.replace('\n', '<br/>'), estilo_texto))

    if projetos_txt and "[PROJETOS E RESULTADOS]" not in projetos_txt:
        story.append(Paragraph("<b>Projetos e Resultados</b>", estilo_titulo_secao))
        story.append(Paragraph(projetos_txt.replace('\n', '<br/>'), estilo_texto))

    if voluntariado_txt and "[TRABALHO VOLUNTÁRIO]" not in voluntariado_txt:
        story.append(Paragraph("<b>Trabalho Voluntário</b>", estilo_titulo_secao))
        story.append(Paragraph(voluntariado_txt.replace('\n', '<br/>'), estilo_texto))

    if habilidades_txt:
        hab_formatadas = habilidades_txt.replace('-', '•').replace('\n', '<br/>')
        story.append(Paragraph("<b>Habilidades e Competências</b>", estilo_titulo_secao))
        story.append(Paragraph(hab_formatadas, estilo_texto))

    if experiencia_txt:
        story.append(Paragraph("<b>Experiência Profissional</b>", estilo_titulo_secao))
        
        blocos_exp = experiencia_txt.split("\n\n")
        for bloco in blocos_exp:
            linhas_bloco = [l.strip() for l in bloco.split("\n") if l.strip()]
            if not linhas_bloco:
                continue
            
            empresa_periodo = linhas_bloco[0]
            
            if "|" in empresa_periodo:
                partes_emp = empresa_periodo.split("|")
                nome_emp = partes_emp[0].strip()
                resto_emp = partes_emp[1].strip()
                empresa_formatada = f"<b><i>{nome_emp}</i></b> | {resto_emp}"
            else:
                empresa_formatada = f"<b><i>{empresa_periodo}</i></b>"
                
            story.append(Paragraph(empresa_formatada, estilo_exp_empresa))
            
            if len(linhas_bloco) > 1:
                cargo_linha = linhas_bloco[1]
                story.append(Paragraph(f"{cargo_linha}", estilo_exp_cargo))
            
            for item in linhas_bloco[2:]:
                item_limpo = item.lstrip('-•* ').strip()
                story.append(Paragraph(f"• {item_limpo}", estilo_texto))
            
            story.append(Spacer(1, 1))

    pdf_doc.build(story)
    buffer_pdf.seek(0)

    if os.path.exists("temp_foto.png"):
        os.remove("temp_foto.png")

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
