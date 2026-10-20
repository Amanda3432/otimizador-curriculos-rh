import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from google import genai
import time
import io
import os
import re
import pypdf
import urllib.request
import urllib.parse
from html.parser import HTMLParser

st.set_page_config(page_title="Otimizador de Currículos - Padrão Sala de Emprego", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Sala de Emprego)")
st.markdown("Ferramenta automatizada com leitura avançada de links de vagas, CNH/CREA restritas ao cabeçalho, linha divisória no Word e limite de 2 páginas.")

# --- CONFIGURAÇÃO AUTOMÁTICA DA API ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        ativa_api_key = st.secrets["GEMINI_API_KEY"]
    else:
        ativa_api_key = ""
except Exception:
    ativa_api_key = ""

# --- FUNÇÃO PARA EXTRAIR TEXTO DE LINKS DE VAGAS (GUPY, LINKEDIN, ETC.) ---
class HTMLTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.text_accumulator = []

    def handle_data(self, data):
        text = data.strip()
        if text:
            self.text_accumulator.append(text)

    def get_text(self):
        return ' '.join(self.text_accumulator)

def extrair_texto_de_url(url):
    try:
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=10) as response:
            html_bytes = response.read()
            html_content = html_bytes.decode('utf-8', errors='ignore')
            parser = HTMLTextExtractor()
            parser.feed(html_content)
            texto_cru = parser.get_text()
            return texto_cru[:4000]
    except Exception as e:
        return f"[Aviso: Não foi possível realizar o scraping automático da URL ({e}). A IA usará as informações complementares informadas]."

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
        info_adicional_vaga = st.text_area("Resumo ou Requisitos principais da vaga (opcional para auxiliar a IA):", height=80, placeholder="Ex: Vaga de Analista, exige Excel avançado...")
        if link_vaga:
            with st.spinner("Lendo o conteúdo do link da vaga..."):
                conteudo_url = extrair_texto_de_url(link_vaga)
                descricao_vaga = f"Link da Vaga: {link_vaga}\nConteúdo Extraído da Vaga:\n{conteudo_url}\nInformações Complementares: {info_adicional_vaga}"

with col2:
    st.subheader("3️⃣ Instruções e Execução")
    st.info("O sistema fará a leitura do descritivo da vaga, mantendo CNH/CREA exclusivamente no cabeçalho e eliminando qualquer duplicidade.")
    
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
        with st.spinner("Lendo requisitos da vaga e otimizando o currículo..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH e sistemas ATS. A sua tarefa é OTIMIZAR o currículo antigo do candidato com base estritamente nos dados verídicos dele, direcionando-o para a vaga descrita abaixo.
                RESTRIÇÃO ABSOLUTA: NÃO invente dados, experiências, cursos ou informações que não constem no currículo original. Mantenha os factos reais e adapte a linguagem técnica e as palavras-chave para atender rigorosamente aos requisitos do descritivo da vaga.
                
                REGRAS DE NÃO-REPETIÇÃO E ESTRUTURA:
                - No [PERFIL PROFISSIONAL], faça uma síntese executiva em 3 parágrafos focada em trajetória e competências alinhadas à vaga. NUNCA cite detalhes específicos de Idiomas, Cursos, CNH ou CREA no perfil, pois eles têm locais próprios.
                - Na tag [CONTATOS], coloque obrigatoriamente: [Bairro, Cidade | Telefone | E-mail | LinkedIn | CNH (se houver) | CREA (se houver)]. NUNCA repita CNH ou CREA em nenhuma outra parte do currículo (como habilidades ou perfil).
                
                Gere o conteúdo final preenchendo rigorosamente estes blocos com tags em maiúsculas:
                
                [NOME]
                [Nome real do candidato extraído do currículo antigo]
                
                [CARGO]
                [Cargo ou até 3 opções baseadas na vaga e na experiência real, separadas por barra]
                
                [CONTATOS]
                [Bairro, Cidade | Telefone | E-mail | LinkedIn | CNH ativa (se houver) | CREA ativo (se houver)]
                
                [PERFIL PROFISSIONAL]
                [Parágrafo 1 de introdução e trajetória geral]
                
                [Parágrafo 2 de competências e destaques corporativos alinhados à vaga]
                
                [Parágrafo 3 de sistemas, ferramentas e palavras-chave exigidas no descritivo]
                
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
                - [Competência real 1 alinhada à vaga]
                - [Competência real 2 alinhada à vaga]
                
                [EXPERIÊNCIA PROFISSIONAL]
                [Empresa real] | [Período real]
                [Cargo real]
                - [Atividade real reescrita com foco nos requisitos da vaga]
                - [Atividade real reescrita com foco nos requisitos da vaga]
                
                Currículo Antigo do Cliente:
                {curriculo_antigo}
                
                Descritivo da Vaga (Extraído do Link / Informações):
                {descricao_vaga}
                """
                
                resultado_ia = chamar_gemini_com_retry(client, prompt_sistema)
                st.session_state["curriculo_gerado"] = resultado_ia
                st.success("✨ Currículo lido e otimizado com sucesso com base na vaga!")
                
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")

if "curriculo_gerado" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Resultado Gerado")
    st.text_area("Texto Otimizado:", st.session_state["curriculo_gerado"], height=300)
    
    st.info(f"Modo selecionado: **{modelo_escolhido}**. Escolha abaixo o formato de download desejado.")

    texto_gerado = st.session_state["curriculo_gerado"]

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

    contatos_linha_limpa = " • ".join([p.strip() for p in contato_txt.replace('|', '•').split('•') if p.strip()])

    # --- GERAÇÃO DE WORD (.DOCX) COM LINHA DIVISÓRIA E PADRÃO UNIFICADO ---
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    def add_p_word(document, text="", bold_prefix="", space_after=2, font_size=10, is_bold=False, is_italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
        p = document.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)

        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(font_size)
            run_b.font.bold = True
            run_b.font.color.rgb = RGBColor(0, 0, 0)

        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            run.font.italic = is_italic
            run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_titulo_word(document, titulo):
        add_p_word(document, text=titulo, font_size=10.5, is_bold=True, space_after=3)

    add_p_word(doc, text=nome_txt, font_size=15, is_bold=True, space_after=1)
    add_p_word(doc, text=cargo_txt, font_size=10.5, is_italic=True, space_after=2)
    add_p_word(doc, text=contatos_linha_limpa, font_size=9, space_after=4)

    p_linha = doc.add_paragraph()
    p_linha.paragraph_format.space_after = Pt(6)
    p_linha_run = p_linha.add_run("―" * 65)
    p_linha_run.font.name = 'Times New Roman'
    p_linha_run.font.size = Pt(9)
    p_linha_run.font.bold = True
    p_linha_run.font.color.rgb = RGBColor(0, 0, 0)

    if perfil_txt:
        add_titulo_word(doc, "Perfil Profissional")
        for p_line in perfil_txt.split('\n'):
            if p_line.strip():
                add_p_word(doc, text=p_line.strip(), space_after=4, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    if formacao_txt:
        add_titulo_word(doc, "Formação Acadêmica")
        for f_line in formacao_txt.split('\n'):
            if f_line.strip():
                add_p_word(doc, text=f_line.strip(), space_after=2)

    if cursos_txt and "[CURSOS E CERTIFICAÇÕES]" not in cursos_txt:
        add_titulo_word(doc, "Cursos e Certificações")
        for c_line in cursos_txt.split('\n'):
            if c_line.strip():
                add_p_word(doc, text=c_line.strip(), space_after=2)

    if idiomas_txt and "[IDIOMAS]" not in idiomas_txt:
        add_titulo_word(doc, "Idiomas")
        for i_line in idiomas_txt.split('\n'):
            if i_line.strip():
                add_p_word(doc, text=i_line.strip(), space_after=2)

    if projetos_txt and "[PROJETOS E RESULTADOS]" not in projetos_txt:
        add_titulo_word(doc, "Projetos e Resultados")
        for pr_line in projetos_txt.split('\n'):
            if pr_line.strip():
                add_p_word(doc, text=pr_line.strip(), space_after=2)

    if voluntariado_txt and "[TRABALHO VOLUNTÁRIO]" not in voluntariado_txt:
        add_titulo_word(doc, "Trabalho Voluntário")
        for v_line in voluntariado_txt.split('\n'):
            if v_line.strip():
                add_p_word(doc, text=v_line.strip(), space_after=2)

    if habilidades_txt:
        add_titulo_word(doc, "Habilidades e Competências")
        for h_line in habilidades_txt.split('\n'):
            if h_line.strip():
                item_limpo = h_line.lstrip('-•* ').strip()
                add_p_word(doc, text=f"• {item_limpo}", space_after=2)

    if experiencia_txt:
        add_titulo_word(doc, "Experiência Profissional")
        blocos_exp = experiencia_txt.split("\n\n")
        for bloco in blocos_exp:
            linhas_bloco = [l.strip() for l in bloco.split("\n") if l.strip()]
            if not linhas_bloco:
                continue
            empresa_periodo = linhas_bloco[0]
            add_p_word(doc, text=empresa_periodo, space_after=1)
            if len(linhas_bloco) > 1:
                add_p_word(doc, text=linhas_bloco[1], is_bold=True, space_after=1)
            for item in linhas_bloco[2:]:
                item_limpo = item.lstrip('-•* ').strip()
                add_p_word(doc, text=f"• {item_limpo}", space_after=2)

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
    estilo_contato_centro = ParagraphStyle('ContatoCentroEstilo', parent=styles['Normal'], fontSize=9, leading=11, textColor=cor_total, fontName="Times-Roman", alignment=0, spaceAfter=4)
    
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=10.5, leading=12.5, textColor=cor_total, spaceBefore=6, spaceAfter=3, fontName="Times-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=2.5, alignment=4)
    estilo_perfil_texto = ParagraphStyle('PerfilTextoEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=4, alignment=4)
    estilo_exp_empresa = ParagraphStyle('EmpresaExpEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Roman", spaceAfter=1)
    estilo_exp_cargo = ParagraphStyle('CargoExpEstilo', parent=styles['Normal'], fontSize=10, leading=13, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)

    story = []

    if modelo_escolhido == "Modelo Com Foto" and foto_arquivo is not None:
        temp_foto_path = "temp_foto.png"
        with open(temp_foto_path, "wb") as f:
            f.write(foto_arquivo.getbuffer())
        img = RLImage(temp_foto_path, width=48, height=48)
        
        infos_com_foto = [
            Paragraph(f"<b>{nome_txt}</b>", estilo_nome),
            Paragraph(cargo_txt, estilo_cargo),
            Paragraph(contatos_linha_limpa, estilo_contato_centro)
        ]
        
        tabela_esquerda_foto = Table([
            [img, infos_com_foto]
        ], colWidths=[55, 485])
        tabela_esquerda_foto.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'LEFT'),
            ('LEFTPADDING', (1,0), (1,0), 6),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        story.append(tabela_esquerda_foto)
    else:
        story.append(Paragraph(f"<b>{nome_txt}</b>", estilo_nome))
        story.append(Paragraph(cargo_txt, estilo_cargo))
        story.append(Paragraph(contatos_linha_limpa, estilo_contato_centro))

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
