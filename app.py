import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
st.markdown("Ferramenta automatizada com modelos Gemini 3.8 / 3.7 / 3.6, linha única para contatos/LinkedIn, sem linha dupla e voluntariado após a experiência.")

# --- CONFIGURAÇÃO AUTOMÁTICA DA API ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        ativa_api_key = st.secrets["GEMINI_API_KEY"]
    else:
        ativa_api_key = ""
except Exception:
    ativa_api_key = ""

# --- FUNÇÃO PARA EXTRAIR TEXTO DE LINKS DE VAGAS ---
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
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            }
        )
        with urllib.request.urlopen(req, timeout=12) as response:
            html_bytes = response.read()
            html_content = html_bytes.decode('utf-8', errors='ignore')
            parser = HTMLTextExtractor()
            parser.feed(html_content)
            texto_cru = parser.get_text()
            return texto_cru[:6000]
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
            with st.spinner("Lendo o conteúdo completo do link da vaga..."):
                conteudo_url = extrair_texto_de_url(link_vaga)
                descricao_vaga = f"Link da Vaga: {link_vaga}\nConteúdo Extraído da Vaga:\n{conteudo_url}\nInformações Complementares: {info_adicional_vaga}"

with col2:
    st.subheader("3️⃣ Instruções e Execução")
    st.info("O sistema utilizará os modelos Gemini 3.8 / 3.7 / 3.6 para gerar currículos perfeitamente formatados sem duplicações de linhas.")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

# --- CHAMADA À API CONFIGURADA COM GEMINI 3.8, 3.7 E 3.6 ---
def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = [
        'gemini-3.8-flash',
        'gemini-3.7-flash',
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
                
    raise Exception(f"Todos os modelos testados (3.8, 3.7, 3.6) falharam. Detalhes: {erros_acumulados}")

if gerar_btn:
    if not curriculo_antigo.strip() or not descricao_vaga.strip():
        st.warning("⚠️ Por favor, forneça o currículo do cliente e os dados/link da vaga.")
    else:
        with st.spinner("Lendo descritivo da vaga e otimizando o currículo no padrão clássico..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH e sistemas ATS da Sala de Emprego. A sua tarefa é OTIMIZAR o currículo antigo do candidato com base estritamente nos dados verídicos dele, direcionando-o de forma cirúrgica para os requisitos da vaga descrita abaixo.
                RESTRIÇÃO ABSOLUTA: NÃO invente dados, experiências, cursos ou informações que não constem no currículo original. Mantenha os factos reais e adapte a linguagem técnica para atender exatamente ao descritivo da vaga fornecido.
                
                PADRÃO EXIGIDO PARA O PERFIL PROFISSIONAL:
                - Redija o perfil profissional no mesmo estilo analítico e detalhista padrão da consultoria, dividido em parágrafos fluidos (apresentação/trajetória, expertises técnicas e sistemas/ferramentas).
                - NUNCA cite detalhes específicos de Idiomas, Cursos, CNH, CREA ou CRT no perfil profissional.
                
                REGRAS DE ESTRUTURA:
                - Na tag [CARGO], forneça o cargo principal ou até 2 opções alinhadas e fortes, evitando textos muito longos para não quebrar de forma isolada.
                - Na tag [CONTATOS], coloque cada informação separada por quebra de linha exata para o layout clássico: 
                  [Bairro, Cidade - UF]
                  [Telefone]
                  [E-mail]
                  [LinkedIn]
                  [CNH / CREA / CRT se houver no original]
                
                Gere o conteúdo final preenchendo rigorosamente estes blocos com tags em maiúsculas:
                
                [NOME]
                [Nome real do candidato extraído do currículo antigo]
                
                [CARGO]
                [Cargo ou até 2 opções fortes baseadas na vaga e na experiência real, separadas por barra]
                
                [CONTATOS]
                Bairro, Cidade - UF
                Telefone
                E-mail
                LinkedIn
                (CNH, CREA ou CRT se houver)
                
                [PERFIL PROFISSIONAL]
                [Parágrafo 1 de introdução e trajetória profissional analítica]
                
                [Parágrafo 2 de expertises e atuações técnicas alinhadas à vaga]
                
                [Parágrafo 3 de conhecimentos em sistemas, ferramentas e indicadores]
                
                [FORMAÇÃO ACADÊMICA]
                [Curso real | Instituição real - Ano real]
                
                [CURSOS E CERTIFICAÇÕES]
                [Curso real | Instituição real - Ano real] (Apenas se houver no original)
                
                [IDIOMAS]
                [Idioma - Nível real] (Apenas se houver no original)
                
                [PROJETOS E RESULTADOS]
                [Nome/Descrição de projeto ou resultado real com destaque de mérito] (Apenas se houver no original)
                
                [HABILIDADES E COMPETÊNCIAS]
                - [Competência real 1 alinhada à vaga]
                - [Competência real 2 alinhada à vaga]
                
                [EXPERIÊNCIA PROFISSIONAL]
                [Empresa real] | [Período real]
                [Cargo real]
                - [Atividade real reescrita com foco nos requisitos da vaga]
                - [Atividade real reescrita com foco nos requisitos da vaga]
                
                [TRABALHO VOLUNTÁRIO]
                [Atividade ou projeto voluntário real] (Apenas se houver no original - OBRIGATORIAMENTE APÓS A EXPERIÊNCIA PROFISSIONAL)
                
                Currículo Antigo do Cliente:
                {curriculo_antigo}
                
                Descritivo da Vaga (Extraído do Link / Informações):
                {descricao_vaga}
                """
                
                resultado_ia = chamar_gemini_com_retry(client, prompt_sistema)
                st.session_state["curriculo_gerado"] = resultado_ia
                st.success("✨ Currículo otimizado com sucesso no modelo clássico!")
                
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
    contatos_raw = extrair_tag("CONTATOS", texto_gerado) or "Bairro, Cidade\nTelefone\nE-mail\nLinkedIn"
    perfil_txt = extrair_tag("PERFIL PROFISSIONAL", texto_gerado)
    formacao_txt = extrair_tag("FORMAÇÃO ACADÊMICA", texto_gerado)
    cursos_txt = extrair_tag("CURSOS E CERTIFICAÇÕES", texto_gerado)
    idiomas_txt = extrair_tag("IDIOMAS", texto_gerado)
    projetos_txt = extrair_tag("PROJETOS E RESULTADOS", texto_gerado)
    habilidades_txt = extrair_tag("HABILIDADES E COMPETÊNCIAS", texto_gerado)
    experiencia_txt = extrair_tag("EXPERIÊNCIA PROFISSIONAL", texto_gerado)
    voluntariado_txt = extrair_tag("TRABALHO VOLUNTÁRIO", texto_gerado)

    # Processamento e limpeza estrita de contatos/LinkedIn para linha única
    linhas_contato_raw = [c.strip() for c in contatos_raw.split('\n') if c.strip()]
    linhas_contato = []
    for line in linhas_contato_raw:
        if "linkedin.com" in line.lower():
            clean_link = re.sub(r'https?://(www\.)?', '', line)
            linhas_contato.append(clean_link)
        else:
            linhas_contato.append(line)

    # --- GERAÇÃO DE WORD (.DOCX) ---
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    def add_p_word(document, text="", bold_prefix="", space_after=2, font_size=11, is_bold=False, is_italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
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

    def add_secao_com_linha_word(document, titulo):
        p_sec = document.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(8)
        p_sec.paragraph_format.space_after = Pt(3)
        
        # Borda superior sólida única acima do título da seção
        pPr = p_sec._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '12')
        top.set(qn('w:space'), '4')
        top.set(qn('w:color'), '000000')
        pBdr.append(top)
        pPr.append(pBdr)

        r_tit = p_sec.add_run(titulo)
        r_tit.font.name = 'Times New Roman'
        r_tit.font.size = Pt(12)
        r_tit.font.bold = True
        r_tit.font.color.rgb = RGBColor(0, 0, 0)

    # Tabela de Cabeçalho no Word (Larguras calibradas para o LinkedIn não quebrar)
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_esq = table_header.rows[0].cells[0]
    cell_dir = table_header.rows[0].cells[1]
    cell_esq.width = Inches(2.2)
    cell_dir.width = Inches(4.8)

    p_esq_nome = cell_esq.paragraphs[0]
    p_esq_nome.paragraph_format.space_after = Pt(1)
    r_nome = p_esq_nome.add_run(nome_txt)
    r_nome.font.name = 'Times New Roman'
    r_nome.font.size = Pt(15)
    r_nome.font.bold = True
    r_nome.font.color.rgb = RGBColor(0, 0, 0)

    p_esq_cargo = cell_esq.add_paragraph()
    p_esq_cargo.paragraph_format.space_after = Pt(0)
    r_cargo = p_esq_cargo.add_run(cargo_txt)
    r_cargo.font.name = 'Times New Roman'
    r_cargo.font.size = Pt(11) # Ajustado para 11pt para evitar quebra isolada da última palavra
    r_cargo.font.italic = True
    r_cargo.font.color.rgb = RGBColor(0, 0, 0)

    for idx, c_item in enumerate(linhas_contato):
        p_dir = cell_dir.paragraphs[0] if idx == 0 else cell_dir.add_paragraph()
        p_dir.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_dir.paragraph_format.space_after = Pt(1)
        r_c = p_dir.add_run(c_item)
        r_c.font.name = 'Times New Roman'
        r_c.font.size = Pt(9.5) if len(c_item) > 35 else Pt(10.5)
        r_c.font.color.rgb = RGBColor(0, 0, 0)

    # Borda inferior SÓLIDA preta no cabeçalho (sem borda superior extra no perfil para evitar duplicidade)
    tcPr = table_header.rows[0].cells[0]._tc.get_or_add_tcPr()
    tcPr2 = table_header.rows[0].cells[1]._tc.get_or_add_tcPr()
    
    for cell_tc in [tcPr, tcPr2]:
        tcBorders = OxmlElement('w:tcBorders')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '4')
        bottom.set(qn('w:color'), '000000')
        tcBorders.append(bottom)
        for b_name in ['top', 'left', 'right']:
            b = OxmlElement(f'w:{b_name}')
            b.set(qn('w:val'), 'none')
            tcBorders.append(b)
        cell_tc.append(tcBorders)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    if perfil_txt:
        add_secao_com_linha_word(doc, "Perfil Profissional")
        for p_line in perfil_txt.split('\n'):
            if p_line.strip():
                add_p_word(doc, text=p_line.strip(), font_size=11, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    if formacao_txt:
        add_secao_com_linha_word(doc, "Formação Acadêmica")
        for f_line in formacao_txt.split('\n'):
            if f_line.strip():
                add_p_word(doc, text=f_line.strip(), font_size=11, space_after=2)

    if cursos_txt and "[CURSOS E CERTIFICAÇÕES]" not in cursos_txt:
        add_secao_com_linha_word(doc, "Cursos e Certificações")
        for c_line in cursos_txt.split('\n'):
            if c_line.strip():
                add_p_word(doc, text=c_line.strip(), font_size=11, space_after=2)

    if idiomas_txt and "[IDIOMAS]" not in idiomas_txt:
        add_secao_com_linha_word(doc, "Idiomas")
        for i_line in idiomas_txt.split('\n'):
            if i_line.strip():
                add_p_word(doc, text=i_line.strip(), font_size=11, space_after=2)

    if projetos_txt and "[PROJETOS E RESULTADOS]" not in projetos_txt:
        add_secao_com_linha_word(doc, "Projetos e Resultados")
        for pr_line in projetos_txt.split('\n'):
            if pr_line.strip():
                add_p_word(doc, text=pr_line.strip(), font_size=11, space_after=2)

    if habilidades_txt:
        add_secao_com_linha_word(doc, "Habilidades e Competências")
        for h_line in habilidades_txt.split('\n'):
            if h_line.strip():
                item_limpo = h_line.lstrip('-•* ').strip()
                add_p_word(doc, text=f"• {item_limpo}", font_size=11, space_after=2)

    if experiencia_txt:
        add_secao_com_linha_word(doc, "Experiência Profissional")
        blocos_exp = experiencia_txt.split("\n\n")
        for bloco in blocos_exp:
            linhas_bloco = [l.strip() for l in bloco.split("\n") if l.strip()]
            if not linhas_bloco:
                continue
            empresa_periodo = linhas_bloco[0]
            add_p_word(doc, text=empresa_periodo, font_size=11, space_after=1)
            if len(linhas_bloco) > 1:
                add_p_word(doc, text=linhas_bloco[1], font_size=11, is_bold=True, space_after=1)
            for item in linhas_bloco[2:]:
                item_limpo = item.lstrip('-•* ').strip()
                add_p_word(doc, text=f"• {item_limpo}", font_size=11, space_after=2)

    # Trabalho Voluntário rigorosamente após a Experiência Profissional
    if voluntariado_txt and "[TRABALHO VOLUNTÁRIO]" not in voluntariado_txt:
        add_secao_com_linha_word(doc, "Trabalho Voluntário")
        for v_line in voluntariado_txt.split('\n'):
            if v_line.strip():
                add_p_word(doc, text=v_line.strip(), font_size=11, space_after=2)

    buffer_word = io.BytesIO()
    doc.save(buffer_word)
    buffer_word.seek(0)

    # --- GERAÇÃO DE PDF PERSONALIZADO ---
    buffer_pdf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    cor_total = colors.HexColor("#000000")

    estilo_nome = ParagraphStyle('NomeEstilo', parent=styles['Heading1'], fontSize=15, leading=17, textColor=cor_total, fontName="Times-Bold", spaceAfter=2)
    estilo_cargo = ParagraphStyle('CargoEstilo', parent=styles['Normal'], fontSize=11, leading=13, textColor=cor_total, fontName="Times-Italic", spaceAfter=0)
    
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=12, leading=14, textColor=cor_total, spaceBefore=4, spaceAfter=3, fontName="Times-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=11, leading=14, textColor=cor_total, fontName="Times-Roman", spaceAfter=2, alignment=4)
    estilo_perfil_texto = ParagraphStyle('PerfilTextoEstilo', parent=styles['Normal'], fontSize=11, leading=14, textColor=cor_total, fontName="Times-Roman", spaceAfter=2, alignment=4)
    estilo_exp_empresa = ParagraphStyle('EmpresaExpEstilo', parent=styles['Normal'], fontSize=11, leading=14, textColor=cor_total, fontName="Times-Roman", spaceAfter=1)
    estilo_exp_cargo = ParagraphStyle('CargoExpEstilo', parent=styles['Normal'], fontSize=11, leading=14, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)

    story = []

    if modelo_escolhido == "Modelo Com Foto" and foto_arquivo is not None:
        temp_foto_path = "temp_foto.png"
        with open(temp_foto_path, "wb") as f:
            f.write(foto_arquivo.getbuffer())
        img = RLImage(temp_foto_path, width=48, height=48)
        
        coluna_esquerda_conteudo = [img, Paragraph(f"<b>{nome_txt}</b>", estilo_nome), Paragraph(cargo_txt, estilo_cargo)]
        ancho_esq = 55
        ancho_dir = 485
    else:
        coluna_esquerda_conteudo = [Paragraph(f"<b>{nome_txt}</b>", estilo_nome), Paragraph(cargo_txt, estilo_cargo)]
        ancho_esq = 210
        ancho_dir = 330

    parágrafos_contatos = []
    for c in linhas_contato:
        f_sz = 9.5 if len(c) > 35 else 10.5
        e_cont = ParagraphStyle('ContatoDirEstiloPDF', parent=styles['Normal'], fontSize=f_sz, leading=f_sz+2, textColor=cor_total, fontName="Times-Roman", alignment=2, spaceAfter=1)
        parágrafos_contatos.append(Paragraph(c, e_cont))

    tabela_cabecalho = Table([
        [coluna_esquerda_conteudo, parágrafos_contatos]
    ], colWidths=[ancho_esq, ancho_dir])
    
    tabela_cabecalho.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))

    story.append(tabela_cabecalho)

    def add_secao_com_linha_pdf(titulo):
        # Linha preta única e sólida acima de cada seção no PDF (sem duplicidade)
        story.append(HRFlowable(width="100%", thickness=0.8, color=cor_total, spaceAfter=2, spaceBefore=6))
        story.append(Paragraph(f"<b>{titulo}</b>", estilo_titulo_secao))

    if perfil_txt:
        add_secao_com_linha_pdf("Perfil Profissional")
        paragrafos_perfil = [p.strip() for p in perfil_txt.split('\n') if p.strip()]
        for p_paragrafo in paragrafos_perfil:
            story.append(Paragraph(p_paragrafo, estilo_perfil_texto))

    if formacao_txt:
        add_secao_com_linha_pdf("Formação Acadêmica")
        story.append(Paragraph(formacao_txt.replace('\n', '<br/>'), estilo_texto))

    if cursos_txt and "[CURSOS E CERTIFICAÇÕES]" not in cursos_txt:
        add_secao_com_linha_pdf("Cursos e Certificações")
        story.append(Paragraph(cursos_txt.replace('\n', '<br/>'), estilo_texto))

    if idiomas_txt and "[IDIOMAS]" not in idiomas_txt:
        add_secao_com_linha_pdf("Idiomas")
        story.append(Paragraph(idiomas_txt.replace('\n', '<br/>'), estilo_texto))

    if projetos_txt and "[PROJETOS E RESULTADOS]" not in projetos_txt:
        add_secao_com_linha_pdf("Projetos e Resultados")
        story.append(Paragraph(projetos_txt.replace('\n', '<br/>'), estilo_texto))

    if habilidades_txt:
        add_secao_com_linha_pdf("Habilidades e Competências")
        hab_formatadas = habilidades_txt.replace('-', '•').replace('\n', '<br/>')
        story.append(Paragraph(hab_formatadas, estilo_texto))

    if experiencia_txt:
        add_secao_com_linha_pdf("Experiência Profissional")
        
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

    # Trabalho Voluntário posicionado após a Experiência Profissional no PDF
    if voluntariado_txt and "[TRABALHO VOLUNTÁRIO]" not in voluntariado_txt:
        add_secao_com_linha_pdf("Trabalho Voluntário")
        story.append(Paragraph(voluntariado_txt.replace('\n', '<br/>'), estilo_texto))

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
