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

st.set_page_config(page_title="Otimizador de Currículos - Padrão Sala de Emprego", page_icon="📄", layout="wide")

st.title("📄 Otimizador de Currículos Profissional (Padrão Sala de Emprego)")
st.markdown("Ferramenta automatizada ajustada estritamente à fonte Times New Roman, cores pretas sólidas e ícones integrados automaticamente.")

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

# --- FUNÇÃO PARA GERAR O MANUAL DO COLABORADOR EM PDF ---
def gerar_manual_pdf():
    buffer = io.BytesIO()
    doc_m = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle('ManTitulo', parent=styles['Heading1'], fontSize=15, leading=18, fontName="Times-Bold", textColor=colors.black, spaceAfter=4)
    estilo_sub = ParagraphStyle('ManSub', parent=styles['Heading2'], fontSize=11, leading=14, fontName="Times-Bold", textColor=colors.black, spaceBefore=10, spaceAfter=3)
    estilo_texto = ParagraphStyle('ManTexto', parent=styles['Normal'], fontSize=9, leading=13, fontName="Times-Roman", textColor=colors.black, spaceAfter=4)
    
    story_m = [
        Paragraph("<b>MANUAL DO COLABORADOR: OTIMIZADOR DE CURRÍCULOS</b>", estilo_titulo),
        HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=8, spaceBefore=2),
        
        Paragraph("<b>1. Visão Geral da Ferramenta</b>", estilo_sub),
        Paragraph("Ferramenta corporativa desenvolvida para reestruturar e otimizar currículos com base estritamente nos dados reais de cada candidato, direcionando-os para vagas específicas com um padrão visual profissional (fonte Times New Roman, cores pretas, ícones integrados e alinhamento refinado).", estilo_texto),
        
        Paragraph("<b>2. Passo a Passo para Utilização</b>", estilo_sub),
        Paragraph("• <b>Passo 1 (Chave de API):</b> Na barra lateral esquerda, insira a sua chave de API individual no campo indicado.", estilo_texto),
        Paragraph("• <b>Passo 2 (Escolha do Formato):</b> Selecione 'Modelo Clássico (Sem Foto)' ou 'Modelo Com Foto'. Se escolher com foto, faça o upload da imagem (JPG ou PNG) do candidato.", estilo_texto),
        Paragraph("• <b>Passo 3 (Inserção de Dados):</b> No campo superior esquerdo, cole o currículo antigo completo do cliente. No campo abaixo, cole a descrição ou os requisitos da vaga pretendida.", estilo_texto),
        Paragraph("• <b>Passo 4 (Geração):</b> Clique no botão azul <b>'🚀 Otimizar Currículo Agora'</b> e aguarde o processamento.", estilo_texto),
        Paragraph("• <b>Passo 5 (Download):</b> Baixe o resultado final pronto em formato <b>Word (.docx)</b> ou <b>PDF</b>.", estilo_texto),
        
        Paragraph("<b>3. Vantagens e Qualidades</b>", estilo_sub),
        Paragraph("• Padronização visual rigorosa e imediata dos currículos da consultoria.<br/>• Fidelidade total ao histórico verídico do candidato (sem invenção de dados).<br/>• Agilidade operacional, reduzindo o tempo gasto com formatações manuais.", estilo_texto),
        
        Paragraph("<b>4. Cuidados e Limitações (Versão Gratuita)</b>", estilo_sub),
        Paragraph("• <b>Limites de Cota por Minuto:</b> Como a API opera no plano gratuito, um volume massivo de envios simultâneos pode gerar interrupções temporárias.<br/>• <b>Inatividade Temporária:</b> Caso o limite seja atingido, a ferramenta exibirá um aviso e normalizará automaticamente após 1 a 2 minutos.<br/>• <b>Chave Individual:</b> Nunca compartilhe sua chave de API pessoal com colegas para evitar o esgotamento prematuro da cota.", estilo_texto),
        
        Paragraph("<b>5. O que fazer em caso de Erros?</b>", estilo_sub),
        Paragraph("Se ocorrer algum aviso de limite excedido, aguarde de 1 a 2 minutos e tente novamente. Caso o problema persista devido ao alto volume de trabalho, comunique a gestão para avaliarmos a transição para o plano pago corporativo.", estilo_texto)
    ]
    
    doc_m.build(story_m)
    buffer.seek(0)
    return buffer

st.sidebar.markdown("---")
st.sidebar.markdown("### 📘 Documentação")
manual_pdf_buffer = gerar_manual_pdf()
st.sidebar.download_button(
    label="📥 Baixar Manual do Colaborador (PDF)",
    data=manual_pdf_buffer,
    file_name="manual_colaborador_otimizador.pdf",
    mime="application/pdf",
    use_container_width=True
)

ativa_api_key = st.session_state.get("api_key", "")

# --- ÁREA PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Dados do Candidato e Vaga")
    curriculo_antigo = st.text_area("Cole o Currículo Atual do Cliente:", height=200, placeholder="Cole aqui o texto do currículo antigo...")
    descricao_vaga = st.text_area("Cole a Descrição da Vaga ou Palavras-Chave:", height=150, placeholder="Cole a descrição da vaga...")

with col2:
    st.subheader("2️⃣ Instruções e Execução")
    st.info("O sistema gerará o currículo com fonte Times, cores 100% pretas e ícones padronizados integrados automaticamente.")
    
    gerar_btn = st.button("🚀 Otimizar Currículo Agora", type="primary", use_container_width=True)

def chamar_gemini_com_retry(client, prompt_texto):
    modelos_para_tentar = [
        'gemini-3.7-flash',
        'gemini-3.6-flash',
        'gemini-3.1-pro-preview'
    ]
    
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
                erro_str = str(e)
                erros_acumulados.append(f"[{modelo}] {erro_str}")
                
                if "429" in erro_str or "RESOURCE_EXHAUSTED" in erro_str:
                    break 
                
                time.sleep(1)
                continue
                
    raise Exception(f"Todos os modelos testados enfrentaram indisponibilidade ou limite de cota. Detalhes do último erro: {erros_acumulados[-1]}")

if gerar_btn:
    if not ativa_api_key:
        st.error("⚠️ Por favor, insira a sua Chave de API na barra lateral para continuar.")
    elif not curriculo_antigo or not descricao_vaga:
        st.warning("⚠️ Preencha tanto o currículo antigo quanto a descrição da vaga.")
    else:
        with st.spinner("A processar e reestruturando o currículo com os modelos atuais..."):
            try:
                client = genai.Client(api_key=ativa_api_key)
                
                prompt_sistema = f"""
                Você é um consultor especialista em RH. A sua tarefa é OTIMIZAR o currículo antigo do candidato com base estritamente nos dados verídicos dele, direcionando-o para a vaga desejada.
                RESTRIÇÃO ABSOLUTA: NÃO invente dados, experiências, cursos ou informações que não constem no currículo original. Mantenha os factos reais e adapte apenas a linguagem para ficar profissional e alinhada à vaga.
                IMPORTANTE SOBRE O ESPAÇAMENTO: No bloco [PERFIL PROFISSIONAL], escreva os parágrafos em sequência direta, utilizando apenas quebras simples (sem linhas em branco extras entre um parágrafo e outro) para evitar espaçamento vertical excessivo.
                
                Gere o conteúdo final preenchendo rigorosamente estes blocos com tags em maiúsculas:
                
                [NOME]
                [Nome real do candidato extraído do currículo antigo]
                
                [CARGO]
                [Cargo ou até 3 opções baseadas na vaga e na experiência real, separadas por barra]
                
                [CONTATOS]
                [Bairro, Cidade | Telefone | E-mail | LinkedIn reais extraídos do currículo]
                
                [PERFIL PROFISSIONAL]
                (Escrito estritamente em 3ª pessoa e dividido em exatamente 3 parágrafos, um logo abaixo do outro, sem linhas em branco intermediárias, baseando-se no histórico real do candidato):
                1º Parágrafo: Profissional atuante na área [cargo/objetivo], destacando-se pela capacidade em [competências reais do candidato], garantindo eficiência nas rotinas da área.
                2º Parágrafo: Expertises em [atividades e competências reais do candidato alinhadas aos requisitos da vaga].
                3º Parágrafo: Conhecimentos em [ferramentas, sistemas e fundamentos reais apresentados no histórico].
                
                [FORMAÇÃO ACADÊMICA]
                [Curso real | Instituição real - Ano real]
                
                [CURSOS E CERTIFICAÇÕES]
                [Curso real | Instituição real - Ano real] (Apenas se houver no original)
                
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

    # Estilos ajustados: fonte um pouco maior para legibilidade e espaços compactos
    estilo_nome = ParagraphStyle('NomeEstilo', parent=styles['Heading1'], fontSize=13, leading=15, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)
    estilo_cargo = ParagraphStyle('CargoEstilo', parent=styles['Normal'], fontSize=9.5, leading=11, textColor=cor_total, fontName="Times-Italic", spaceAfter=2)
    estilo_contato_dir = ParagraphStyle('ContatoDirEstilo', parent=styles['Normal'], fontSize=9, leading=12, textColor=cor_total, fontName="Times-Roman", alignment=2) 
    
    estilo_titulo_secao = ParagraphStyle('SecaoEstilo', parent=styles['Heading2'], fontSize=9.5, leading=11, textColor=cor_total, spaceBefore=5, spaceAfter=2, fontName="Times-Bold")
    estilo_texto = ParagraphStyle('TextoEstilo', parent=styles['Normal'], fontSize=9, leading=11.5, textColor=cor_total, fontName="Times-Roman", spaceAfter=2, alignment=4)
    estilo_exp_empresa = ParagraphStyle('EmpresaExpEstilo', parent=styles['Normal'], fontSize=9, leading=11, textColor=cor_total, fontName="Times-Roman", spaceAfter=1)
    estilo_exp_cargo = ParagraphStyle('CargoExpEstilo', parent=styles['Normal'], fontSize=9, leading=11, textColor=cor_total, fontName="Times-Bold", spaceAfter=1)

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
    habilidades_txt = extrair_tag("HABILIDADES E COMPETÊNCIAS", texto_gerado)
    experiencia_txt = extrair_tag("EXPERIÊNCIA PROFISSIONAL", texto_gerado)

    partes_contato = [p.strip() for p in contato_txt.replace('|', ',').split(',') if p.strip()]
    cidade_bairro = partes_contato[0] if len(partes_contato) > 0 else "Bairro, Cidade"
    telefone = partes_contato[1] if len(partes_contato) > 1 else "Telefone"
    email = partes_contato[2] if len(partes_contato) > 2 else "E-mail"
    linkedin = partes_contato[3] if len(partes_contato) > 3 else "https://www.linkedin.com/in/"

    def criar_icone_casa():
        d = Drawing(14, 14)
        d.add(Rect(0, 0, 14, 14, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Polygon([3, 7.5, 7, 3.5, 11, 7.5], fillColor=colors.white, strokeColor=colors.white))
        d.add(Rect(4.5, 2, 5, 5.5, fillColor=colors.white, strokeColor=colors.white))
        d.add(Rect(5.5, 2, 3, 3, fillColor=colors.black, strokeColor=colors.black))
        return d

    def criar_icone_telefone():
        d = Drawing(14, 14)
        d.add(Rect(0, 0, 14, 14, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Polygon([3.5, 10, 5, 11.5, 6.5, 10, 5.5, 9, 8.5, 6, 9.5, 7, 11, 5.5, 9.5, 4, 8.5, 5, 5.5, 8], fillColor=colors.white, strokeColor=colors.white))
        return d

    def criar_icone_email():
        d = Drawing(14, 14)
        d.add(Rect(0, 0, 14, 14, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(Rect(2, 3.5, 10, 7, rx=1, ry=1, fillColor=colors.white, strokeColor=colors.white))
        d.add(Polygon([2.2, 10.3, 7, 7, 11.8, 10.3], fillColor=colors.black, strokeColor=colors.black))
        return d

    def criar_icone_linkedin():
        d = Drawing(14, 14)
        d.add(Rect(0, 0, 14, 14, rx=2, ry=2, fillColor=colors.black, strokeColor=colors.black))
        d.add(String(2.5, 3, "in", fontName="Helvetica-Bold", fontSize=9, fillColor=colors.white))
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
        ], colWidths=[54, 291])
        tabela_esquerda_foto.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'LEFT'),
            ('LEFTPADDING', (1,0), (1,0), 6),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        coluna_esquerda = tabela_esquerda_foto
    else:
        coluna_esquerda = [
            Paragraph(f"<b>{nome_txt}</b>", estilo_nome),
            Paragraph(cargo_txt, estilo_cargo)
        ]

    tabela_contatos_direita = Table([
        [Paragraph(cidade_bairro, estilo_contato_dir), criar_icone_casa()],
        [Paragraph(telefone, estilo_contato_dir), criar_icone_telefone()],
        [Paragraph(email, estilo_contato_dir), criar_icone_email()],
        [Paragraph(linkedin, estilo_contato_dir), criar_icone_linkedin()]
    ], colWidths=[160, 35])
    
    tabela_contatos_direita.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,-1), 'RIGHT'),
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ('LEFTPADDING', (1,0), (1,-1), 8),
        ('RIGHTPADDING', (0,0), (0,-1), 4),
    ]))

    t_header = Table([ [coluna_esquerda, tabela_contatos_direita] ], colWidths=[345, 195])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(t_header)
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.6, color=cor_total, spaceAfter=4, spaceBefore=0))

    def adicionar_secao(titulo, conteudo_html):
        if conteudo_html:
            story.append(Paragraph(f"<b>{titulo}</b>", estilo_titulo_secao))
            story.append(HRFlowable(width="100%", thickness=0.3, color=cor_total, spaceAfter=2, spaceBefore=1))
            story.append(Paragraph(conteudo_html, estilo_texto))

    if perfil_txt:
        paragrafos_perfil = [p.strip() for p in perfil_txt.split('\n\n') if p.strip()]
        perfil_formatado = "<br/><br/>".join(paragrafos_perfil)
        adicionar_secao("Perfil Profissional", perfil_formatado)

    if formacao_txt:
        adicionar_secao("Formação Acadêmica", formacao_txt.replace('\n', '<br/>'))

    if cursos_txt and "[CURSOS E CERTIFICAÇÕES]" not in cursos_txt:
        adicionar_secao("Cursos e Certificações", cursos_txt.replace('\n', '<br/>'))

    if habilidades_txt:
        hab_formatadas = habilidades_txt.replace('-', '•').replace('\n', '<br/>')
        adicionar_secao("Habilidades e Competências", hab_formatadas)

    if experiencia_txt:
        story.append(Paragraph("<b>Experiência Profissional</b>", estilo_titulo_secao))
        story.append(HRFlowable(width="100%", thickness=0.3, color=cor_total, spaceAfter=2, spaceBefore=1))
        
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
            
            story.append(Spacer(1, 2))

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
