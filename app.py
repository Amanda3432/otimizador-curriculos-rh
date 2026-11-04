import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable, KeepTogether
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
from xml.sax.saxutils import escape

st.set_page_config(
    page_title="Sala de Emprego | Otimizador de Currículos", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CABEÇALHO PROFISSIONAL DA CONSULTORIA ---
st.markdown("<h1 style='text-align: left; color: #111111; font-family: sans-serif; font-size: 2.2rem; margin-bottom: 0px;'>Sala de Emprego — Otimizador de Currículos</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: left; color: #555555; font-size: 1.1rem; margin-top: 5px; margin-bottom: 25px;'>Ferramenta de Otimização de Currículo para Triagens em ATS do Sala de Emprego</p>", unsafe_allow_html=True)
st.markdown("---")

# --- CONFIGURAÇÃO AUTOMÁTICA DA API ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        ativa_api_key = st.secrets["GEMINI_API_KEY"]
    else:
        ativa_api_key = ""
except Exception:
    ativa_api_key = ""

# --- FUNÇÃO DE SEGURANÇA: SANITIZAÇÃO DE TEXTO CONTRA PROMPT INJECTION ---
def sanitizar_texto_entrada(texto):
    if not texto:
        return ""
    padroes_maliciosos = [
        r"ignore previous instructions",
        r"ignore as instruções anteriores",
        r"system prompt",
        r"revele a chave",
        r"show api key",
        r"developer mode"
    ]
    texto_limpo = texto
    for padrao in padroes_maliciosos:
        texto_limpo = re.sub(padrao, "[REDACTED_SECURITY]", texto_limpo, flags=re.IGNORECASE)
    return texto_limpo

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

# --- BARRA LATERAL (MODELOS E FOTO) ---
st.sidebar.markdown("### 📋 Modelos Disponíveis")
modelo_escolhido = st.sidebar.radio(
