import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import io
import PyPDF2
import docx

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Acta de Evaluación", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS ---
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 2px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0px 0px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff;
        border-top: 2px solid #4e8cff;
    }
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .highlight-box {
        background-color: #e8f0fe;
        border-left: 5px solid #4285f4;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
        color: #004085;
    }
</style>
""", unsafe_allow_html=True)

# --- GESTIÓN DE ESTADO ---
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0
if 'data' not in st.session_state:
    st.session_state.data = None

def reiniciar_app():
    st.session_state.data = None
    st.session_state.uploader_key += 1
    st.rerun()

# --- FUNCIONES DE EXTRACCIÓN ---
def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
        return text
    except: return ""

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except: return ""

def process_data_with_ai(text_data, api_key, filename):
    if not text_data: return None
    client = openai.OpenAI(api_key=api_key)
    prompt = f"""
    Analiza el texto académico de '{filename}'.
    Extrae calificaciones en formato CSV. Columnas: "Alumno", "Materia", "Nota".
    Reglas:
    1. Materia: abreviatura (MAT, LE, ING).
    2. Nota: número decimal (usar punto). Texto a número (Bien=6, Notable=7.5).
    3. SOLO CSV.
    Texto: {text_data[:15000]}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}], temperature=0
        )
        csv = response.choices[0].message.content.replace("```csv", "").replace("```", "").strip()
        return pd.read_csv(io.StringIO(csv))
    except: return None

# --- FUNCIONES DE WORD (INDIVIDUAL Y MASIVO) ---

def add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos):
    """Función auxiliar para añadir una página de alumno al documento"""
    doc.add_heading(f'Informe Individual: {alumno}', 0)
    p = doc.add_paragraph(f"Nota Media: {media:.2f} | Materias Suspensas: {suspensos}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    t = doc.add_table(rows=1, cols=2)
    t.style = 'Table Grid'
    t.autofit = False 
    t.columns[0].width = Inches(4)
    t.columns[1].width = Inches(1.5)
    
    # Cabecera tabla
    hdr = t.rows[0].cells
    hdr[0].text = 'Materia'
    hdr[1].text = 'Cal
