import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import io
import pdfplumber
import PyPDF2  # Recuperamos esto como respaldo
import docx
from datetime import datetime
import re # Para limpieza de texto avanzada

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
    div[data-testid="stDataEditor"] {
        border: 2px solid #4e8cff;
        border-radius: 5px;
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

# --- FUNCIONES DE EXTRACCIÓN ROBUSTAS ---
def get_pdf_text(file):
    """Intenta extraer texto con múltiples métodos"""
    text = ""
    
    # Método 1: PDFPlumber (Mejor para tablas visuales)
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text(x_tolerance=2, y_tolerance=2)
                if extracted:
                    text += extracted + "\n"
    except: pass
    
    # Método 2: PyPDF2 (Respaldo si el anterior falla o da vacío)
    if len(text) < 50: 
        try:
            file.seek(0)
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() or ""
        except: pass
        
    return text

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except: return ""

# --- LIMPIEZA DE DATOS (PYTHON PURO) ---
def limpiar_nombre_alumno(texto):
    """
    1. Quita números iniciales (1 ANTHONY -> ANTHONY)
    2. Formatea Apellidos, Nombre -> Nombre Apellidos
    """
    if not isinstance(texto, str): return str(texto)
    
    # Paso 1: Quitar números e índices al principio (ej: "1 ", "2.")
    texto_limpio = re.sub(r'^\d+[\.\-\s]+', '', texto.strip())
    
    # Paso 2: Ordenar Nombre Apellidos
    if ',' in texto_limpio:
        partes = texto_limpio.split(',')
        if len(partes) >= 2:
            apellidos = partes[0].strip()
            nombre = partes[1].strip()
            return f"{nombre} {apellidos}"
            
    return texto_limpio

def process_data_with_ai(text_data, api_key, filename):
    if not text_data or len(text_data) < 10: 
        return None
        
    client = openai.OpenAI(api_key=api_key)
    
    prompt = f"""
    Analiza este texto de acta de evaluación ('{filename}').
    
    OBJETIVO: Extraer calificaciones.
    
    PROBLEMA CONOCIDO:
    - Las notas (números del 1 al 10) suelen estar agrupadas al final o a la derecha.
    - Los nombres tienen un índice delante (ej: "1 PEREZ, JUAN"). ESE "1" NO ES LA NOTA.
    
    INSTRUCCIONES:
    1. Extrae cada alumno con su asignatura y su nota REAL.
    2. Si ves un número de índice delante del nombre, IGNÓRALO.
    3. Asocia la primera fila de notas con el primer alumno.
    
    SALIDA CSV (3 columnas):
    Alumno, Materia, Nota
    
    Ejemplo de salida esperada:
    PEREZ GOMEZ, JUAN, MAT, 8
    PEREZ GOMEZ, JUAN, LE, 5
    
    Texto:
    {text_data[:15000]}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}], temperature=0
        )
        content = response.choices[0].message.content
        # Limpieza de bloques de código markdown
        csv_str = content.replace("```csv", "").replace("```", "").strip()
        
        # Validación mínima
        if "," not in csv_str: return None
        
        df = pd.read_csv(io.StringIO(csv_str))
        
        # FUERZA BRUTA PARA NOMBRES DE COLUMNA
        # Si tiene 3 columnas, las renombramos nosotros para evitar errores
        if len(df.columns) == 3:
            df.columns = ['Alumno', 'Materia', 'Nota']
        
        # Aplicar limpieza de nombres (Quitar índice y reordenar)
        if 'Alumno' in df.columns:
            df['Alumno'] = df['Alumno'].apply(limpiar_nombre_alumno)
            
        return df
        
    except Exception as e:
        st.error(f"Error IA: {str(e)}")
        return None

# --- GENERACIÓN DE TEXTOS AUTOMÁTICOS ---
def generar_comentario_individual(alumno, datos_alumno):
    suspensos = datos_alumno[datos_alumno['Nota'] < 5]
    num_suspensos = len(suspensos)
    lista_suspensas = suspensos['Materia'].tolist()
    
    txt = f"El alumno/a {alumno} tiene actualmente {num_suspensos} materias suspensas."
    
    if num_suspensos == 0:
        txt = "No tiene ninguna materia suspensa. ¡Excelente trabajo! Se recomienda mantener la constancia."
    elif num_suspensos == 1:
        txt += f" La materia pendiente es: {', '.join(lista_suspensas)}. Recuperación factible con plan de refuerzo."
    elif num_suspensos == 2:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. Situación límite. Se aconseja refuerzo urgente."
    else:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. Situación preocupante que compromete la promoción."
    return txt

def generar_valoracion_detallada(res):
    txt = f"Nota media global: {res['media_grupo']:.2f}. "
    if res['pct_pasan'] >= 85: txt += "Promoción excelente."
    elif res['pct_pasan'] >= 70: txt += "Promoción satisfactoria."
    else: txt += "Promoción baja, se requiere intervención."
    return txt

# --- WORD INDIVIDUAL ---
def add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat):
    doc.add_heading(f'Informe Individual: {alumno}', 0)
    doc.add_paragraph(f"Nota Media: {media:.2f} | Materias Suspensas: {suspensos}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('Análisis y Recomendaciones', level=2)
    p = doc.add_paragraph(generar_comentario_individual(alumno, datos_alumno))
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    doc.add_heading('Detalle de Calificaciones', level=2)
    t = doc.add_table(rows=1, cols=4)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text = 'Materia'; hdr[1].text = 'Nota'; hdr[2].text = 'Media Clase'; hdr[3].text = 'Dif.'
    
    medias = stats_mat.set_index('Materia')['Media'].to_dict()
    for _, row in datos_alumno.iterrows():
        c = t.add_row().cells
        c[0].text = str(row['Materia'])
        c[1].text = str(row['Nota'])
        media_c = medias.get(row['Materia'], 0)
        c[2].text = f"{media_c:.2f}"
        dif = row['Nota'] - media_c
        c[3].text = f"{dif:+.2f}"
        
        if row['Nota'] < 5:
            run = c[1].paragraphs[0].runs[0]
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.bold = True

    # PIE DE PÁGINA
    doc.add_paragraph("\n\n")
    now = datetime.now()
    meses = ["enero"
