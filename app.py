import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import io
import PyPDF2
import docx

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de Informes 1º Bach 7", layout="wide")

# --- GESTIÓN DEL ESTADO (Para el botón de borrar) ---
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

def reiniciar_app():
    # Incrementamos la clave para forzar que el uploader se cree de nuevo vacío
    st.session_state.uploader_key += 1
    # Recargamos la página
    st.rerun()

# --- TÍTULOS ---
st.title("📊 Analizador Multidocumento - IES Lucía de Medrano")
st.subheader("1º Bachillerato 7 - Tutor: Jose Carlos Tejedor")

# --- BARRA LATERAL ---
st.sidebar.header("Configuración")
api_key = st.sidebar.text_input("Introduce tu API Key de OpenAI", type="password")

st.sidebar.markdown("---")
st.sidebar.write("¿Quieres subir nuevos archivos?")
# Botón para borrar
if st.sidebar.button("🗑️ Borrar todo y subir nuevos archivos", type="primary"):
    reiniciar_app()

# --- FUNCIONES DE EXTRACCIÓN Y CÁLCULO ---

def extract_text_from_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(file):
    doc = docx.Document(file)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

def process_data_with_ai(text_data, api_key, filename):
    client = openai.OpenAI(api_key=api_key)
    prompt = f"""
    Analiza el texto del archivo '{filename}'.
    Extrae calificaciones y devuelve EXCLUSIVAMENTE un CSV.
    Columnas: "Alumno", "Materia", "Nota".
    Reglas:
    1. Materia: abreviaturas (MAT, LE, ING).
    2. Nota: número decimal (5.0). Texto a número (Bien=6, Notable=7.5).
    3. SOLO el CSV, nada más.
    
    Texto: {text_data[:15000]}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "system", "content": "Experto en datos académicos."},
                      {"role": "user", "content": prompt}],
            temperature=0
        )
        csv_content = response.choices[0].message.content
        csv_content = csv_content.replace("```csv", "").replace("```", "")
        return pd.read_csv(io.StringIO(csv_content))
    except Exception as e:
        st.error(f"Error en {filename}: {e}")
        return None

def analyze_data(df):
    df['Nota'] = pd.to_numeric(df['Nota'], errors='coerce')
    df['Aprobado'] = df['Nota'] >= 5
    
    stats_alumno = df.groupby('Alumno').agg(
        Suspensos=('Nota', lambda x: (x < 5).sum()),
        Media=('Nota', 'mean')
    ).reset_index()
    
    total = len(stats_alumno)
    cero = stats_alumno[stats_alumno['Suspensos'] == 0].shape[0]
    uno = stats_alumno[stats_alumno['Suspensos'] == 1].shape[0]
    dos = stats_alumno[stats_alumno['Suspensos'] == 2].shape[0]
    mas_dos = stats_alumno[stats_alumno['Suspensos'] > 2].shape[0]
    
    stats_materia = df.groupby('Materia').agg(
        Total=('Nota', 'count'),
        Aprobados=('Aprobado', 'sum'),
        Suspensos=('Nota', lambda x: (x < 5).sum()),
        Media=('Nota', 'mean')
    ).reset_index()
    
    stats_materia['Pct_Suspensos'] = (stats_materia['Suspensos'] / stats_materia['Total']) * 100
    
    return {
        "total": total, "cero": cero, "uno": uno, "dos": dos, "mas_dos": mas_dos,
        "pasan": cero + uno + dos, "no_pasan": mas_dos,
        "stats_materia": stats_materia,
        "ranking": stats_materia.sort_values('Pct_Suspensos', ascending=False),
        "mejor": stats_alumno.loc[stats_alumno['Media'].idxmax()],
        "peor": stats_alumno.loc[stats_alumno['Media'].idxmin()],
        "media_global": df['Nota'].mean()
    }

def generate_word_report(r, plots):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = s.page_height, s.page_width
    s.orientation = WD_ORIENT.LANDSCAPE

    doc.add_heading('Informe 1º Bach 7 - IES Lucía de Medrano', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('Tutor: Jose Carlos Tejedor').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. Resumen', 1)
    doc.add_paragraph(f"• Aprueban todo: {r['cero']} ({r['cero']/r['total']:.1%})")
    doc.add_paragraph(f"• 1 Suspenso: {r['uno']}")
    doc.add_paragraph(f"• 2 Suspensos: {r['dos']}")
    doc.add_paragraph(f"• Promocionan: {r['pasan']} ({r['pasan']/r['total']:.1%})").bold = True
    doc.add_paragraph(f"• No promocionan: {r['no_pasan']} ({r['no_pasan']/r['total']:.1%})").bold = True
    
    doc.add_heading('2. Materias', 1)
    t = doc.add_table(1, 4)
    t.style = 'Table Grid'
    h = t.rows[0].cells
    h[0].text='Materia'; h[1].text='Suspensos'; h[2].text='%'; h[3].text='Media'
    for _, row in r['ranking'].iterrows():
        c = t.add_row().cells
        c[0].text=str(row['Materia']); c[1].text=str(row['Suspensos'])
        c[2].text=f"{row['Pct_Suspensos']:.1f}%"; c[3].text=f"{row['Media']:.2f}"
        
    doc.add_heading('3. Gráficas', 1)
    for p in plots:
        doc.add_picture(p, width=Inches(7))
        doc.add_paragraph(" ")
        
    return doc

# --- INTERFAZ PRINCIPAL ---

# El key dinámico es lo que permite "resetear" el botón
uploaded_files = st.file_uploader(
    "Sube tus archivos (Excel, PDF, Word)", 
    type=['xlsx', 'pdf', 'docx', 'doc'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}" 
)

col1, col2 = st.columns([1, 4])
with col1:
    generate_btn = st.button("🚀 Generar Informe", type="primary")

if uploaded_files and generate_btn:
    if not api_key:
        st.error("⚠️ Introduce la API Key en la barra lateral.")
    else:
        dfs = []
        bar = st.progress(0)
        status = st.empty()
        
        for i, f in enumerate(uploaded_files):
            status.text(f"Leyendo {f.name}...")
            if f.name.endswith('.xlsx'):
                d = pd.read_excel(f)
                if 'Nota' not in d.columns:
                    d = d.melt(id_vars=[d.columns[0]], var_name="Materia", value_name="Nota")
                    d.columns = ['Alumno', 'Materia', 'Nota']
                dfs.append(d)
            elif f.name.endswith('.pdf'):
                dfs.append(process_data_with_ai(extract_text_from_pdf(f), api_key, f.name))
            elif 'doc' in f.name:
                dfs.append(process_data_with_ai(extract_text_from_docx(f), api_key, f.name))
            bar.progress((i+1)/len(uploaded_files))
            
        if dfs:
            full_df = pd.concat(dfs, ignore_index=True)
            res = analyze_data(full_df)
            
            # Gráficas
            plots = []
            fig, ax = plt.subplots(figsize=(10,4))
            res['stats_materia'].sort_values('Suspensos').plot(
                x='Materia', y=['Aprobados','Suspensos'], kind='bar', stacked=True, 
                color=['#81c784','#e57373'], ax=ax)
            img = io.BytesIO(); plt.savefig(img, format='png'); img.seek(0); plots.append(img)
            
            fig2, ax2 = plt.subplots(figsize=(6,4))
            ax2.bar(['0','1','2','+2'], [res['cero'], res['uno'], res['dos'], res['mas_dos']], color='orange')
            img2 = io.BytesIO(); plt.savefig(img2, format='png'); img2.seek(0); plots.append(img2)
            
            # Word
            doc = generate_word_report(res, plots)
            bio = io.BytesIO(); doc.save(bio); bio.seek(0)
            
            status.success("✅ Informe listo")
            st.download_button("📥 Descargar Word", data=bio, file_name="Informe_1Bach7.docx")
            
            # Resumen visual
            c1, c2, c3 = st.columns(3)
            c1.metric("Promocionan", f"{res['pasan']} ({res['pasan']/(res['total'] or 1):.0%})")
            c2.metric("No promocionan", f"{res['no_pasan']}")
            c3.metric("Nota Media", f"{res['media_global']:.2f}")
            st.dataframe(res['ranking'][['Materia','Suspensos','Pct_Suspensos']])