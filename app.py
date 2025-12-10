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
    st.session_state.uploader_key += 1
    st.rerun()

# --- TÍTULOS ---
st.title("📊 Analizador Multidocumento - IES Lucía de Medrano")
st.subheader("1º Bachillerato 7 - Tutor: Jose Carlos Tejedor")

# --- BARRA LATERAL ---
st.sidebar.header("Configuración")
api_key = st.sidebar.text_input("Introduce tu API Key de OpenAI (sk-...)", type="password")

st.sidebar.markdown("---")
st.sidebar.write("¿Quieres subir nuevos archivos?")
if st.sidebar.button("🗑️ Borrar todo y reiniciar", type="primary"):
    reiniciar_app()

# --- FUNCIONES ---

def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
        return text
    except Exception:
        return ""

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        text = "\n".join([para.text for para in doc.paragraphs])
        return text
    except Exception:
        return ""

def process_data_with_ai(text_data, api_key, filename):
    if not text_data:
        return None
        
    client = openai.OpenAI(api_key=api_key)
    prompt = f"""
    Analiza el texto del archivo '{filename}'.
    Extrae calificaciones y devuelve EXCLUSIVAMENTE un CSV.
    Columnas: "Alumno", "Materia", "Nota".
    Reglas:
    1. Materia: abreviaturas (MAT, LE, ING).
    2. Nota: número decimal (5.0). Texto a número (Bien=6, Notable=7.5).
    3. SOLO el CSV, sin explicaciones ni comillas de código.
    
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
        # Limpiar posibles marcas de markdown
        csv_content = csv_content.replace("```csv", "").replace("```", "").strip()
        return pd.read_csv(io.StringIO(csv_content))
    except Exception as e:
        st.error(f"Error procesando {filename}: {e}")
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
    total_safe = r['total'] if r['total'] > 0 else 1
    doc.add_paragraph(f"• Aprueban todo: {r['cero']} ({r['cero']/total_safe:.1%})")
    doc.add_paragraph(f"• 1 Suspenso: {r['uno']}")
    doc.add_paragraph(f"• 2 Suspensos: {r['dos']}")
    doc.add_paragraph(f"• Promocionan: {r['pasan']} ({r['pasan']/total_safe:.1%})").bold = True
    doc.add_paragraph(f"• No promocionan: {r['no_pasan']} ({r['no_pasan']/total_safe:.1%})").bold = True
    
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

uploaded_files = st.file_uploader(
    "Sube tus archivos (Excel, PDF, Word)", 
    type=['xlsx', 'pdf', 'docx', 'doc'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}" 
)

col1, col2 = st.columns([1, 4])
with col1:
    generate_btn = st.button("🚀 Generar Informe", type="primary")

# --- LÓGICA DE EJECUCIÓN ---

if uploaded_files and generate_btn:
    if not api_key:
        st.error("⚠️ Error: Introduce la API Key en la barra lateral izquierda.")
    else:
        dfs = []
        bar = st.progress(0)
        status = st.empty()
        
        for i, f in enumerate(uploaded_files):
            status.text(f"Leyendo {f.name}...")
            df_temp = None
            
            # Procesar según tipo
            if f.name.endswith('.xlsx'):
                try:
                    d = pd.read_excel(f)
                    if 'Nota' not in d.columns:
                        d = d.melt(id_vars=[d.columns[0]], var_name="Materia", value_name="Nota")
                        d.columns = ['Alumno', 'Materia', 'Nota']
                    df_temp = d
                except:
                    pass
            elif f.name.endswith('.pdf'):
                text = extract_text_from_pdf(f)
                if text: df_temp = process_data_with_ai(text, api_key, f.name)
            elif 'doc' in f.name:
                text = extract_text_from_docx(f)
                if text: df_temp = process_data_with_ai(text, api_key, f.name)
            
            # --- PROTECCIÓN CONTRA ERRORES ---
            # Solo añadimos si se extrajeron datos correctamente
            if df_temp is not None and not df_temp.empty:
                dfs.append(df_temp)
            
            bar.progress((i+1)/len(uploaded_files))
            
        if dfs:
            full_df = pd.concat(dfs, ignore_index=True)
            res = analyze_data(full_df)
            
            # Gráficas
            plots = []
            try:
                fig, ax = plt.subplots(figsize=(10,4))
                res['stats_materia'].sort_values('Suspensos').plot(
                    x='Materia', y=['Aprobados','Suspensos'], kind='bar', stacked=True, 
                    color=['#81c784','#e57373'], ax=ax)
                plt.title("Aprobados vs Suspensos")
                img = io.BytesIO(); plt.savefig(img, format='png'); img.seek(0); plots.append(img)
                
                fig2, ax2 = plt.subplots(figsize=(6,4))
                ax2.bar(['0','1','2','+2'], [res['cero'], res['uno'], res['dos'], res['mas_dos']], color='orange')
                plt.title("Distribución de suspensos")
                img2 = io.BytesIO(); plt.savefig(img2, format='png'); img2.seek(0); plots.append(img2)
            except Exception as e:
                st.warning("No se pudieron generar todas las gráficas.")

            # Word
            try:
                doc = generate_word_report(res, plots)
                bio = io.BytesIO(); doc.save(bio); bio.seek(0)
                status.success("✅ Informe listo")
                st.download_button("📥 Descargar Word", data=bio, file_name="Informe_1Bach7.docx")
            except Exception as e:
                st.error(f"Error generando Word: {e}")
            
            # Resumen visual
            c1, c2, c3 = st.columns(3)
            total_div = res['total'] if res['total'] > 0 else 1
            c1.metric("Promocionan", f"{res['pasan']} ({res['pasan']/total_div:.0%})")
            c2.metric("No promocionan", f"{res['no_pasan']}")
            c3.metric("Nota Media", f"{res['media_global']:.2f}")
            
        else:
            st.error("❌ No se pudieron extraer datos. Revisa que la API Key sea correcta (empieza por 'sk-') y que los archivos tengan texto legible.")
