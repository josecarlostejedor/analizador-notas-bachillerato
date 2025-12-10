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
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 0.5rem;
        padding: 1rem;
        color: #0f5132;
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
def extract_table_data_from_pdf(file):
    text_content = ""
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        clean_row = [str(cell).replace("\n", " ") if cell is not None else "" for cell in row]
                        if any(len(c) > 0 for c in clean_row):
                            text_content += " | ".join(clean_row) + "\n"
        return text_content
    except Exception as e:
        return ""

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except: return ""

def process_data_with_ai(text_data, api_key, filename):
    if not text_data: return None
    client = openai.OpenAI(api_key=api_key)
    prompt = f"""
    Analiza esta tabla de acta de evaluación ('{filename}').
    Filas por saltos de línea, columnas por '|'.
    TAREA: Extrae CSV con columnas: "Alumno", "Materia", "Nota".
    REGLAS:
    1. Primera columna es Alumno.
    2. Siguientes son Materias (EF, FILO, etc).
    3. Ignora cabeceras repetidas.
    4. Convierte notas a numérico. Si es texto (ej: EX), pon vacío.
    Datos: {text_data[:15000]}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}], temperature=0
        )
        csv = response.choices[0].message.content.replace("```csv", "").replace("```", "").strip()
        if "," not in csv: return None
        return pd.read_csv(io.StringIO(csv))
    except: return None

# --- LÓGICA DE VALORACIÓN ---
def obtener_valoracion(pct_promocion):
    if pct_promocion >= 90:
        return "Los resultados son excelentes. La inmensa mayoría del grupo ha alcanzado los objetivos previstos."
    elif pct_promocion >= 75:
        return "Los resultados son muy positivos. Una gran parte del alumnado promociona."
    elif pct_promocion >= 60:
        return "Los resultados son aceptables, aunque existe un porcentaje significativo que no alcanza los mínimos."
    else:
        return "Los resultados son preocupantes. Se recomienda revisar estrategias metodológicas."

# --- FUNCIONES DE WORD ---

def add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat):
    doc.add_heading(f'Informe Individual: {alumno}', 0)
    doc.add_paragraph(f"Nota Media: {media:.2f} | Materias Suspensas: {suspensos}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
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
            c[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255,0,0)

def crear_informe_individual(alumno, datos_alumno, media, suspensos, stats_mat):
    doc = Document()
    add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat)
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

def generar_informe_todos_alumnos(df, stats_al, stats_mat):
    doc = Document()
    for i, al in enumerate(stats_al['Alumno'].unique()):
        d_al = df[df['Alumno'] == al]
        info = stats_al[stats_al['Alumno'] == al].iloc[0]
        add_alumno_to_doc(doc, al, d_al, info['Media'], info['Suspensos'], stats_mat)
        if i < len(stats_al)-1: doc.add_page_break()
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

def generate_global_report(datos_resumen, plots, ranking_materias, centro, grupo):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    doc.add_heading(f'Informe de Evaluación - {centro}', 0)
    doc.add_paragraph('Tutor: Jose Carlos Tejedor')

    doc.add_heading('1. Datos Generales del Grupo', 1)
    doc.add_paragraph(f"a) Grupo evaluado: {grupo}")
    doc.add_paragraph(f"b) Número de alumnos en el grupo: {datos_resumen['total_alumnos']}")
    doc.add_paragraph(f"c) Media de notas del grupo: {datos_resumen['media_grupo']:.2f}")
    doc.add_paragraph(f"d) Alumno(s) con media más elevada: {', '.join(datos_resumen['mejores_alumnos'])}")
    doc.add_paragraph(f"e) Alumno(s) con media más baja: {', '.join(datos_resumen['peores_alumnos'])}")

    doc.add_heading('2. Nivel de Promoción', 1)
    p = doc.add_paragraph()
    p.add_run(f"a) Aprueban todo: {datos_resumen['cero']} ({datos_resumen['pct_cero']:.1f}%)\n")
    p.add_run(f"b) Suspenden 1 materia: {datos_resumen['uno']} ({datos_resumen['pct_uno']:.1f}%)\n")
    p.add_run(f"c) Suspenden 2 materias: {datos_resumen['dos']} ({datos_resumen['pct_dos']:.1f}%)\n")
    p.add_run(f"d) Suspenden >2 materias: {datos_resumen['mas_dos']} ({datos_resumen['pct_mas_dos']:.1f}%)")

    doc.add_paragraph("-" * 30)
    p2 = doc.add_paragraph()
    p2.add_run(f"• Promocionan: {datos_resumen['pasan']} ({datos_resumen['pct_pasan']:.1f}%)\n").bold = True
    p2.add_run(f"• No promocionan: {datos_resumen['no_pasan']} ({datos_resumen['pct_no_pasan']:.1f}%)").bold = True

    doc.add_heading('3. Valoración', 1)
    doc.add_paragraph(datos_resumen['valoracion']).italic = True

    doc.add_heading('4. Análisis por Materias', 1)
    t = doc.add_table(rows=1, cols=6)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text = 'Materia'; hdr[1].text = 'Aprobados'; hdr[2].text = '% Apr.'; 
    hdr[3].text = 'Suspensos'; hdr[4].text = '% Susp.'; hdr[5].text = 'Nota Media'
    
    for _, row in ranking_materias.iterrows():
        c = t.add_row().cells
        c[0].text = str(row['Materia']); c[1].text = str(row['Aprobados']); c[2].text = f"{row['Pct_Aprobados']:.1f}%"
        c[3].text = str(row['Suspensos']); c[4].text = f"{row['Pct_Suspensos']:.1f}%"; c[5].text = f"{row['Media']:.2f}"

    doc.add_heading('5. Gráficas', 1)
    for p in plots: doc.add_picture(p, width=Inches(6))

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# --- INTERFAZ ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=50)
    st.title("Configuración")
    api_key = st.text_input("🔑 API Key OpenAI", type="password")
    st.markdown("---")
    centro = st.text_input("Centro", "IES Lucía de Medrano")
    grupo = st.text_input("Grupo", "1º BACH 7")
    curso = st.text_input("Curso", "2024-2025")
    st.markdown("---")
    uploaded_files = st.file_uploader("📂 Subir Actas", type=['xlsx', 'pdf', 'docx', 'doc'], accept_multiple_files=True, key=f"up_{st.session_state.uploader_key}")
    
    if uploaded_files and st.session_state.data is None:
        if st.button("Analizar Archivos", type="primary"):
            if not api_key: st.error("Falta la API Key")
            else:
                dfs = []
                bar = st.progress(0)
                for i, f in enumerate(uploaded_files):
                    df_t = None
                    if f.name.endswith('.xlsx'):
                        try:
                            d = pd.read_excel(f)
                            if 'Nota' not in d.columns:
                                d = d.melt(id_vars=[d.columns[0]], var_name="Materia", value_name="Nota")
                                d.columns = ['Alumno', 'Materia', 'Nota']
                            df_t = d
                        except: pass
                    elif f.name.endswith('.pdf'):
                        txt = extract_table_data_from_pdf(f)
                        if txt: df_t = process_data_with_ai(txt, api_key, f.name)
                    elif 'doc' in f.name:
                        txt = extract_text_from_docx(f)
                        if txt: df_t = process_data_with_ai(txt, api_key, f.name)
                    
                    if df_t is not None and not df_t.empty: dfs.append(df_t)
                    bar.progress((i+1)/len(uploaded_files))
                
                if dfs:
                    st.session_state.data = pd.concat(dfs, ignore_index=True)
                    st.rerun()
                else: st.error("No se extrajeron datos.")

    if st.session_state.data is not None:
        if st.button("🔄 Subir otro archivo (Reiniciar)"): reiniciar_app()

st.title("Acta de Evaluación")
col_b1, col_b2, col_b3 = st.columns([1,1,1])
col_b1.info(f"🏫 **Centro:** {centro}")
col_b2.info(f"👥 **Grupo:** {grupo}")
col_b3.info(f"📅 **Curso:** {curso}")

if st.session_state.data is not None:
    # --- CORRECCIÓN CRÍTICA DE COLUMNAS (SOLUCIÓN KEYERROR) ---
    # Limpiar espacios en nombres de columnas
    st.session_state.data.columns = st.session_state.data.columns.str.strip()
    
    # Renombrar si la IA devolvió nombres distintos
    cols_map = {
        'Student': 'Alumno', 'Nombre': 'Alumno', 'Apellidos y Nombre': 'Alumno',
        'Subject': 'Materia', 'Asignatura': 'Materia',
        'Grade': 'Nota', 'Calificacion': 'Nota'
    }
    st.session_state.data.rename(columns=cols_map, inplace=True)

    # Verificar existencia de columnas antes de seguir
    required_cols = ['Alumno', 'Materia', 'Nota']
    missing = [c for c in required_cols if c not in st.session_state.data.columns]
    
    if missing:
        st.error(f"❌ Error: La IA no encontró las columnas correctas. Faltan: {missing}. Intenta subir el archivo de nuevo.")
    else:
        # Aquí ya es seguro ejecutar el código
        df = st.session_state.data.drop_duplicates(subset=['Alumno', 'Materia'], keep='last')
        df['Nota'] = pd.to_numeric(df['Nota'], errors='coerce')
        df['Aprobado'] = df['Nota'] >= 5
        
        # CÁLCULOS
        stats_al = df.groupby('Alumno').agg(
            Suspensos=('Nota', lambda x: (x<5).sum()),
            Media=('Nota', 'mean')
        ).reset_index()
        
        stats_mat = df.groupby('Materia').agg(
            Total=('Nota', 'count'),
            Aprobados=('Aprobado', 'sum'),
            Suspensos=('Nota', lambda x: (x<5).sum()),
            Media=('Nota', 'mean')
        ).reset_index()
        stats_mat['Pct_Aprobados'] = (stats_mat['Aprobados']/stats_mat['Total'])*100
        stats_mat['Pct_Suspensos'] = (stats_mat['Suspensos']/stats_mat['Total'])*100
        
        total_alumnos = len(stats_al)
        media_grupo = df['Nota'].mean()
        max_media = stats_al['Media'].max()
        min_media = stats_al['Media'].min()
        mejores = stats_al[stats_al['Media'] == max_media]['Alumno'].tolist()
        peores = stats_al[stats_al['Media'] == min_media]['Alumno'].tolist()
        
        cero = stats_al[stats_al['Suspensos'] == 0].shape[0]
        uno = stats_al[stats_al['Suspensos'] == 1].shape[0]
        dos = stats_al[stats_al['Suspensos'] == 2].shape[0]
        mas_dos = stats_al[stats_al['Suspensos'] > 2].shape[0]
        
        base = total_alumnos if total_alumnos > 0 else 1
        pasan = cero + uno + dos
        no_pasan = mas_dos
        
        res = {
            'total_alumnos': total_alumnos, 'media_grupo': media_grupo,
            'max_media': max_media, 'min_media': min_media,
            'mejores_alumnos': mejores, 'peores_alumnos': peores,
            'cero': cero, 'pct_cero': (cero/base)*100,
            'uno': uno, 'pct_uno': (uno/base)*100,
            'dos': dos, 'pct_dos': (dos/base)*100,
            'mas_dos': mas_dos, 'pct_mas_dos': (mas_dos/base)*100,
            'pasan': pasan, 'pct_pasan': (pasan/base)*100,
            'no_pasan': no_pasan, 'pct_no_pasan': (no_pasan/base)*100,
            'valoracion': obtener_valoracion((pasan/base)*100)
        }

        # VISUALIZACIÓN
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Informe General", "📚 Por Materias", "🎓 Por Alumnos", "📄 Descargas"])
        
        with tab1:
            st.subheader("1. Datos Generales")
            st.write(f"**a) Grupo:** {grupo} | **b) Alumnos:** {total_alumnos} | **c) Media:** {media_grupo:.2f}")
            st.write(f"**d) Mejor Media:** {', '.join(mejores)} ({max_media:.2f})")
            st.write(f"**e) Peor Media:** {', '.join(peores)} ({min_media:.2f})")
            
            st.markdown("---")
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Todo Aprobado", f"{cero} ({res['pct_cero']:.1f}%)")
            c2.metric("1 Suspenso", f"{uno} ({res['pct_uno']:.1f}%)")
            c3.metric("2 Suspensos", f"{dos} ({res['pct_dos']:.1f}%)")
            c4.metric(">2 Suspensos", f"{mas_dos} ({res['pct_mas_dos']:.1f}%)", delta_color="inverse")
            
            st.success(f"✅ PROMOCIONAN: {pasan} ({res['pct_pasan']:.1f}%)")
            st.info(f"📝 Valoración: {res['valoracion']}")
            
            g1, g2 = st.columns(2)
            with g1:
                fig, ax = plt.subplots()
                ax.pie([pasan, no_pasan], labels=['Sí', 'No'], autopct='%1.1f%%', colors=['#2ecc71', '#e74c3c'])
                ax.set_title("Promoción")
                st.pyplot(fig)
            with g2:
                fig2, ax2 = plt.subplots()
                ax2.bar(['0','1','2','>2'], [cero, uno, dos, mas_dos], color='#3498db')
                ax2.set_title("Suspensos")
                st.pyplot(fig2)

        with tab2:
            st.dataframe(stats_mat.style.format({'Pct_Aprobados':'{:.1f}%','Pct_Suspensos':'{:.1f}%','Media':'{:.2f}'}), use_container_width=True)
            fig3, ax3 = plt.subplots(figsize=(10,4))
            ax3.bar(stats_mat['Materia'], stats_mat['Media'], color='#9b59b6')
            ax3.set_title("Nota Media por Materia")
            st.pyplot(fig3)

        with tab3:
            pivot = df.pivot_table(index='Alumno', columns='Materia', values='Nota', aggfunc='first')
            st.dataframe(pivot)

        with tab4:
            img_buf = io.BytesIO(); fig.savefig(img_buf, format='png'); img_buf.seek(0)
            img_buf2 = io.BytesIO(); fig2.savefig(img_buf2, format='png'); img_buf2.seek(0)
            img_buf3 = io.BytesIO(); fig3.savefig(img_buf3, format='png'); img_buf3.seek(0)
            
            st.download_button("📄 Informe General (Word)", generate_global_report(res, [img_buf, img_buf2, img_buf3], stats_mat, centro, grupo), f"Informe_{grupo}.docx", type="primary")
            st.markdown("---")
            c_izq, c_der = st.columns(2)
            with c_izq:
                sel = st.selectbox("Individual", stats_al['Alumno'].unique())
                if sel:
                    inf = stats_al[stats_al['Alumno']==sel].iloc[0]
                    st.download_button(f"Descargar {sel}", crear_informe_individual(sel, df[df['Alumno']==sel], inf['Media'], inf['Suspensos'], stats_mat), f"{sel}.docx")
            with c_der:
                if st.button("Generar Todos"):
                    st.download_button("Descargar ZIP Todos", generar_informe_todos_alumnos(df, stats_al, stats_mat), f"Todos_{grupo}.docx")
else:
    st.info("👈 Sube las actas en el menú lateral.")
