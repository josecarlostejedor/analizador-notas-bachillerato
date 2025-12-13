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
    /* Resaltar celdas editables */
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

# --- FUNCIONES DE EXTRACCIÓN ---
def extract_text_with_pdfplumber(file):
    text_content = ""
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text_content += page.extract_text(x_tolerance=2, y_tolerance=2) + "\n"
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
    
    # PROMPT MEJORADO PARA EVITAR CONFUSIÓN CON ÍNDICES (1, 2, 3...)
    prompt = f"""
    Analiza el siguiente texto de un acta de evaluación ('{filename}').
    
    ATENCIÓN AL FORMATO:
    1. Al principio aparece una lista de alumnos. CUIDADO: Delante del nombre suele haber un número índice (1, 2, 3...). NO confundas ese número con una nota.
       Ejemplo: "1 ANTHONY..." -> El '1' es el índice, ignóralo. La nota viene después o al final.
    2. Las NOTAS (números) suelen aparecer AL FINAL DEL BLOQUE DE TEXTO, separadas de los nombres.
    3. Tu tarea es ASOCIAR la primera fila de notas al primer alumno, la segunda al segundo, etc.
    
    TAREA:
    Genera un CSV con columnas: "Alumno", "Materia", "Nota".
    REGLAS:
    - Materia: Usa abreviaturas (EF, ING1, etc).
    - Nota: Numérico (0-10). Si ves un número > 10, probablemente sea un código, ignóralo.
    - Devuelve SOLO el CSV.
    
    Texto:
    {text_data[:15000]}
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

# --- GENERACIÓN DE TEXTOS AUTOMÁTICOS ---
def generar_comentario_individual(alumno, datos_alumno):
    suspensos = datos_alumno[datos_alumno['Nota'] < 5]
    num_suspensos = len(suspensos)
    lista_suspensas = suspensos['Materia'].tolist()
    
    txt = f"El alumno tiene actualmente {num_suspensos} materias suspensas."
    
    if num_suspensos == 0:
        txt = "El alumno no tiene ninguna materia suspensa. ¡Excelente trabajo! Se recomienda mantener la constancia en el estudio y, si es posible, ayudar a compañeros en las materias donde destaca."
    elif num_suspensos == 1:
        txt += f" La materia pendiente es: {', '.join(lista_suspensas)}. "
        txt += "Al ser solo una materia, la recuperación es muy factible. Se recomienda hablar con el profesor de la asignatura para establecer un plan de trabajo específico y reforzar los contenidos base."
    elif num_suspensos == 2:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. "
        txt += "Se encuentra en el límite de la promoción. Es vital organizar un horario de estudio que priorice estas dos asignaturas sin descuidar el resto. Se aconseja asistencia a refuerzos."
    else:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. "
        txt += "La situación es preocupante y compromete la promoción al curso siguiente. Se requiere un cambio radical en los hábitos de estudio, supervisión familiar diaria y tutorías urgentes con el equipo docente."
    return txt

def generar_valoracion_detallada(res):
    txt = f"El grupo presenta una nota media global de {res['media_grupo']:.2f}. "
    if res['pct_pasan'] >= 85:
        txt += "El nivel de promoción es excelente, con la gran mayoría del alumnado superando los objetivos. Esto indica un grupo con buena dinámica de trabajo. "
    elif res['pct_pasan'] >= 70:
        txt += "El nivel de promoción es satisfactorio. La mayoría del grupo avanza adecuadamente, aunque existe un sector que requiere seguimiento. "
    else:
        txt += "El nivel de promoción es bajo, lo que alerta sobre dificultades generalizadas en el aprendizaje o adaptación al curso. "
    
    if res['pct_mas_dos'] > 20:
        txt += f"Preocupa especialmente que un {res['pct_mas_dos']:.1f}% de alumnos acumula más de dos suspensos. "
    elif res['pct_cero'] > 50:
        txt += "Destaca positivamente que más de la mitad de la clase ha aprobado todas las materias. "
        
    txt += "Se recomienda mantener la comunicación con las familias de los alumnos con dificultades y reforzar las materias con medias más bajas."
    return txt

# --- FUNCIONES DE WORD ---
def add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat):
    doc.add_heading(f'Informe Individual: {alumno}', 0)
    doc.add_paragraph(f"Nota Media: {media:.2f} | Materias Suspensas: {suspensos}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('Análisis y Recomendaciones', level=2)
    comentario = generar_comentario_individual(alumno, datos_alumno)
    p = doc.add_paragraph(comentario)
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
            c[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255,0,0)
            c[1].paragraphs[0].runs[0].bold = True

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
    p2.add_run(f"• Promocionan (0-2 suspensos): {datos_resumen['pasan']} ({datos_resumen['pct_pasan']:.1f}%)\n").bold = True
    p2.add_run(f"• No promocionan (>2 suspensos): {datos_resumen['no_pasan']} ({datos_resumen['pct_no_pasan']:.1f}%)").bold = True

    doc.add_heading('3. Valoración General del Grupo', 1)
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
    t_graf = doc.add_table(rows=2, cols=2)
    t_graf.autofit = True
    r1 = t_graf.rows[0].cells
    r1[0].paragraphs[0].add_run().add_picture(plots[0], width=Inches(4.5))
    r1[1].paragraphs[0].add_run().add_picture(plots[3], width=Inches(4.5))
    r2 = t_graf.rows[1].cells
    r2[0].paragraphs[0].add_run().add_picture(plots[2], width=Inches(4.5))
    r2[1].paragraphs[0].add_run().add_picture(plots[1], width=Inches(4.5))

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
                        txt = extract_text_with_pdfplumber(f)
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
    # Corrección de columnas
    st.session_state.data.columns = st.session_state.data.columns.str.strip()
    cols_map = {'Student': 'Alumno', 'Nombre': 'Alumno', 'Apellidos y Nombre': 'Alumno', 'Subject': 'Materia', 'Asignatura': 'Materia', 'Grade': 'Nota'}
    st.session_state.data.rename(columns=cols_map, inplace=True)

    if 'Alumno' not in st.session_state.data.columns:
        st.error("❌ Error de columnas.")
    else:
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
        mejores = stats_al[stats_al['Media'] == stats_al['Media'].max()]['Alumno'].tolist()
        peores = stats_al[stats_al['Media'] == stats_al['Media'].min()]['Alumno'].tolist()
        
        cero = stats_al[stats_al['Suspensos'] == 0].shape[0]
        uno = stats_al[stats_al['Suspensos'] == 1].shape[0]
        dos = stats_al[stats_al['Suspensos'] == 2].shape[0]
        mas_dos = stats_al[stats_al['Suspensos'] > 2].shape[0]
        
        base = total_alumnos if total_alumnos > 0 else 1
        pasan = cero + uno + dos
        no_pasan = mas_dos
        
        res = {
            'total_alumnos': total_alumnos, 'media_grupo': media_grupo,
            'mejores_alumnos': mejores, 'peores_alumnos': peores,
            'cero': cero, 'pct_cero': (cero/base)*100,
            'uno': uno, 'pct_uno': (uno/base)*100,
            'dos': dos, 'pct_dos': (dos/base)*100,
            'mas_dos': mas_dos, 'pct_mas_dos': (mas_dos/base)*100,
            'pasan': pasan, 'pct_pasan': (pasan/base)*100,
            'no_pasan': no_pasan, 'pct_no_pasan': (no_pasan/base)*100,
        }
        res['valoracion'] = generar_valoracion_detallada(res)

        # VISUALIZACIÓN
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Informe General", "📚 Por Materias", "🎓 Por Alumnos (Editor)", "📄 Descargas"])
        
        with tab1:
            st.subheader("Datos Generales")
            c1,c2,c3 = st.columns(3)
            c1.metric("Alumnos", total_alumnos); c2.metric("Media Grupo", f"{media_grupo:.2f}")
            c3.metric("Promoción", f"{res['pct_pasan']:.1f}%")
            
            st.markdown(f"**Valoración:** _{res['valoracion']}_")
            
            g1, g2 = st.columns(2)
            with g1:
                fig, ax = plt.subplots(figsize=(5,4))
                ax.pie([pasan, no_pasan], labels=['Promocionan', 'No'], autopct='%1.1f%%', colors=['#2ecc71', '#e74c3c'], startangle=90)
                ax.set_title("Ratio de Promoción")
                st.pyplot(fig)
            with g2:
                fig2, ax2 = plt.subplots(figsize=(5,4))
                bars = ax2.bar(['0', '1', '2', '>2'], [cero, uno, dos, mas_dos], color='#3498db')
                ax2.bar_label(bars)
                ax2.set_title("Distribución de Suspensos")
                st.pyplot(fig2)
                
            st.write("---")
            fig4, ax4 = plt.subplots(figsize=(8, 3))
            labels_prom = ['Promocionan', 'No Promocionan']
            counts_prom = [pasan, no_pasan]
            pcts_prom = [res['pct_pasan'], res['pct_no_pasan']]
            bars4 = ax4.bar(labels_prom, counts_prom, color=['#2ecc71', '#e74c3c'])
            for bar, pct in zip(bars4, pcts_prom):
                height = bar.get_height()
                ax4.text(bar.get_x() + bar.get_width()/2., height, f'{int(height)}\n({pct:.1f}%)', ha='center', va='bottom')
            ax4.set_title("Detalle Promoción")
            st.pyplot(fig4)

        with tab2:
            st.dataframe(stats_mat.style.format({'Pct_Aprobados':'{:.1f}%','Media':'{:.2f}'}), use_container_width=True)
            fig3, ax3 = plt.subplots(figsize=(10,5))
            datos_graf = stats_mat.sort_values('Media', ascending=False)
            bars3 = ax3.bar(datos_graf['Materia'], datos_graf['Media'], color='#9b59b6')
            ax3.set_ylim(0, 10.5)
            ax3.set_title("Nota Media por Materia")
            ax3.bar_label(bars3, fmt='%.2f', padding=3)
            plt.xticks(rotation=45)
            st.pyplot(fig3)

        with tab3:
            st.markdown("### 📝 Editor de Calificaciones")
            st.info("Haz doble clic en una celda para corregir un dato. Al terminar, pulsa el botón rojo abajo.")
            
            # CREAR TABLA PIVOT EDITABLE
            pivot_df = df.pivot_table(index='Alumno', columns='Materia', values='Nota', aggfunc='first')
            edited_df = st.data_editor(pivot_df, use_container_width=True, num_rows="dynamic")
            
            # BOTÓN DE RECALCULAR
            if st.button("🔄 Recalcular Análisis con Datos Corregidos", type="primary"):
                # Convertir de vuelta a formato largo (Alumno, Materia, Nota)
                try:
                    new_long_df = edited_df.reset_index().melt(id_vars='Alumno', var_name='Materia', value_name='Nota')
                    new_long_df.dropna(subset=['Nota'], inplace=True) # Eliminar vacíos
                    st.session_state.data = new_long_df
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar datos: {e}")

        with tab4:
            plots = []
            for f in [fig, fig2, fig3, fig4]:
                buf = io.BytesIO(); f.savefig(buf, format='png', bbox_inches='tight'); buf.seek(0)
                plots.append(buf)
            
            st.download_button("📄 Informe General (Word)", generate_global_report(res, plots, stats_mat, centro, grupo), f"Informe_{grupo}.docx", type="primary")
            
            c_izq, c_der = st.columns(2)
            with c_izq:
                sel = st.selectbox("Individual", stats_al['Alumno'].unique())
                if sel:
                    inf = stats_al[stats_al['Alumno']==sel].iloc[0]
                    st.info(generar_comentario_individual(sel, df[df['Alumno']==sel]))
                    st.download_button(f"Descargar {sel}", crear_informe_individual(sel, df[df['Alumno']==sel], inf['Media'], inf['Suspensos'], stats_mat), f"{sel}.docx")
            with c_der:
                if st.button("Generar Todos"):
                    st.download_button("Descargar ZIP Todos", generar_informe_todos_alumnos(df, stats_al, stats_mat), f"Todos_{grupo}.docx")
else:
    st.info("👈 Sube las actas en el menú lateral.")
