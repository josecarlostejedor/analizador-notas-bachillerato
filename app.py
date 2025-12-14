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
    
    prompt = f"""
    Analiza el siguiente texto de un acta de evaluación ('{filename}').
    
    ATENCIÓN AL FORMATO:
    1. Al principio aparece una lista de alumnos. CUIDADO: Delante del nombre suele haber un número índice (1, 2, 3...). NO confundas ese número con una nota.
    2. Las NOTAS (números) suelen aparecer AL FINAL DEL BLOQUE DE TEXTO.
    3. Asocia la primera fila de notas al primer alumno, etc.
    
    TAREA:
    Genera un CSV con columnas: "Alumno", "Materia", "Nota".
    REGLAS:
    - Materia: Usa abreviaturas (EF, ING1, etc).
    - Nota: Numérico (0-10).
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
        txt = "El alumno no tiene ninguna materia suspensa. ¡Excelente trabajo! Se recomienda mantener la constancia en el estudio."
    elif num_suspensos == 1:
        txt += f" La materia pendiente es: {', '.join(lista_suspensas)}. Recuperación factible, se recomienda plan específico."
    elif num_suspensos == 2:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. Situación límite para la promoción. Se aconseja refuerzo urgente."
    else:
        txt += f" Las materias son: {', '.join(lista_suspensas)}. Situación preocupante que compromete la promoción."
    return txt

def generar_valoracion_detallada(res):
    txt = f"El grupo presenta una nota media global de {res['media_grupo']:.2f}. "
    if res['pct_pasan'] >= 85:
        txt += "El nivel de promoción es excelente."
    elif res['pct_pasan'] >= 70:
        txt += "El nivel de promoción es satisfactorio."
    else:
        txt += "El nivel de promoción es bajo."
    return txt

# --- FUNCIONES DE WORD ---
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
    s = doc.sections[0]; s.orientation = WD_ORIENT.LANDSCAPE; s.page_width, s.page_height = s.page_height, s.page_width
    
    doc.add_heading(f'Informe de Evaluación - {centro}', 0)
    doc.add_heading('Datos Generales', 1)
    doc.add_paragraph(f"Media del grupo: {datos_resumen['media_grupo']:.2f}")
    doc.add_paragraph(f"Promoción: {datos_resumen['pasan']} ({datos_resumen['pct_pasan']:.1f}%)")
    doc.add_paragraph(f"No Promocionan: {datos_resumen['no_pasan']} ({datos_resumen['pct_no_pasan']:.1f}%)")
    
    doc.add_heading('Gráficas', 1)
    # Tabla 2x2 para las gráficas
    t = doc.add_table(rows=2, cols=2)
    t.autofit = True
    
    # Insertar imágenes asegurando que existen
    if len(plots) >= 4:
        t.rows[0].cells[0].paragraphs[0].add_run().add_picture(plots[0], width=Inches(4.5))
        t.rows[0].cells[1].paragraphs[0].add_run().add_picture(plots[3], width=Inches(4.5))
        t.rows[1].cells[0].paragraphs[0].add_run().add_picture(plots[2], width=Inches(4.5))
        t.rows[1].cells[1].paragraphs[0].add_run().add_picture(plots[1], width=Inches(4.5))

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# --- NUEVA FUNCIÓN: INFORME PADRES ---
def generate_parents_report(res, stats_mat, plot_suspensos, plot_pct_materias):
    doc = Document()
    
    # Configurar Apaisado
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    # Título
    doc.add_heading('RESUMEN DE EVALUACIÓN PARA FAMILIAS', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Tabla Estructural (1 Fila, 2 Columnas)
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(5)
    table.columns[1].width = Inches(5)
    
    # --- COLUMNA 1: TEXTO ---
    cell_text = table.rows[0].cells[0]
    p = cell_text.paragraphs[0]
    p.add_run("En esta gráfica vemos el número de materias suspensas en este trimestre, así como una estimación de cuántos alumnos promocionarían.\n\n").italic = True
    
    p.add_run(f"• Alumnos que promocionan: {res['pasan']} ({res['pct_pasan']:.1f}%)\n")
    p.add_run(f"• Alumnos que no promocionan: {res['no_pasan']} ({res['pct_no_pasan']:.1f}%)\n")
    p.add_run(f"• Media de suspensos del grupo: {res['media_suspensos_grupo']:.2f}\n\n")
    
    p.add_run("Porcentaje de aprobados por materia:\n").bold = True
    for _, row in stats_mat.iterrows():
        p.add_run(f"- {row['Materia']}: {row['Pct_Aprobados']:.1f}%\n")
    
    # Más y menos aprobados
    mejor = stats_mat.loc[stats_mat['Pct_Aprobados'].idxmax()]
    peor = stats_mat.loc[stats_mat['Pct_Aprobados'].idxmin()]
    
    p.add_run(f"\nAsignatura con menos aprobados: {peor['Materia']} ({peor['Pct_Aprobados']:.1f}%)\n").bold = True
    p.add_run(f"Asignatura con más aprobados: {mejor['Materia']} ({mejor['Pct_Aprobados']:.1f}%)\n").bold = True

    # --- COLUMNA 2: GRÁFICAS ---
    cell_graphs = table.rows[0].cells[1]
    
    p_g1 = cell_graphs.paragraphs[0]
    p_g1.add_run("1º) Materias no superadas (Distribución):\n").bold = True
    p_g1.add_run().add_picture(plot_suspensos, width=Inches(4.5))
    
    p_g2 = cell_graphs.add_paragraph()
    p_g2.add_run("\n2º) % de Suspensos por Materia:\n").bold = True
    p_g2.add_run().add_picture(plot_pct_materias, width=Inches(4.5))

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# --- INTERFAZ ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=50)
    st.title("Configuración")
    api_key = st.text_input("🔑 API Key OpenAI", type="password")
    st.markdown("---")
    centro = st.text_input("Centro", "IES Lucía de Medrano")
    grupo = st.text_input("Grupo", "1º BACH 4")
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
        tres = stats_al[stats_al['Suspensos'] == 3].shape[0]
        mas_tres = stats_al[stats_al['Suspensos'] > 3].shape[0]
        
        base = total_alumnos if total_alumnos > 0 else 1
        pasan = cero + uno + dos
        no_pasan = tres + mas_tres
        media_suspensos_grupo = stats_al['Suspensos'].mean()
        
        res = {
            'total_alumnos': total_alumnos, 'media_grupo': media_grupo,
            'media_suspensos_grupo': media_suspensos_grupo,
            'pasan': pasan, 'pct_pasan': (pasan/base)*100,
            'no_pasan': no_pasan, 'pct_no_pasan': (no_pasan/base)*100,
            'pct_mas_dos': ((tres+mas_tres)/base)*100
        }
        res['valoracion'] = generar_valoracion_detallada(res)

        # TABS
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Informe General", "📚 Por Materias", "🎓 Por Alumnos (Editor)", "📄 Informes", "👨‍👩‍👧 Resumen Padres"])
        
        # 1. GENERAL
        with tab1:
            st.metric("Media Grupo", f"{media_grupo:.2f}")
            c1, c2 = st.columns(2)
            with c1: 
                fig_pie, ax_pie = plt.subplots(figsize=(4,3))
                ax_pie.pie([pasan, no_pasan], labels=['Sí', 'No'], autopct='%1.1f%%', colors=['#2ecc71', '#e74c3c'])
                st.pyplot(fig_pie)
            with c2:
                fig_bars, ax_bars = plt.subplots(figsize=(4,3))
                ax_bars.bar(['0','1','2','3','>3'], [cero, uno, dos, tres, mas_tres], color='#3498db')
                st.pyplot(fig_bars)

        # 2. MATERIAS
        with tab2:
            st.dataframe(stats_mat.style.format({'Pct_Aprobados':'{:.1f}%'}), use_container_width=True)

        # 3. EDITOR
        with tab3:
            st.markdown("### 📝 Editor de Calificaciones")
            pivot_df = df.pivot_table(index='Alumno', columns='Materia', values='Nota', aggfunc='first')
            edited_df = st.data_editor(pivot_df, use_container_width=True, num_rows="dynamic")
            if st.button("🔄 Recalcular", type="primary"):
                try:
                    new_long_df = edited_df.reset_index().melt(id_vars='Alumno', var_name='Materia', value_name='Nota')
                    new_long_df.dropna(subset=['Nota'], inplace=True)
                    st.session_state.data = new_long_df
                    st.rerun()
                except: pass

        # 4. INFORMES INDIVIDUALES Y GENERAL (CORREGIDO)
        with tab4:
            st.write("Generador de informes.")
            
            # Regenerar gráficas generales para el Word
            # 1. Pie
            fig_p, ax_p = plt.subplots(figsize=(5,4))
            ax_p.pie([pasan, no_pasan], labels=['Sí', 'No'], autopct='%1.1f%%', colors=['#2ecc71', '#e74c3c'], startangle=90)
            
            # 2. Distribución suspensos
            fig_d, ax_d = plt.subplots(figsize=(5,4))
            bars_d = ax_d.bar(['0', '1', '2', '>2'], [cero, uno, dos, tres+mas_tres], color='#3498db')
            ax_d.bar_label(bars_d)

            # 3. Media Notas
            fig_m, ax_m = plt.subplots(figsize=(10,5))
            d_graf = stats_mat.sort_values('Media', ascending=False)
            bars_m = ax_m.bar(d_graf['Materia'], d_graf['Media'], color='#9b59b6')
            ax_m.bar_label(bars_m, fmt='%.2f')
            plt.xticks(rotation=45)

            # 4. Promocion vs No
            fig_pr, ax_pr = plt.subplots(figsize=(8,3))
            ax_pr.bar(['Sí', 'No'], [pasan, no_pasan], color=['green', 'red'])

            plots_general = []
            for f in [fig_p, fig_d, fig_m, fig_pr]:
                buf = io.BytesIO()
                f.savefig(buf, format='png', bbox_inches='tight')
                buf.seek(0)
                plots_general.append(buf)

            if st.button("Generar Informe General Word"):
                st.download_button("Descargar", generate_global_report(res, plots_general, stats_mat, centro, grupo), f"Global_{grupo}.docx")

        # 5. RESUMEN PADRES
        with tab5:
            st.header("Resumen para Reunión de Padres")
            
            # Gráfica 1: Distribución suspensos
            fig_padres1, ax_p1 = plt.subplots(figsize=(6, 4))
            labels_p = ['0', '1', '2', '3', '>3']
            vals_p = [cero, uno, dos, tres, mas_tres]
            bars_p = ax_p1.bar(labels_p, vals_p, color=['#2ecc71', '#f1c40f', '#e67e22', '#e74c3c', '#c0392b'])
            ax_p1.bar_label(bars_p)
            ax_p1.set_title("Alumnos según nº de materias suspensas")
            ax_p1.set_ylabel("Nº Alumnos")
            
            # Gráfica 2: % Suspensos por materia
            fig_padres2, ax_p2 = plt.subplots(figsize=(6, 4))
            df_p2 = stats_mat.sort_values('Pct_Suspensos', ascending=True)
            ax_p2.barh(df_p2['Materia'], df_p2['Pct_Suspensos'], color='#3498db')
            ax_p2.set_title("% de Suspensos por Materia")
            ax_p2.set_xlabel("% Suspensos")
            
            c1, c2 = st.columns(2)
            with c1: st.pyplot(fig_padres1)
            with c2: st.pyplot(fig_padres2)
            
            buf_p1 = io.BytesIO(); fig_padres1.savefig(buf_p1, format='png', bbox_inches='tight'); buf_p1.seek(0)
            buf_p2 = io.BytesIO(); fig_padres2.savefig(buf_p2, format='png', bbox_inches='tight'); buf_p2.seek(0)
            
            if st.button("📄 Generar Word Resumen Padres"):
                doc_padres = generate_parents_report(res, stats_mat, buf_p1, buf_p2)
                st.download_button(
                    label="⬇️ Descargar Resumen Padres (.docx)",
                    data=doc_padres,
                    file_name=f"Resumen_Padres_{grupo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
else:
    st.info("👈 Sube las actas en el menú lateral.")
