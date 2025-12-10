import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
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

# --- LÓGICA DE ANÁLISIS CUALITATIVO (NUEVA) ---

def generar_texto_analisis(alumno, datos_alumno, stats_mat):
    """Genera un texto automático analizando al alumno vs la clase"""
    
    # Preparar datos
    notas_alumno = datos_alumno.set_index('Materia')['Nota']
    medias_clase = stats_mat.set_index('Materia')['Media']
    
    # Comparativa
    comparativa = notas_alumno - medias_clase
    
    # Mejores y peores
    if notas_alumno.empty: return "No hay datos suficientes."
    
    mejor_materia = notas_alumno.idxmax()
    mejor_nota = notas_alumno.max()
    peor_materia = notas_alumno.idxmin()
    peor_nota = notas_alumno.min()
    
    suspensos = notas_alumno[notas_alumno < 5].index.tolist()
    num_suspensos = len(suspensos)
    
    # Construcción del texto
    texto = []
    
    # 1. Análisis de rendimiento
    if num_suspensos == 0:
        texto.append(f"El alumno {alumno} ha tenido un rendimiento excelente, aprobando todas las materias.")
        texto.append(f"Destaca especialmente en {mejor_materia} con un {mejor_nota}.")
    elif num_suspensos <= 2:
        texto.append(f"El alumno presenta un buen rendimiento general, aunque necesita reforzar {', '.join(suspensos)}.")
        texto.append(f"Su mejor resultado ha sido en {mejor_materia} ({mejor_nota}).")
    else:
        texto.append(f"El alumno presenta dificultades significativas, con {num_suspensos} materias insuficientes ({', '.join(suspensos)}).")
        texto.append("Es necesario un plan de recuperación urgente.")

    # 2. Comparativa con la media
    materias_top = comparativa[comparativa > 0].index.tolist()
    if materias_top:
        texto.append(f"Supera la media de la clase en {len(materias_top)} asignaturas, destacando su desempeño por encima del promedio.")
    else:
        texto.append("Actualmente se encuentra por debajo de la media del grupo en las materias evaluadas.")

    # 3. Recomendaciones (Consejos)
    texto.append("\nRecomendaciones:")
    if num_suspensos > 0:
        texto.append(f"- Priorizar el estudio de {peor_materia}, donde se registra la calificación más baja ({peor_nota}).")
        texto.append("- Se recomienda asistir a tutorías de refuerzo y revisar los conceptos base de las materias no superadas.")
        if num_suspensos > 2:
            texto.append("- Es aconsejable una reunión con la familia para establecer un horario de estudio supervisado.")
    else:
        texto.append("- Mantener la constancia actual.")
        if mejor_nota < 9:
            texto.append(f"- Intentar subir nota en {peor_materia} participando más en clase.")
        else:
            texto.append("- ¡Felicidades! Continúa ayudando a compañeros en tus materias fuertes.")

    return " ".join(texto)

# --- FUNCIONES DE WORD MODIFICADAS ---

def add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat):
    """Añade página completa de alumno con análisis y comparativa"""
    
    # Título
    doc.add_heading(f'Informe Individual: {alumno}', 0)
    p_info = doc.add_paragraph(f"Nota Media: {media:.2f} | Materias Suspensas: {suspensos}")
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # --- SECCIÓN 1: ANÁLISIS CUALITATIVO ---
    doc.add_heading('Análisis y Recomendaciones', level=2)
    texto_analisis = generar_texto_analisis(alumno, datos_alumno, stats_mat)
    p_analisis = doc.add_paragraph(texto_analisis)
    p_analisis.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # --- SECCIÓN 2: TABLA COMPARATIVA ---
    doc.add_heading('Comparativa de Rendimiento', level=2)
    
    t = doc.add_table(rows=1, cols=4)
    t.style = 'Table Grid'
    t.autofit = False 
    
    # Anchos
    t.columns[0].width = Inches(2.5) # Materia
    t.columns[1].width = Inches(1.2) # Nota
    t.columns[2].width = Inches(1.2) # Media Clase
    t.columns[3].width = Inches(1.5) # Situación
    
    # Cabeceras
    hdr = t.rows[0].cells
    hdr[0].text = 'Materia'
    hdr[1].text = 'Nota Alumno'
    hdr[2].text = 'Media Clase'
    hdr[3].text = 'Diferencia'
    
    # Datos para comparar
    medias_dict = stats_mat.set_index('Materia')['Media'].to_dict()
    
    for _, row in datos_alumno.iterrows():
        materia = row['Materia']
        nota = row['Nota']
        media_clase = medias_dict.get(materia, 0)
        diferencia = nota - media_clase
        
        c = t.add_row().cells
        c[0].text = str(materia)
        c[1].text = str(nota)
        c[2].text = f"{media_clase:.2f}"
        
        # Lógica de colores y texto para la diferencia
        if diferencia > 0:
            c[3].text = f"+{diferencia:.2f} (Superior)"
            c[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0) # Verde
        else:
            c[3].text = f"{diferencia:.2f} (Inferior)"
            c[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0) # Rojo
            
        # Resaltar suspensos del alumno en columna Nota
        if nota < 5:
            c[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
            c[1].paragraphs[0].runs[0].bold = True

def crear_informe_individual(alumno, datos_alumno, media, suspensos, stats_mat):
    doc = Document()
    add_alumno_to_doc(doc, alumno, datos_alumno, media, suspensos, stats_mat)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generar_informe_todos_alumnos(df, stats_al, stats_mat):
    doc = Document()
    alumnos_lista = stats_al['Alumno'].unique()
    
    for i, alumno in enumerate(alumnos_lista):
        datos_alumno = df[df['Alumno'] == alumno]
        info_alumno = stats_al[stats_al['Alumno'] == alumno].iloc[0]
        
        # Añadir contenido del alumno PASANDO stats_mat para la comparativa
        add_alumno_to_doc(doc, alumno, datos_alumno, info_alumno['Media'], info_alumno['Suspensos'], stats_mat)
        
        # Añadir salto de página si no es el último alumno
        if i < len(alumnos_lista) - 1:
            doc.add_page_break()
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_global_report(res, plots):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    doc.add_heading('Informe de Evaluación - IES Lucía de Medrano', 0)
    doc.add_paragraph('Tutor: Jose Carlos Tejedor')
    
    doc.add_heading('Resumen Ejecutivo', 1)
    p = doc.add_paragraph()
    p.add_run(f"Total Alumnos: {res['total']}\n").bold = True
    p.add_run(f"Promoción: {res['pasan']} alumnos ({res['pct_pasan']:.1f}%)\n")
    p.add_run(f"No Promoción: {res['no_pasan']} alumnos")

    doc.add_heading('Estadísticas por Materia', 1)
    t = doc.add_table(1, 4)
    t.style = 'Table Grid'
    h = t.rows[0].cells
    h[0].text='Materia'; h[1].text='Suspensos'; h[2].text='% Susp'; h[3].text='Media'
    for _, row in res['ranking'].iterrows():
        c = t.add_row().cells
        c[0].text=str(row['Materia']); c[1].text=str(row['Suspensos'])
        c[2].text=f"{row['Pct_Suspensos']:.1f}%"; c[3].text=f"{row['Media']:.2f}"
    
    doc.add_heading('Gráficas', 1)
    for p in plots:
        doc.add_picture(p, width=Inches(6))
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFAZ BARRA LATERAL ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=50)
    st.title("Configuración")
    api_key = st.text_input("🔑 API Key OpenAI", type="password")
    
    st.markdown("---")
    
    centro = st.text_input("Centro", "IES Lucía de Medrano")
    grupo = st.text_input("Grupo", "1º BACH 7")
    curso = st.text_input("Curso", "2024-2025")
    
    st.markdown("---")
    uploaded_files = st.file_uploader(
        "📂 Subir Actas", 
        type=['xlsx', 'pdf', 'docx', 'doc'], 
        accept_multiple_files=True,
        key=f"up_{st.session_state.uploader_key}"
    )
    
    if uploaded_files and st.session_state.data is None:
        if st.button("Analizar Archivos", type="primary"):
            if not api_key:
                st.error("Falta la API Key")
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
                        txt = extract_text_from_pdf(f)
                        if txt: df_t = process_data_with_ai(txt, api_key, f.name)
                    elif 'doc' in f.name:
                        txt = extract_text_from_docx(f)
                        if txt: df_t = process_data_with_ai(txt, api_key, f.name)
                    
                    if df_t is not None and not df_t.empty: dfs.append(df_t)
                    bar.progress((i+1)/len(uploaded_files))
                
                if dfs:
                    st.session_state.data = pd.concat(dfs, ignore_index=True)
                    st.rerun()
                else:
                    st.error("No se extrajeron datos.")

    if st.session_state.data is not None:
        if st.button("🔄 Subir otro archivo (Reiniciar)"):
            reiniciar_app()

# --- ÁREA PRINCIPAL ---
st.title("Acta de Evaluación")
# Badges superiores
col_b1, col_b2, col_b3 = st.columns([1,1,1])
col_b1.info(f"🏫 **Centro:** {centro}")
col_b2.info(f"👥 **Grupo:** {grupo}")
col_b3.info(f"📅 **Curso:** {curso}")

if st.session_state.data is not None:
    # --- LIMPIEZA DATOS ---
    df = st.session_state.data.drop_duplicates(subset=['Alumno', 'Materia'], keep='last')
    df['Nota'] = pd.to_numeric(df['Nota'], errors='coerce')
    
    # CÁLCULOS
    stats_al = df.groupby('Alumno').agg(
        Suspensos=('Nota', lambda x: (x<5).sum()),
        Media=('Nota', 'mean')
    ).reset_index()
    
    stats_mat = df.groupby('Materia').agg(
        Total=('Nota', 'count'),
        Suspensos=('Nota', lambda x: (x<5).sum()),
        Media=('Nota', 'mean')
    ).reset_index()
    stats_mat['Pct_Suspensos'] = (stats_mat['Suspensos']/stats_mat['Total'])*100
    
    total_alumnos = len(stats_al)
    pasan = stats_al[stats_al['Suspensos'] <= 2].shape[0]
    no_pasan = total_alumnos - pasan
    pct_pasan = (pasan/total_alumnos)*100 if total_alumnos > 0 else 0
    
    if not stats_mat.empty:
        peor_materia = stats_mat.loc[stats_mat['Suspensos'].idxmax()]
    else:
        peor_materia = pd.Series({'Materia': 'N/A', 'Suspensos': 0})
    
    # --- PESTAÑAS ---
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Informe General", "📚 Por Materias", "🎓 Por Alumnos", "📄 Informes Individuales"])
    
    # 1. INFORME GENERAL
    with tab1:
        st.markdown(f"""
        <div class="highlight-box">
            <h4>📄 Resumen Ejecutivo del Análisis</h4>
            <p>Se ha analizado el acta del grupo <b>{grupo}</b>. El grupo consta de <b>{total_alumnos}</b> alumnos evaluados.</p>
            <p>En términos de rendimiento global, <b>{pasan} alumnos ({pct_pasan:.1f}%)</b> cumplen los requisitos de promoción 
            (0, 1 o 2 materias suspensas).</p>
        </div>
        """, unsafe_allow_html=True)
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Alumnos Totales", total_alumnos)
        kpi2.metric("Tasa Promoción", f"{pct_pasan:.1f}%")
        kpi3.metric("Media del Grupo", f"{df['Nota'].mean():.2f}")
        kpi4.metric("Suspensos Totales", int(stats_mat['Suspensos'].sum()))
        
        st.markdown("---")
        
        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Promoción del Alumnado")
            fig, ax = plt.subplots()
            ax.pie([pasan, no_pasan], labels=['Promocionan', 'No Promocionan'], autopct='%1.1f%%', colors=['#00CC96', '#EF553B'])
            st.pyplot(fig)
            
        with g2:
            st.subheader("Distribución de Suspensos")
            fig2, ax2 = plt.subplots()
            conteos = stats_al['Suspensos'].value_counts().sort_index()
            ax2.bar(conteos.index.astype(str), conteos.values, color='#636EFA')
            st.pyplot(fig2)

        img_buf = io.BytesIO(); fig.savefig(img_buf, format='png'); img_buf.seek(0)
        img_buf2 = io.BytesIO(); fig2.savefig(img_buf2, format='png'); img_buf2.seek(0)
        
        res_global = {
            'total': total_alumnos, 'pasan': pasan, 'pct_pasan': pct_pasan, 
            'no_pasan': no_pasan, 'ranking': stats_mat.sort_values('Pct_Suspensos', ascending=False)
        }
        word_global = generate_global_report(res_global, [img_buf, img_buf2])
        st.download_button("📥 Descargar Informe Completo (Word)", word_global, "Informe_Global.docx")

    # 2. POR MATERIAS
    with tab2:
        st.subheader("Análisis Detallado por Asignatura")
        col_m1, col_m2 = st.columns([2, 1])
        with col_m1:
            st.dataframe(stats_mat.sort_values('Suspensos', ascending=False), use_container_width=True)
        with col_m2:
            st.info(f"📉 **Más difícil:** {peor_materia['Materia']} ({peor_materia['Suspensos']} suspensos)")
            if not stats_mat.empty:
                mejor_mat = stats_mat.loc[stats_mat['Media'].idxmax()]
                st.success(f"📈 **Mejor media:** {mejor_mat['Materia']} ({mejor_mat['Media']:.2f})")

    # 3. POR ALUMNOS
    with tab3:
        st.subheader("Listado de Calificaciones")
        st.dataframe(stats_al.sort_values('Suspensos'), use_container_width=True)
        
        st.subheader("Detalle de Notas (Todos)")
        pivot = df.pivot_table(index='Alumno', columns='Materia', values='Nota', aggfunc='first')
        st.dataframe(pivot)

    # 4. INFORMES INDIVIDUALES (MEJORADO)
    with tab4:
        st.subheader("🖨️ Boletines con Análisis y Consejos")
        
        c_izq, c_der = st.columns(2)
        
        # --- INDIVIDUAL ---
        with c_izq:
            st.markdown("### 👤 Alumno Individual")
            alumno_sel = st.selectbox("Selecciona un alumno:", stats_al['Alumno'].unique())
            
            if alumno_sel:
                datos_alumno = df[df['Alumno'] == alumno_sel]
                info_alumno = stats_al[stats_al['Alumno'] == alumno_sel].iloc[0]
                
                # Vista previa del análisis
                with st.expander("Ver análisis generado"):
                    st.write(generar_texto_analisis(alumno_sel, datos_alumno, stats_mat))
                
                word_indiv = crear_informe_individual(alumno_sel, datos_alumno, info_alumno['Media'], info_alumno['Suspensos'], stats_mat)
                
                st.download_button(
                    label=f"⬇️ Descargar solo {alumno_sel}",
                    data=word_indiv,
                    file_name=f"Informe_{alumno_sel}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        
        # --- MASIVO ---
        with c_der:
            st.markdown("### 🏫 Toda la Clase")
            st.info("Genera un documento único con el análisis detallado, recomendaciones y comparativas de cada alumno.")
            
            if st.button("🚀 Generar Informe Masivo"):
                with st.spinner("Analizando y generando informes..."):
                    word_todos = generar_informe_todos_alumnos(df, stats_al, stats_mat)
                    
                    st.success("¡Documento listo!")
                    st.download_button(
                        label="⬇️ Descargar TODOS los alumnos (.docx)",
                        data=word_todos,
                        file_name=f"Boletines_Analisis_{grupo}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )

else:
    st.info("👈 Por favor, sube las actas en el menú lateral para ver el análisis.")
