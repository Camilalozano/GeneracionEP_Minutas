import streamlit as st
import pandas as pd
from docx import Document
import zipfile
import io
from datetime import datetime, date
import time
import base64
from pathlib import Path
import uuid

# ============== CONFIGURACIÓN DE PÁGINA ==============
st.set_page_config(
    page_title="ATENEA: Generador Estudios Previos y Minutas",
    page_icon="Generador de Documentos",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============== ESTILOS CSS PERSONALIZADOS ==============
st.markdown("""
<style>
    /* Fuentes y colores principales */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%);
        padding: 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }

    .main-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .subtitle {
        color: rgba(255,255,255,0.8);
        font-size: 1.1rem;
        margin-top: 0.5rem;
    }

    .stat-card {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        border-left: 4px solid #2d5a87;
    }

    .step-badge {
        background: #2d5a87;
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .sidebar-info {
        background: rgba(255,255,255,0.1);
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ============== FUNCIONES ==============
def replace_text_in_paragraph(paragraph, key, value):
    placeholder = "{{" + key + "}}"
    value_str = str(value) if value is not None else ""
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder in full_text:
        replaced_text = full_text.replace(placeholder, value_str)
        if paragraph.runs:
            paragraph.runs[0].text = replaced_text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(replaced_text)


def generar_documentos(df, word_file, progress_bar, status_text):
    try:
        zip_buffer = io.BytesIO()
        documentos_generados = 0
        errores = []
        total = len(df)

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                try:
                    word_file.seek(0)
                    doc = Document(word_file)

                    for paragraph in doc.paragraphs:
                        for key in row.index:
                            replace_text_in_paragraph(paragraph, key, row[key])

                    for table in doc.tables:
                        for row_table in table.rows:
                            for cell in row_table.cells:
                                for paragraph in cell.paragraphs:
                                    for key in row.index:
                                        replace_text_in_paragraph(paragraph, key, row[key])

                    doc_bytes = io.BytesIO()
                    doc.save(doc_bytes)
                    zipf.writestr(f"Documento_{idx + 1}.docx", doc_bytes.getvalue())
                    documentos_generados += 1

                    progress = (idx + 1) / total
                    progress_bar.progress(progress)
                    status_text.text(f"📄 Procesando documento {idx + 1} de {total}...")

                except Exception as e:
                    errores.append(f"Fila {idx + 1}: {str(e)}")

        status_text.text("✅ ¡Proceso completado!")
        return zip_buffer, documentos_generados, errores

    except Exception as e:
        raise Exception(f"Error procesando documentos: {str(e)}")


def validar_formulario(data):
    errores = {}

    if not data["dependencia"]:
        errores["dependencia"] = "La dependencia es obligatoria."

    if not data["objeto"] or len(data["objeto"].strip()) < 20:
        errores["objeto"] = "El objeto es obligatorio y debe tener al menos 20 caracteres."
    elif len(data["objeto"]) > 500:
        errores["objeto"] = "El objeto no puede superar 500 caracteres."

    if data["plazo_inicio"] > data["plazo_fin"]:
        errores["plazo"] = "La fecha de inicio no puede ser posterior a la fecha final."

    if data["modalidad"] not in ["Contratación directa", "Licitación pública", "Selección abreviada", "Concurso de méritos"]:
        errores["modalidad"] = "Selecciona una modalidad válida de la lista cerrada."

    if not data["supervisor"] or len(data["supervisor"].strip()) < 5:
        errores["supervisor"] = "El supervisor es obligatorio y debe tener al menos 5 caracteres."

    if data["valor_estimado"] <= 0:
        errores["valor_estimado"] = "El valor estimado debe ser mayor a cero."

    return errores


def construir_dataframe_desde_formulario(data):
    fila = {
        "dependencia": data["dependencia"],
        "descripción": data["objeto"],
        "plazo_inicio": data["plazo_inicio"].strftime("%Y-%m-%d"),
        "plazo_fin": data["plazo_fin"].strftime("%Y-%m-%d"),
        "plazo": f"{data['plazo_inicio'].strftime('%Y-%m-%d')} a {data['plazo_fin'].strftime('%Y-%m-%d')}",
        "modalidad": data["modalidad"],
        "supervisor": data["supervisor"],
        "valor_estimado": f"{data['valor_estimado']:,.2f}",
        "observaciones": data["observaciones"],
        "fecha_solicitud": data["fecha_solicitud"].strftime("%Y-%m-%d"),
    }
    return pd.DataFrame([fila])


if "form_borrador" not in st.session_state:
    st.session_state.form_borrador = {}
if "df_captura" not in st.session_state:
    st.session_state.df_captura = None

if "resultado_zip" not in st.session_state:
    st.session_state.resultado_zip = None
if "resultado_nombre" not in st.session_state:
    st.session_state.resultado_nombre = ""
if "resultado_generados" not in st.session_state:
    st.session_state.resultado_generados = 0
if "resultado_errores" not in st.session_state:
    st.session_state.resultado_errores = []
if "descarga_automatica_pendiente" not in st.session_state:
    st.session_state.descarga_automatica_pendiente = False
if "auditoria_acciones" not in st.session_state:
    st.session_state.auditoria_acciones = []


def registrar_evento_auditoria(accion, actor, detalle, id_caso="CASO-EN-SESION"):
    st.session_state.auditoria_acciones.append(
        {
            "id_evento": str(uuid.uuid4())[:8],
            "id_caso": id_caso,
            "fecha_hora_utc": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
            "actor": actor if actor else "No especificado",
            "accion": accion,
            "detalle": detalle,
        }
    )


def cargar_plantilla_precargada():
    """Carga la plantilla institucional por defecto desde el repositorio."""
    candidatos = [
        Path("Plantilla.docx"),
        Path("templates/Plantilla.docx"),
    ]

    for ruta in candidatos:
        if ruta.exists() and ruta.is_file():
            return io.BytesIO(ruta.read_bytes()), ruta.name

    return None, None


def disparar_descarga_automatica(zip_bytes, file_name):
    b64 = base64.b64encode(zip_bytes).decode()
    st.markdown(
        f"""
        <a id="auto-download-link" href="data:application/zip;base64,{b64}" download="{file_name}"></a>
        <script>
            const enlace = document.getElementById('auto-download-link');
            if (enlace) {{ enlace.click(); }}
        </script>
        """,
        unsafe_allow_html=True,
    )

# ============== SIDEBAR ==============
with st.sidebar:
    st.markdown("## Generador Inteligente de Documentos")
    st.markdown("---")
    st.markdown("### 📖 Guía Rápida")
    st.markdown("### ℹ️ Información")
    st.markdown("""
    <div class="sidebar-info">
        <p><strong>Versión:</strong> 1.1</p>
        <p><strong>Gerencia de Gestión Corporativa</p>
    </div>
    """, unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title" style="color: white;">ATENEA: Generador Inteligente de Documentos</h1>
    <p class="subtitle">Estudios Previos y Minutas | Gerencia de Gestión Corporativa</p>
</div>
""", unsafe_allow_html=True)

st.markdown("### 📋 Proceso de Generación")
actor_actual = st.text_input(
    "👤 Responsable de la acción (usuario actual)",
    value=st.session_state.get("actor_actual", ""),
    help="Este nombre se registra en la bitácora para control institucional y trazabilidad legal.",
)
st.session_state.actor_actual = actor_actual

col1, col2, col3 = st.columns(3)
for col, paso, titulo, texto in [
    (col1, "Paso 1", "🧾 Capturar Datos", "Diligencia formulario guiado o carga Excel"),
    (col2, "Paso 2", "📝 Cargar Plantilla", "Sube plantilla Word con placeholders"),
    (col3, "Paso 3", "🚀 Generar", "Genera y descarga tus documentos"),
]:
    with col:
        st.markdown(f"""
        <div class="stat-card">
            <span class="step-badge">{paso}</span>
            <h4 style="margin: 1rem 0 0.5rem 0;">{titulo}</h4>
            <p style="color: #64748b; font-size: 0.9rem;">{texto}</p>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 📁 Captura de datos")
modo_captura = st.radio(
    "Selecciona el modo de captura",
    ["Formulario guiado (principal)", "Carga masiva por Excel (secundario)"],
    horizontal=True,
)

excel_file = None
df = None

if modo_captura == "Formulario guiado (principal)":
    st.info("🧭 Diligencia los bloques del formulario. Puedes guardar un borrador temporal en sesión.")

    defaults = st.session_state.form_borrador
    with st.form("formulario_guiado"):
        st.markdown("#### 1) Datos de la dependencia")
        dep_col1, dep_col2 = st.columns(2)
        with dep_col1:
            dependencia = st.selectbox(
                "Dependencia *",
                ["", "GEP", "GGC", "Jurídica", "Financiera", "Talento Humano"],
                index=0 if not defaults.get("dependencia") else ["", "GEP", "GGC", "Jurídica", "Financiera", "Talento Humano"].index(defaults.get("dependencia")),
            )
        with dep_col2:
            fecha_solicitud = st.date_input("Fecha de solicitud *", value=defaults.get("fecha_solicitud", date.today()))

        st.markdown("#### 2) Objeto y modalidad")
        objeto = st.text_area("Objeto contractual *", value=defaults.get("objeto", ""), max_chars=500)
        modalidad = st.selectbox(
            "Modalidad de selección *",
            ["Contratación directa", "Licitación pública", "Selección abreviada", "Concurso de méritos"],
            index=["Contratación directa", "Licitación pública", "Selección abreviada", "Concurso de méritos"].index(defaults.get("modalidad", "Contratación directa")),
        )

        st.markdown("#### 3) Plazo y responsables")
        pcol1, pcol2, pcol3 = st.columns([1, 1, 2])
        with pcol1:
            plazo_inicio = st.date_input("Plazo inicio *", value=defaults.get("plazo_inicio", date.today()))
        with pcol2:
            plazo_fin = st.date_input("Plazo fin *", value=defaults.get("plazo_fin", date.today()))
        with pcol3:
            supervisor = st.text_input("Supervisor *", value=defaults.get("supervisor", ""), max_chars=120)

        st.markdown("#### 4) Valor y observaciones")
        valor_estimado = st.number_input("Valor estimado (COP) *", min_value=0.0, step=100000.0, value=float(defaults.get("valor_estimado", 0.0)))
        observaciones = st.text_area("Observaciones", value=defaults.get("observaciones", ""), max_chars=1000)

        c1, c2 = st.columns(2)
        guardar_borrador = c1.form_submit_button("💾 Guardar borrador")
        cargar_registro = c2.form_submit_button("✅ Usar este registro")

    form_data = {
        "dependencia": dependencia,
        "fecha_solicitud": fecha_solicitud,
        "objeto": objeto,
        "modalidad": modalidad,
        "plazo_inicio": plazo_inicio,
        "plazo_fin": plazo_fin,
        "supervisor": supervisor,
        "valor_estimado": valor_estimado,
        "observaciones": observaciones,
    }
    errores = validar_formulario(form_data)

    if "dependencia" in errores:
        st.error(f"Dependencia: {errores['dependencia']}")
    if "objeto" in errores:
        st.error(f"Objeto contractual: {errores['objeto']}")
    if "plazo" in errores:
        st.error(f"Plazo: {errores['plazo']}")
    if "modalidad" in errores:
        st.error(f"Modalidad: {errores['modalidad']}")
    if "supervisor" in errores:
        st.error(f"Supervisor: {errores['supervisor']}")
    if "valor_estimado" in errores:
        st.error(f"Valor estimado: {errores['valor_estimado']}")

    if guardar_borrador:
        st.session_state.form_borrador = form_data
        registrar_evento_auditoria(
            "Guardar borrador",
            actor_actual,
            "Se guardó el borrador del formulario guiado.",
        )
        st.success("✅ Borrador guardado en la sesión actual.")

    if cargar_registro:
        if errores:
            st.warning("⚠️ Corrige las validaciones antes de usar el registro.")
        else:
            st.session_state.df_captura = construir_dataframe_desde_formulario(form_data)
            registrar_evento_auditoria(
                "Cargar registro",
                actor_actual,
                "Se cargó un registro del formulario guiado para generar documentos.",
            )
            st.success("✅ Registro cargado correctamente para generar documentos.")

    if st.session_state.df_captura is not None:
        df = st.session_state.df_captura

        with st.expander("👀 Vista previa del registro capturado", expanded=True):
            st.dataframe(st.session_state.df_captura, use_container_width=True)

else:
    st.markdown("##### 📊 Datos (Excel)")
    excel_file = st.file_uploader(
        "Arrastra o selecciona tu archivo Excel",
        type="xlsx",
        key="excel",
        help="Archivo Excel con múltiples registros para generación masiva"
    )
    if excel_file:
        try:
            df = pd.read_excel(excel_file)
            registrar_evento_auditoria(
                "Cargar Excel",
                actor_actual,
                f"Se cargó archivo Excel: {excel_file.name} con {len(df)} registros.",
            )
            st.success(f"✅ {excel_file.name}")
        except Exception as e:
            st.error(f"Error al leer Excel: {e}")

st.markdown("##### 📝 Plantilla (Word)")
plantilla_precargada, nombre_plantilla_precargada = cargar_plantilla_precargada()

if plantilla_precargada:
    st.success(f"✅ Plantilla precargada disponible: {nombre_plantilla_precargada}")
else:
    st.warning("⚠️ No se encontró la plantilla precargada. Sube una plantilla para continuar.")

word_file_upload = st.file_uploader(
    "Opcional: sube otra plantilla Word para reemplazar la plantilla por defecto",
    type="docx",
    key="word",
    help="Si no subes archivo, se usará la plantilla institucional precargada."
)

if word_file_upload:
    word_file = word_file_upload
    registrar_evento_auditoria(
        "Cargar plantilla personalizada",
        actor_actual,
        f"Se seleccionó plantilla Word personalizada: {word_file.name}.",
    )
    st.success(f"✅ Plantilla personalizada seleccionada: {word_file.name}")
else:
    word_file = plantilla_precargada
    if word_file:
        st.info("ℹ️ Se usará la plantilla precargada por defecto.")

if df is not None and word_file:
    with st.expander("👀 Vista previa de datos", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Documentos a generar", len(df))
        with col2:
            st.metric("📊 Columnas/Variables", len(df.columns))
        with col3:
            st.metric("🔤 Placeholders detectados", len(df.columns))
        st.dataframe(df, use_container_width=True, height=220)
        placeholders = " | ".join([f"`{{{{{col}}}}}`" for col in df.columns])
        st.code(placeholders, language=None)

    generate_btn = st.button("🚀 Generar Documentos", use_container_width=True, type="primary")
    if generate_btn:
        if df.empty:
            st.error("❌ No hay datos para procesar")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            zip_buffer, generados, errores = generar_documentos(df, word_file, progress_bar, status_text)
            time.sleep(0.5)
            if generados > 0:
                zip_buffer.seek(0)
                zip_data = zip_buffer.getvalue()
                nombre_zip = f"Documentos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

                st.session_state.resultado_zip = zip_data
                st.session_state.resultado_nombre = nombre_zip
                st.session_state.resultado_generados = generados
                st.session_state.resultado_errores = errores
                st.session_state.descarga_automatica_pendiente = True
                registrar_evento_auditoria(
                    "Generar documentos",
                    actor_actual,
                    f"Se generaron {generados} documentos y {len(errores)} errores.",
                )

                st.success("✅ Documentos generados correctamente. Iniciando descarga automática...")
            else:
                st.error("❌ No se pudieron generar documentos")

if st.session_state.resultado_zip:
    st.download_button(
        label=f"📥 Descargar resultados ({st.session_state.resultado_generados} documentos)",
        data=st.session_state.resultado_zip,
        file_name=st.session_state.resultado_nombre,
        mime="application/zip",
        use_container_width=True,
        type="secondary"
    )

    if st.session_state.descarga_automatica_pendiente:
        disparar_descarga_automatica(st.session_state.resultado_zip, st.session_state.resultado_nombre)
        registrar_evento_auditoria(
            "Descarga automática",
            actor_actual,
            f"Se disparó descarga automática del archivo {st.session_state.resultado_nombre}.",
        )
        st.session_state.descarga_automatica_pendiente = False

    if st.session_state.resultado_errores:
        with st.expander(f"⚠️ Ver {len(st.session_state.resultado_errores)} errores"):
            for error in st.session_state.resultado_errores:
                st.warning(error)

st.markdown("---")
st.markdown("### 🧾 Bitácora de auditoría del caso")
st.caption(
    "Registro de acciones para control institucional, responsabilidad legal y reconstrucción histórica del proceso."
)

if st.session_state.auditoria_acciones:
    df_auditoria = pd.DataFrame(st.session_state.auditoria_acciones)
    st.dataframe(df_auditoria, use_container_width=True, height=220)
    st.download_button(
        label="📥 Descargar bitácora de auditoría (CSV)",
        data=df_auditoria.to_csv(index=False).encode("utf-8"),
        file_name=f"auditoria_caso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=True,
    )
else:
    st.info("Aún no hay acciones registradas en esta sesión.")
