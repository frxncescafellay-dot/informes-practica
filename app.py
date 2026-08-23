import os
import io
import json
from datetime import datetime
import pytz
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
from google import genai
from google.genai import types
import pypdf
from docx import Document
from pptx import Presentation

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Plataforma de Analítica & Intervención IA",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- RUTAS DE ALMACENAMIENTO PERSISTENTE ---
DIR_BASE = "almacen_datos"
DIR_DATASETS = os.path.join(DIR_BASE, "datasets")
DIR_INTERVENCION = os.path.join(DIR_BASE, "intervencion")
DIR_AVATARS = os.path.join(DIR_BASE, "avatares")
FILE_DB = os.path.join(DIR_BASE, "base_datos.json")

for d in [DIR_BASE, DIR_DATASETS, DIR_INTERVENCION, DIR_AVATARS]:
    os.makedirs(d, exist_ok=True)

MODELO_GEMINI = "gemini-3.6-flash"

# --- ESTILOS VISUALES TEAL SUAVE INTEGRAL (SIN NEGRO NI BLANCO PURO) ---
st.markdown("""
<style>
    /* Fondo principal: Teal Profundo */
    .stApp {
        background: radial-gradient(circle at 50% 30%, #007f8c 0%, #005f69 55%, #003e45 100%) !important;
        color: #001e28;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    }
    
    /* Encabezados y títulos */
    h1, h2, h3, h4, h5, h6 {
        color: #001e28 !important;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    
    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #5c9e9b !important;
        border-right: 2px solid #3d7976 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #001e28 !important;
        font-weight: 600;
    }

    /* Tarjetas y Contenedores */
    .modern-card {
        background: #83c5be !important;
        border: 1.5px solid #63a9a2 !important;
        border-radius: 14px;
        padding: 22px;
        margin-bottom: 20px;
        box-shadow: 0 8px 20px rgba(0, 30, 40, 0.25);
    }

    /* Desplegables (Expanders): corrección de barra negra superior */
    div[data-testid="stExpander"] {
        background-color: #83c5be !important;
        border: 1.5px solid #63a9a2 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stExpander"] summary {
        background-color: #5c9e9b !important;
        color: #001e28 !important;
        border-radius: 8px !important;
        font-weight: 750 !important;
    }
    div[data-testid="stExpander"] summary * {
        color: #001e28 !important;
        font-weight: 750 !important;
    }
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #83c5be !important;
        color: #001e28 !important;
    }

    /* Campos de texto y entradas */
    input[type="text"], 
    input[type="password"], 
    textarea {
        background-color: #a5d8d3 !important;
        color: #001e28 !important;
        border: 1.5px solid #63a9a2 !important;
        border-radius: 8px !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
    }
    input[type="text"]:focus, 
    input[type="password"]:focus, 
    textarea:focus {
        background-color: #c4ece8 !important;
        color: #00141a !important;
        border-color: #e76f51 !important;
        box-shadow: 0 0 0 3px rgba(231, 111, 81, 0.3) !important;
    }

    /* Selectores */
    div[data-baseweb="select"] > div {
        background-color: #a5d8d3 !important;
        color: #001e28 !important;
        border: 1.5px solid #63a9a2 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] * {
        color: #001e28 !important;
        font-weight: 600 !important;
    }

    /* Zona de carga de archivos (Uploader): corrección de botón negro */
    div[data-testid="stFileUploader"] {
        background-color: #83c5be !important;
        border: 2px dashed #005f69 !important;
        border-radius: 12px !important;
        padding: 12px !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: #83c5be !important;
    }
    div[data-testid="stFileUploader"] section * {
        color: #001e28 !important;
    }
    div[data-testid="stFileUploader"] button {
        background-color: #4a8b88 !important;
        color: #001e28 !important;
        border: 1.5px solid #2b5f5c !important;
        font-weight: 750 !important;
    }
    div[data-testid="stFileUploader"] button * {
        color: #001e28 !important;
    }

    /* Tablas de datos Excel / Data Editor: corrección de fondo negro */
    div[data-testid="stDataFrame"], 
    div[data-testid="stDataEditor"],
    div[data-testid="stDataFrame"] > div,
    div[data-testid="stDataEditor"] > div {
        background-color: #83c5be !important;
        border: 1.5px solid #3d7976 !important;
        border-radius: 10px !important;
    }
    
    /* Celdas del canvas de Glide Data Grid (Excel de Streamlit) */
    .glide-data-grid,
    .gdg-container,
    div[data-testid="stDataFrame"] canvas,
    div[data-testid="stDataEditor"] canvas {
        filter: invert(0.88) hue-rotate(180deg) saturate(1.8) contrast(1.1) !important;
        border-radius: 8px !important;
    }

    /* Botones primarios (Tono Coral) */
    .stButton>button {
        background: #e76f51 !important;
        color: #001e28 !important;
        border: 1.5px solid #cf5436 !important;
        border-radius: 8px !important;
        padding: 8px 20px !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button:hover {
        background: #f4a261 !important;
        color: #00141a !important;
        border-color: #e76f51 !important;
        transform: translateY(-1px);
    }
    
    /* Botones secundarios (Borrar / Eliminar) */
    div[data-testid="stBaseButton-secondary"] button {
        background: #e63946 !important;
        color: #ffffff !important;
        border: 1.5px solid #ba181b !important;
    }
    div[data-testid="stBaseButton-secondary"] button:hover {
        background: #ba181b !important;
        color: #ffffff !important;
    }

    /* Pestañas de navegación */
    button[data-baseweb="tab"] {
        background-color: #5c9e9b !important;
        color: #001e28 !important;
        border-radius: 8px 8px 0 0 !important;
        font-weight: 750 !important;
        margin-right: 4px !important;
        padding: 9px 18px !important;
        border: 1px solid #3d7976 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #83c5be !important;
        color: #001e28 !important;
        border-bottom: 3.5px solid #e76f51 !important;
    }

    label, p, span {
        color: #001e28;
        font-weight: 600;
    }

    .badge-role {
        display: inline-block;
        background: #e76f51;
        color: #001e28 !important;
        padding: 4px 12px;
        border-radius: 14px;
        font-size: 0.82rem;
        font-weight: 800;
        border: 1px solid #cf5436;
    }
</style>
""", unsafe_allow_html=True)

# --- COMPONENTE DE RELOJ DUAL (DIGITAL + ANÁLOGO SVG) ---
def renderizar_reloj_chile():
    html_reloj = """
    <div style="
        background: #83c5be;
        border: 1.5px solid #63a9a2;
        border-radius: 14px;
        padding: 10px 18px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        box-shadow: 0 4px 14px rgba(0, 30, 40, 0.25);
        max-width: 480px;
        margin-left: auto;
        margin-bottom: 15px;
        font-family: 'Segoe UI', system-ui, sans-serif;
    ">
        <!-- Reloj Análogo -->
        <div style="position: relative; width: 58px; height: 58px;">
            <svg id="analog-clock" width="58" height="58" viewBox="0 0 100 100">
                <circle cx="50" cy="50" r="46" fill="#a5d8d3" stroke="#005f69" stroke-width="4"/>
                <line x1="50" y1="10" x2="50" y2="16" stroke="#003e45" stroke-width="3" stroke-linecap="round"/>
                <line x1="90" y1="50" x2="84" y2="50" stroke="#003e45" stroke-width="3" stroke-linecap="round"/>
                <line x1="50" y1="90" x2="50" y2="84" stroke="#003e45" stroke-width="3" stroke-linecap="round"/>
                <line x1="10" y1="50" x2="16" y2="50" stroke="#003e45" stroke-width="3" stroke-linecap="round"/>
                <line id="hour-hand" x1="50" y1="50" x2="50" y2="28" stroke="#001e28" stroke-width="4.5" stroke-linecap="round"/>
                <line id="min-hand" x1="50" y1="50" x2="50" y2="18" stroke="#005f69" stroke-width="3" stroke-linecap="round"/>
                <line id="sec-hand" x1="50" y1="50" x2="50" y2="14" stroke="#e76f51" stroke-width="2" stroke-linecap="round"/>
                <circle cx="50" cy="50" r="3.5" fill="#001e28"/>
            </svg>
        </div>
        <!-- Reloj Digital y Fecha -->
        <div style="display: flex; flex-direction: column; justify-content: center;">
            <div style="font-size: 0.75rem; font-weight: 800; color: #003e45; text-transform: uppercase; letter-spacing: 0.5px;">
                🇨🇱 Hora Oficial de Chile
            </div>
            <div id="digital-clock" style="font-size: 1.5rem; font-weight: 900; color: #001e28; line-height: 1.1;">
                --:--:--
            </div>
            <div id="digital-date" style="font-size: 0.78rem; font-weight: 700; color: #003e45; margin-top: 2px;">
                Cargando fecha...
            </div>
        </div>
    </div>

    <script>
        function actualizarReloj() {
            const opciones = { timeZone: 'America/Santiago', hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' };
            const opcionesFecha = { timeZone: 'America/Santiago', weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
            
            const ahora = new Date();
            const formateadorHora = new Intl.DateTimeFormat('es-CL', opciones);
            const formateadorFecha = new Intl.DateTimeFormat('es-CL', opcionesFecha);
            
            const partesHora = formateadorHora.formatToParts(ahora);
            let h = 0, m = 0, s = 0;
            partesHora.forEach(p => {
                if (p.type === 'hour') h = parseInt(p.value);
                if (p.type === 'minute') m = parseInt(p.value);
                if (p.type === 'second') s = parseInt(p.value);
            });

            const hStr = h.toString().padStart(2, '0');
            const mStr = m.toString().padStart(2, '0');
            const sStr = s.toString().padStart(2, '0');
            document.getElementById('digital-clock').textContent = `${hStr}:${mStr}:${sStr}`;
            
            let fechaStr = formateadorFecha.format(ahora);
            fechaStr = fechaStr.charAt(0).toUpperCase() + fechaStr.slice(1);
            document.getElementById('digital-date').textContent = fechaStr;

            const secDeg = s * 6;
            const minDeg = m * 6 + s * 0.1;
            const hourDeg = (h % 12) * 30 + m * 0.5;

            function rotar(elemId, deg) {
                const el = document.getElementById(elemId);
                if (el) el.setAttribute('transform', `rotate(${deg} 50 50)`);
            }
            rotar('sec-hand', secDeg);
            rotar('min-hand', minDeg);
            rotar('hour-hand', hourDeg);
        }
        setInterval(actualizarReloj, 1000);
        actualizarReloj();
    </script>
    """
    components.html(html_reloj, height=95)

# --- GESTOR DE PERSISTENCIA (JSON) ---
def cargar_estado():
    if not os.path.exists(FILE_DB):
        data_inicial = {
            "usuarios": {
                "admin1": {
                    "nombre": "Administrador Principal",
                    "rol": "Admin",
                    "pin": "1234",
                    "permiso_editar": True,
                    "permiso_eliminar": True,
                    "avatar": ""
                }
            },
            "datasets": {},
            "intervencion": [],
            "informes": []
        }
        guardar_estado(data_inicial)
        return data_inicial
    with open(FILE_DB, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return {"usuarios": {"admin1": {"nombre": "Administrador Principal", "rol": "Admin", "pin": "1234", "permiso_editar": True, "permiso_eliminar": True, "avatar": ""}}, "datasets": {}, "intervencion": [], "informes": []}

def guardar_estado(data):
    with open(FILE_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

db = cargar_estado()

# --- CLIENTE GENAI ---
API_KEY_GEMINI = os.environ.get("GEMINI_API_KEY", "")

def obtener_cliente_ia():
    if not API_KEY_GEMINI:
        return None
    return genai.Client(api_key=API_KEY_GEMINI)

# --- EXTRACCIÓN DE TEXTO PARA RESÚMENES ---
def extraer_texto_archivo(ruta, extension):
    texto = ""
    try:
        if extension == "pdf":
            lector = pypdf.PdfReader(ruta)
            for pag in lector.pages:
                t = pag.extract_text()
                if t: texto += t + "\n"
        elif extension == "docx":
            doc = Document(ruta)
            texto = "\n".join([p.text for p in doc.paragraphs])
        elif extension == "pptx":
            prs = Presentation(ruta)
            for diap in prs.slides:
                for forma in diap.shapes:
                    if hasattr(forma, "text"):
                        texto += forma.text + "\n"
        elif extension in ["xlsx", "xls"]:
            df_tmp = pd.read_excel(ruta)
            texto = f"Estadísticas:\n{df_tmp.describe(include='all').to_string()}\nPrimeras filas:\n{df_tmp.head(10).to_string()}"
    except Exception as e:
        texto = f"Error al extraer texto: {e}"
    return texto[:8000]

# --- CONTROL DE ACCESO (LOGIN) ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None

if not st.session_state.autenticado:
    st.markdown("<h1 style='text-align:center; color:#83c5be !important; margin-top:35px;'>🔐 Acceso a la Plataforma</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:#a5d8d3;'>Ingrese sus credenciales registradas</p>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.4, 1])
    with c2:
        with st.form("login_form"):
            u_in = st.text_input("Usuario")
            p_in = st.text_input("PIN / Contraseña", type="password")
            if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                if u_in in db["usuarios"] and db["usuarios"][u_in]["pin"] == p_in:
                    st.session_state.autenticado = True
                    st.session_state.usuario_clave = u_in
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
    st.stop()

# --- DATOS DEL USUARIO ACTUAL ---
usr_key = st.session_state.usuario_clave
usr = db["usuarios"].get(usr_key, {"nombre": "Invitado", "rol": "Usuario", "permiso_editar": False, "permiso_eliminar": False, "avatar": ""})
es_admin = (usr.get("rol") == "Admin") or (usr_key == "admin1")

# --- BARRA LATERAL: PERFIL Y FOTO ---
st.sidebar.markdown("### 👤 Mi Perfil")

avatar_path = usr.get("avatar", "")
if avatar_path and os.path.exists(avatar_path):
    st.sidebar.image(avatar_path, width=105)
else:
    st.sidebar.markdown("""
        <div style='width:80px; height:80px; border-radius:50%; background:#a5d8d3; border:2px solid #005f69; display:flex; align-items:center; justify-content:center; font-size:2.2rem; color:#001e28; margin-bottom:10px;'>
            👤
        </div>
    """, unsafe_allow_html=True)

st.sidebar.markdown(f"**{usr.get('nombre')}**")
st.sidebar.markdown(f"<span class='badge-role'>{usr.get('rol')}</span>", unsafe_allow_html=True)

with st.sidebar.expander("📷 Cambiar foto de perfil"):
    nueva_foto = st.file_uploader("Subir imagen (JPG, PNG):", type=["jpg", "jpeg", "png"], key="upload_avatar")
    if nueva_foto is not None:
        if st.button("Guardar Foto"):
            ruta_avatar = os.path.join(DIR_AVATARS, f"{usr_key}_avatar.png")
            with open(ruta_avatar, "wb") as f_av:
                f_av.write(nueva_foto.getbuffer())
            db["usuarios"][usr_key]["avatar"] = ruta_avatar
            guardar_estado(db)
            st.success("Foto actualizada.")
            st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None
    st.rerun()

# --- ENCABEZADO SUPERIOR CON RELOJ ---
col_head_title, col_head_clock = st.columns([1.2, 1])
with col_head_title:
    st.markdown("<h1 style='color:#83c5be !important; margin:0;'>⚡ Centro de Gestión & Analítica</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#a5d8d3; margin-top:4px;'>Plataforma colaborativa multi-formato con procesamiento inteligente.</p>", unsafe_allow_html=True)
with col_head_clock:
    renderizar_reloj_chile()

# --- NAVEGACIÓN MODULAR POR PESTAÑAS ---
titulos_pestanas = ["📊 Datasets & Archivos", "📂 Intervención", "📄 Informes Compartidos"]
if es_admin:
    titulos_pestanas.append("👥 Gestión de Usuarios")

pestanas_principales = st.tabs(titulos_pestanas)

# ==========================================
# 1. DATASETS EXCEL
# ==========================================
with pestanas_principales[0]:
    st.markdown("""
        <div class='modern-card'>
            <h3 style='margin:0;'>📊 Repositorio de Datasets</h3>
            <p style='margin:0; color:#001e28;'>Suba, visualice y edite hojas de cálculo en ventanas independientes.</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.expander("➕ Cargar Nuevo Dataset Excel", expanded=True):
        c_tit, c_arc = st.columns([1, 1])
        with c_tit:
            t_data = st.text_input("Título del dataset:")
        with c_arc:
            f_data = st.file_uploader("Archivo Excel (.xlsx, .xls):", type=["xlsx", "xls"])
            
        if st.button("Guardar Dataset"):
            if not t_data.strip() or f_data is None:
                st.error("Complete el título y seleccione un archivo.")
            else:
                nom_arc = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f_data.name}"
                ruta_dest = os.path.join(DIR_DATASETS, nom_arc)
                with open(ruta_dest, "wb") as f:
                    f.write(f_data.getbuffer())
                
                db["datasets"][t_data] = {
                    "titulo": t_data,
                    "autor": usr["nombre"],
                    "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "ruta": ruta_dest
                }
                guardar_estado(db)
                st.success("Dataset guardado con éxito.")
                st.rerun()

    st.markdown("---")
    if not db["datasets"]:
        st.info("No hay datasets subidos actualmente.")
    else:
        titulos_ds = list(db["datasets"].keys())
        pestanas_ds = st.tabs(titulos_ds)
        
        for idx_ds, tab_ds in enumerate(pestanas_ds):
            t_act = titulos_ds[idx_ds]
            info_ds = db["datasets"][t_act]
            
            with tab_ds:
                st.markdown(f"### 📋 {info_ds['titulo']}")
                st.caption(f"Subido por: **{info_ds['autor']}** | Fecha: {info_ds['fecha']}")
                
                if os.path.exists(info_ds["ruta"]):
                    df_actual = pd.read_excel(info_ds["ruta"])
                    
                    if usr.get("permiso_editar") or es_admin:
                        st.markdown("**✏️ Editor de Datos en Vivo:**")
                        df_edit = st.data_editor(df_actual, key=f"d_edit_{t_act}", use_container_width=True)
                        if st.button("💾 Guardar Cambios en Excel", key=f"s_df_{t_act}"):
                            df_edit.to_excel(info_ds["ruta"], index=False)
                            st.success("Datos actualizados.")
                            st.rerun()
                    else:
                        st.dataframe(df_actual, use_container_width=True)
                    
                    if usr.get("permiso_eliminar") or es_admin:
                        st.markdown("---")
                        if st.button("🗑️ Eliminar Dataset", key=f"del_ds_{t_act}", type="secondary"):
                            if os.path.exists(info_ds["ruta"]):
                                os.remove(info_ds["ruta"])
                            del db["datasets"][t_act]
                            guardar_estado(db)
                            st.success("Dataset eliminado.")
                            st.rerun()
                else:
                    st.error("Archivo físico no encontrado.")

# ==========================================
# 2. INTERVENCIÓN MULTI-FORMATO
# ==========================================
with pestanas_principales[1]:
    st.markdown("""
        <div class='modern-card'>
            <h3 style='margin:0;'>📂 Módulo de Intervención</h3>
            <p style='margin:0; color:#001e28;'>Soporte y resúmenes automáticos para Documentos, Imágenes y Audios (M4A/MP3/WAV).</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.form("form_intervencion"):
        tit_int = st.text_input("Título descriptivo del material:")
        archivos = st.file_uploader(
            "Cargar archivos:",
            type=["xlsx", "xls", "docx", "pdf", "pptx", "jpg", "jpeg", "png", "mp4", "m4a", "mp3", "wav"],
            accept_multiple_files=True
        )
        if st.form_submit_button("Subir y Procesar"):
            if not tit_int.strip() or not archivos:
                st.error("Complete el título y cargue al menos un archivo.")
            else:
                client = obtener_cliente_ia()
                for a in archivos:
                    ext = a.name.split(".")[-1].lower()
                    nom_dest = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{a.name}"
                    ruta_guardada = os.path.join(DIR_INTERVENCION, nom_dest)
                    
                    with open(ruta_guardada, "wb") as f:
                        f.write(a.getbuffer())
                    
                    resumen_txt = "Archivo guardado."
                    if client:
                        with st.spinner(f"Analizando {a.name} con Gemini 3.6 Flash..."):
                            try:
                                if ext in ["m4a", "mp3", "wav"]:
                                    mime_map = {"m4a": "audio/mp4", "mp3": "audio/mp3", "wav": "audio/wav"}
                                    with open(ruta_guardada, "rb") as f_aud:
                                        audio_bytes = f_aud.read()
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=[
                                            types.Part.from_bytes(data=audio_bytes, mime_type=mime_map.get(ext, "audio/mp4")),
                                            "Sintetiza los puntos clave tratados en este audio, acuerdos y diagnóstico para intervención."
                                        ]
                                    )
                                    resumen_txt = resp.text
                                elif ext in ["jpg", "jpeg", "png"]:
                                    img = Image.open(ruta_guardada)
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=["Describe y resume los elementos clave de esta imagen para un informe de intervención:", img]
                                    )
                                    resumen_txt = resp.text
                                elif ext in ["pdf", "docx", "pptx", "xlsx", "xls"]:
                                    t_doc = extraer_texto_archivo(ruta_guardada, ext)
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=f"Elabora un resumen y diagnóstico clave de este documento ({a.name}):\n\n{t_doc}"
                                    )
                                    resumen_txt = resp.text
                                elif ext == "mp4":
                                    resumen_txt = "Video registrado (reproducción multimedia disponible)."
                            except Exception as e:
                                resumen_txt = f"Archivo guardado. Diagnóstico no generado: {e}"
                    
                    db["intervencion"].append({
                        "titulo": tit_int,
                        "nombre_original": a.name,
                        "tipo": ext,
                        "autor": usr["nombre"],
                        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "ruta": ruta_guardada,
                        "resumen": resumen_txt
                    })
                guardar_estado(db)
                st.success("Materiales integrados correctamente.")
                st.rerun()

    st.markdown("---")
    st.subheader("📚 Materiales Guardados")
    
    if not db["intervencion"]:
        st.info("No hay registros en intervención.")
    else:
        for idx, item in enumerate(db["intervencion"]):
            titulo_item = item.get("titulo") or item.get("titulo_registro") or "Sin Título"
            nombre_arc = item.get("nombre_original") or item.get("filename") or "Archivo"
            tipo_arc = item.get("tipo", "").lower()
            autor_arc = item.get("autor", "Desconocido")
            fecha_arc = item.get("fecha", "")
            resumen_arc = item.get("resumen", "Sin resumen disponible.")
            ruta_arc = item.get("ruta", "")
            
            with st.container():
                col_head, col_del_btn = st.columns([5, 1])
                with col_head:
                    st.markdown(f"### 📁 {titulo_item} — `{nombre_arc}`")
                    st.caption(f"Autor: **{autor_arc}** | Fecha: {fecha_arc} | Formato: **{tipo_arc.upper()}**")
                with col_del_btn:
                    if usr.get("permiso_eliminar") or es_admin:
                        if st.button("🗑️ Borrar", key=f"del_int_{idx}", type="secondary"):
                            if ruta_arc and os.path.exists(ruta_arc):
                                os.remove(ruta_arc)
                            db["intervencion"].pop(idx)
                            guardar_estado(db)
                            st.success("Archivo eliminado.")
                            st.rerun()
                
                t_vis, t_res = st.tabs(["👁️ Multimedia / Descarga", "📝 Diagnóstico y Resumen"])
                with t_vis:
                    if ruta_arc and os.path.exists(ruta_arc):
                        if tipo_arc == "mp4":
                            st.video(ruta_arc)
                        elif tipo_arc in ["m4a", "mp3", "wav"]:
                            st.audio(ruta_arc)
                        elif tipo_arc in ["jpg", "jpeg", "png"]:
                            st.image(ruta_arc, use_container_width=True)
                        else:
                            with open(ruta_arc, "rb") as fl:
                                st.download_button("📥 Descargar Archivo", data=fl.read(), file_name=nombre_arc, key=f"dl_a_{idx}")
                    else:
                        st.error("Archivo físico no encontrado.")
                with t_res:
                    st.info(resumen_arc)
                st.markdown("---")

# ==========================================
# 3. INFORMES COMPARTIDOS
# ==========================================
with pestanas_principales[2]:
    st.markdown("""
        <div class='modern-card'>
            <h3 style='margin:0;'>📄 Informes Compartidos</h3>
            <p style='margin:0; color:#001e28;'>Generación y repositorio colaborativo con exportación en formato Carta.</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.expander("🤖 Redactar Nuevo Informe con IA", expanded=False):
        nom_i = st.text_input("Título del Informe:")
        enf_i = st.selectbox("Enfoque:", ["Resumen Ejecutivo", "Diagnóstico Técnico", "Evaluación Estratégica"])
        ins_i = st.text_area("Instrucciones complementarias:")
        
        if st.button("🚀 Generar Informe"):
            client = obtener_cliente_ia()
            if not nom_i.strip():
                st.error("Ingrese un título para el informe.")
            elif not client:
                st.error("API Key de Gemini no configurada.")
            else:
                with st.spinner("Redactando informe consolidado con Gemini 3.6 Flash..."):
                    resumenes_int = "\n".join([f"- {x.get('nombre_original', 'Archivo')}: {x.get('resumen', '')}" for x in db["intervencion"]])
                    
                    prompt = f"""
                    Actúa como especialista analítico senior.
                    Genera un informe estructurado con enfoque '{enf_i}'.
                    Título: {nom_i}
                    Autor solicitante: {usr['nombre']}
                    Instrucciones: {ins_i}
                    Materiales analizados: {resumenes_int if resumenes_int else 'Sin materiales adjuntos.'}
                    
                    Estructura:
                    1. Diagnóstico y Visión Global
                    2. Hallazgos Analíticos Relevantes
                    3. Conclusiones y Plan de Acción
                    """
                    try:
                        resp = client.models.generate_content(
                            model=MODELO_GEMINI,
                            contents=prompt
                        )
                        db["informes"].append({
                            "titulo": nom_i,
                            "autor": usr["nombre"],
                            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "contenido": resp.text
                        })
                        guardar_estado(db)
                        st.success("Informe publicado.")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Error al generar informe: {err}")

    st.markdown("---")
    st.subheader("📚 Repositorio de Informes")
    
    if not db["informes"]:
        st.info("No hay informes registrados.")
    else:
        for idx_inf, inf in enumerate(db["informes"]):
            titulo_inf = inf.get("titulo") or inf.get("titulo_informe") or "Sin Título"
            autor_inf = inf.get("autor", "Desconocido")
            fecha_inf = inf.get("fecha", "")
            contenido_inf = inf.get("contenido", "")
            
            with st.container():
                c_inf_t, c_inf_del = st.columns([5, 1])
                with c_inf_t:
                    st.markdown(f"### 📄 [{autor_inf}] - {titulo_inf}")
                    st.caption(f"Generado el: {fecha_inf}")
                with c_inf_del:
                    if usr.get("permiso_eliminar") or es_admin:
                        if st.button("🗑️ Borrar", key=f"del_inf_{idx_inf}", type="secondary"):
                            db["informes"].pop(idx_inf)
                            guardar_estado(db)
                            st.success("Informe eliminado.")
                            st.rerun()
                
                st.markdown(contenido_inf)
                
                cd1, cd2 = st.columns([1, 1])
                with cd1:
                    df_out = pd.DataFrame([{"Autor": autor_inf, "Título": titulo_inf, "Fecha": fecha_inf, "Contenido": contenido_inf}])
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_out.to_excel(writer, index=False)
                    st.download_button(
                        label="📥 Descargar Excel (.xlsx)",
                        data=buf.getvalue(),
                        file_name=f"[{autor_inf}] - {titulo_inf}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"xls_dl_{idx_inf}"
                    )
                with cd2:
                    html_carta = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>{titulo_inf}</title>
                        <style>
                            @page {{ size: letter portrait; margin: 25mm; }}
                            body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #001e28; padding: 20px; background-color: #83c5be; }}
                            .header {{ border-bottom: 2px solid #005f69; padding-bottom: 8px; margin-bottom: 20px; }}
                            .title {{ font-size: 20pt; font-weight: bold; color: #003e45; }}
                            .meta {{ font-size: 10pt; color: #005f69; margin-top: 4px; }}
                            .body {{ font-size: 11pt; line-height: 1.6; white-space: pre-wrap; }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <div class="title">{titulo_inf}</div>
                            <div class="meta">Autor: {autor_inf} | Fecha: {fecha_inf}</div>
                        </div>
                        <div class="body">{contenido_inf}</div>
                    </body>
                    </html>
                    """
                    st.download_button(
                        label="🖼️ Descargar Formato Carta (HTML / PDF)",
                        data=html_carta,
                        file_name=f"[{autor_inf}] - {titulo_inf}_Carta.html",
                        mime="text/html",
                        key=f"doc_dl_{idx_inf}"
                    )
                st.markdown("---")

# ==========================================
# 4. GESTIÓN DE USUARIOS (SOLO ADMIN)
# ==========================================
if es_admin:
    with pestanas_principales[3]:
        st.markdown("""
            <div class='modern-card'>
                <h3 style='margin:0;'>👥 Gestión de Usuarios y Permisos</h3>
                <p style='margin:0; color:#001e28;'>Control centralizado de cuentas accesible exclusivamente por Administradores.</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.expander("➕ Registrar Nuevo Usuario", expanded=True):
            cu1, cu2, cu3 = st.columns([1, 1, 1])
            with cu1:
                n_u = st.text_input("Usuario (Login):")
                n_nom = st.text_input("Nombre Completo:")
            with cu2:
                n_pin = st.text_input("PIN / Clave:", type="password")
                n_rol = st.selectbox("Rol:", ["Usuario", "Analista", "Especialista", "Gerencia", "Admin"])
            with cu3:
                st.markdown("**Permisos Iniciales:**")
                p_e = st.checkbox("Permiso para Editar Datasets")
                p_d = st.checkbox("Permiso para Eliminar Datasets / Archivos")
                
            if st.button("Crear Usuario"):
                if not n_u or not n_pin or not n_nom:
                    st.error("Todos los campos son obligatorios.")
                elif n_u in db["usuarios"]:
                    st.error("El nombre de usuario ya existe.")
                else:
                    db["usuarios"][n_u] = {
                        "nombre": n_nom,
                        "rol": n_rol,
                        "pin": n_pin,
                        "permiso_editar": p_e,
                        "permiso_eliminar": p_d,
                        "avatar": ""
                    }
                    guardar_estado(db)
                    st.success(f"Usuario '{n_u}' registrado correctamente.")
                    st.rerun()

        st.markdown("---")
        st.subheader("📜 Cuentas Registradas")
        
        for u_k in list(db["usuarios"].keys()):
            u_d = db["usuarios"][u_k]
            with st.container():
                col_u_inf, col_u_e, col_u_d, col_u_del = st.columns([2, 1, 1, 1])
                with col_u_inf:
                    st.markdown(f"**{u_d.get('nombre', '')}** (`{u_k}`) — Rol: *{u_d.get('rol', '')}*")
                
                if u_k == "admin1":
                    st.caption("Administrador Principal (Cuenta protegida)")
                else:
                    with col_u_e:
                        val_e = st.checkbox("Editar", value=u_d.get("permiso_editar", False), key=f"pe_{u_k}")
                    with col_u_d:
                        val_d = st.checkbox("Eliminar", value=u_d.get("permiso_eliminar", False), key=f"pd_{u_k}")
                    with col_u_del:
                        if st.button("🗑️ Eliminar", key=f"del_user_{u_k}", type="secondary"):
                            del db["usuarios"][u_k]
                            guardar_estado(db)
                            st.success(f"Usuario {u_k} eliminado.")
                            st.rerun()
                            
                    if val_e != u_d.get("permiso_editar") or val_d != u_d.get("permiso_eliminar"):
                        db["usuarios"][u_k]["permiso_editar"] = val_e
                        db["usuarios"][u_k]["permiso_eliminar"] = val_d
                        guardar_estado(db)
            st.markdown("---")
