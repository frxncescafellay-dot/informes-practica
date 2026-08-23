import os
import io
import json
from datetime import datetime
import pytz
import pandas as pd
import streamlit as st
from PIL import Image
from google import genai
from google.genai import types
import pypdf
from docx import Document
from pptx import Presentation

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Plataforma de Intervención & Analítica IA",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CREACIÓN DE DIRECTORIOS PARA PERSISTENCIA ---
DIR_BASE = "almacen_datos"
DIR_DATASETS = os.path.join(DIR_BASE, "datasets")
DIR_INTERVENCION = os.path.join(DIR_BASE, "intervencion")
FILE_DB = os.path.join(DIR_BASE, "base_datos.json")

for d in [DIR_BASE, DIR_DATASETS, DIR_INTERVENCION]:
    os.makedirs(d, exist_ok=True)

# --- MODELO DESIGNADO ---
MODELO_GEMINI = "gemini-3.6-flash"

# --- ESTILOS VISUALES MEJORADOS ---
st.markdown("""
<style>
    .stApp {
        background: radial-gradient(circle at 15% 15%, #0f172a 0%, #030712 100%);
        color: #f8fafc;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .top-bar {
        display: flex;
        justify-content: flex-end;
        align-items: center;
        padding: 8px 16px;
        background: rgba(30, 41, 59, 0.4);
        border: 1px solid rgba(56, 189, 248, 0.2);
        border-radius: 10px;
        margin-bottom: 20px;
        font-size: 0.9rem;
        color: #38bdf8;
    }
    
    .header-card {
        background: linear-gradient(135deg, rgba(14, 116, 144, 0.25) 0%, rgba(3, 105, 161, 0.15) 100%);
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 16px;
        padding: 24px;
        display: flex;
        align-items: center;
        gap: 20px;
        margin-bottom: 25px;
    }
    .header-img {
        width: 65px;
        height: 65px;
        filter: drop-shadow(0 0 8px rgba(56, 189, 248, 0.6));
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        transition: 0.3s all;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%);
        color: #030712;
        box-shadow: 0 0 14px rgba(56, 189, 248, 0.5);
    }
    
    section[data-testid="stSidebar"] {
        background-color: #030712 !important;
        border-right: 1px solid #1e293b;
    }
</style>
""", unsafe_allow_html=True)

# --- COMPONENTE HORA CHILE ---
def obtener_fecha_hora_chile():
    tz_cl = pytz.timezone("America/Santiago")
    hora_cl = datetime.now(tz_cl)
    return hora_cl.strftime("🇨🇱 Chile: %d/%m/%Y | %H:%M:%S")

st.markdown(f"<div class='top-bar'>🕒 {obtener_fecha_hora_chile()}</div>", unsafe_allow_html=True)

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
                    "permiso_eliminar": True
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
            return {"usuarios": {"admin1": {"nombre": "Administrador Principal", "rol": "Admin", "pin": "1234", "permiso_editar": True, "permiso_eliminar": True}}, "datasets": {}, "intervencion": [], "informes": []}

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

# --- EXTRACCIÓN DE TEXTO PARA DOCUMENTOS ---
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

# --- CONTROL DE ACCESO ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None

if not st.session_state.autenticado:
    st.markdown("<h1 style='text-align:center;'>🔐 Acceso a la Plataforma</h1>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
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

usr_key = st.session_state.usuario_clave
usr = db["usuarios"].get(usr_key, {"nombre": "Invitado", "rol": "Invitado", "permiso_editar": False, "permiso_eliminar": False})

# --- SIDEBAR DE NAVEGACIÓN ---
st.sidebar.markdown(f"### 👤 {usr['nombre']}")
st.sidebar.caption(f"Rol: **{usr['rol']}**")
st.sidebar.markdown("---")

menu = ["📊 Datasets & Archivos", "📂 Intervención", "📄 Informes Compartidos"]
if usr_key == "admin1" or usr.get("rol") == "Admin":
    menu.append("👥 Gestión de Usuarios")

opcion = st.sidebar.radio("Navegación:", menu)

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None
    st.rerun()

# ==========================================
# 1. DATASETS EXCEL
# ==========================================
if opcion == "📊 Datasets & Archivos":
    st.markdown("""
    <div class='header-card'>
        <img class='header-img' src='https://cdn-icons-png.flaticon.com/512/732/732220.png'>
        <div>
            <h2 style='margin:0;'>Módulo de Datasets & Tablas</h2>
            <p style='margin:0; color:#94a3b8;'>Gestión persistente, edición y visualización por pestañas individuales.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("➕ Cargar Nuevo Archivo Excel", expanded=True):
        c_tit, c_arc = st.columns([1, 1])
        with c_tit:
            t_data = st.text_input("Título de la ventana/dataset:")
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
        titulos = list(db["datasets"].keys())
        pestanas = st.tabs(titulos)
        
        for idx, tab in enumerate(pestanas):
            t_act = titulos[idx]
            info = db["datasets"][t_act]
            
            with tab:
                st.markdown(f"### 📋 {info['titulo']}")
                st.caption(f"Subido por: **{info['autor']}** ({info['fecha']})")
                
                if os.path.exists(info["ruta"]):
                    df_actual = pd.read_excel(info["ruta"])
                    
                    if usr.get("permiso_editar") or usr_key == "admin1":
                        st.markdown("**✏️ Editor de Datos en Vivo:**")
                        df_edit = st.data_editor(df_actual, key=f"d_edit_{t_act}", use_container_width=True)
                        if st.button("💾 Guardar Cambios en Excel", key=f"s_df_{t_act}"):
                            df_edit.to_excel(info["ruta"], index=False)
                            st.success("Datos actualizados.")
                            st.rerun()
                    else:
                        st.dataframe(df_actual, use_container_width=True)
                    
                    if usr.get("permiso_eliminar") or usr_key == "admin1":
                        st.markdown("---")
                        if st.button("🗑️ Eliminar Dataset Completo", key=f"del_ds_{t_act}", type="secondary"):
                            if os.path.exists(info["ruta"]):
                                os.remove(info["ruta"])
                            del db["datasets"][t_act]
                            guardar_estado(db)
                            st.success("Dataset eliminado.")
                            st.rerun()
                else:
                    st.error("Archivo físico no encontrado.")

# ==========================================
# 2. INTERVENCIÓN MULTI-FORMATO (CON AUDIO)
# ==========================================
elif opcion == "📂 Intervención":
    st.markdown("""
    <div class='header-card'>
        <img class='header-img' src='https://cdn-icons-png.flaticon.com/512/3135/3135715.png'>
        <div>
            <h2 style='margin:0;'>Módulo de Intervención Multi-Formato</h2>
            <p style='margin:0; color:#94a3b8;'>Soporte y resúmenes automáticos para Documentos, Imágenes y Audios (M4A/MP3/WAV).</p>
        </div>
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
                        with st.spinner(f"Analizando y resumiendo {a.name} con Gemini..."):
                            try:
                                # 1. Audio (M4A, MP3, WAV)
                                if ext in ["m4a", "mp3", "wav"]:
                                    mime_map = {"m4a": "audio/mp4", "mp3": "audio/mp3", "wav": "audio/wav"}
                                    mime_type = mime_map.get(ext, "audio/mp4")
                                    
                                    with open(ruta_guardada, "rb") as f_aud:
                                        audio_bytes = f_aud.read()
                                        
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=[
                                            types.Part.from_bytes(data=audio_bytes, mime_type=mime_type),
                                            "Escucha con atención este archivo de audio. Realiza una transcripción sintetizada de los puntos clave tratados, detallando acuerdos, situaciones descritas y un diagnóstico para el informe de intervención."
                                        ]
                                    )
                                    resumen_txt = resp.text
                                
                                # 2. Imágenes (JPG, PNG)
                                elif ext in ["jpg", "jpeg", "png"]:
                                    img = Image.open(ruta_guardada)
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=["Describe y resume detalladamente los elementos clave de esta imagen para un informe de intervención:", img]
                                    )
                                    resumen_txt = resp.text
                                
                                # 3. Documentos (PDF, DOCX, PPTX, XLSX)
                                elif ext in ["pdf", "docx", "pptx", "xlsx", "xls"]:
                                    t_doc = extraer_texto_archivo(ruta_guardada, ext)
                                    resp = client.models.generate_content(
                                        model=MODELO_GEMINI,
                                        contents=f"Elabora un resumen y diagnóstico clave de este documento ({a.name}):\n\n{t_doc}"
                                    )
                                    resumen_txt = resp.text
                                    
                                # 4. Video (MP4)
                                elif ext == "mp4":
                                    resumen_txt = "Archivo de video registrado (reproducción disponible en visor)."
                                    
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
                st.success("Materiales integrados y analizados correctamente.")
                st.rerun()

    st.markdown("---")
    st.subheader("📚 Archivos y Diagnósticos Guardados")
    
    if not db["intervencion"]:
        st.info("No hay materiales en el módulo de intervención.")
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
                    if usr.get("permiso_eliminar") or usr_key == "admin1":
                        if st.button("🗑️ Borrar", key=f"del_int_{idx}"):
                            if ruta_arc and os.path.exists(ruta_arc):
                                os.remove(ruta_arc)
                            db["intervencion"].pop(idx)
                            guardar_estado(db)
                            st.success("Archivo eliminado.")
                            st.rerun()
                
                t_vis, t_res = st.tabs(["👁️ Visualizador Multimedia / Descarga", "📝 Diagnóstico y Resumen Visual"])
                
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
                        st.error("Archivo físico no encontrado en el servidor.")
                        
                with t_res:
                    st.info(resumen_arc)
                st.markdown("---")

# ==========================================
# 3. INFORMES COMPARTIDOS
# ==========================================
elif opcion == "📄 Informes Compartidos":
    st.markdown("""
    <div class='header-card'>
        <img class='header-img' src='https://cdn-icons-png.flaticon.com/512/2991/2991108.png'>
        <div>
            <h2 style='margin:0;'>Repositorio de Informes Compartidos</h2>
            <p style='margin:0; color:#94a3b8;'>Generación analítica integral y exportación en formato Carta.</p>
        </div>
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
                with st.spinner("Redactando informe consolidado con Gemini..."):
                    resumenes_int = "\n".join([f"- {x.get('nombre_original', 'Archivo')}: {x.get('resumen', '')}" for x in db["intervencion"]])
                    
                    prompt = f"""
                    Actúa como un especialista y analista de datos senior.
                    Genera un informe estructurado con enfoque '{enf_i}'.
                    Título: {nom_i}
                    Autor solicitante: {usr['nombre']}
                    
                    Instrucciones adicionales: {ins_i}
                    
                    Materiales de Intervención analizados (audios, documentos e imágenes):
                    {resumenes_int if resumenes_int else 'Sin materiales adjuntos.'}
                    
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
                        st.success("Informe publicado exitosamente.")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Error al generar informe: {err}")

    st.markdown("---")
    st.subheader("📚 Informes Publicados")
    
    if not db["informes"]:
        st.info("No hay informes registrados.")
    else:
        for idx, inf in enumerate(db["informes"]):
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
                    if usr.get("permiso_eliminar") or usr_key == "admin1":
                        if st.button("🗑️ Borrar", key=f"del_inf_{idx}"):
                            db["informes"].pop(idx)
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
                        key=f"xls_dl_{idx}"
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
                            body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #0f172a; padding: 20px; }}
                            .header {{ border-bottom: 2px solid #0284c7; padding-bottom: 8px; margin-bottom: 20px; }}
                            .title {{ font-size: 20pt; font-weight: bold; color: #0369a1; }}
                            .meta {{ font-size: 10pt; color: #64748b; margin-top: 4px; }}
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
                        label="🖼️ Descargar Formato Carta (Imprimible / PDF)",
                        data=html_carta,
                        file_name=f"[{autor_inf}] - {titulo_inf}_Carta.html",
                        mime="text/html",
                        key=f"doc_dl_{idx}"
                    )
                st.markdown("---")

# ==========================================
# 4. GESTIÓN DE USUARIOS
# ==========================================
elif opcion == "👥 Gestión de Usuarios":
    st.markdown("""
    <div class='header-card'>
        <img class='header-img' src='https://cdn-icons-png.flaticon.com/512/1256/1256650.png'>
        <div>
            <h2 style='margin:0;'>Administración y Control de Usuarios</h2>
            <p style='margin:0; color:#94a3b8;'>Creación, asignación de permisos y eliminación de perfiles.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("➕ Crear Nuevo Usuario", expanded=True):
        cu1, cu2, cu3 = st.columns([1, 1, 1])
        with cu1:
            n_u = st.text_input("Usuario (Login):")
            n_nom = st.text_input("Nombre Completo:")
        with cu2:
            n_pin = st.text_input("PIN / Clave:", type="password")
            n_rol = st.selectbox("Rol:", ["Analista", "Especialista", "Gerencia", "Admin"])
        with cu3:
            st.markdown("**Permisos Asignados:**")
            p_e = st.checkbox("Editar Datasets")
            p_d = st.checkbox("Eliminar Datasets / Archivos")
            
        if st.button("Guardar Perfil"):
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
                    "permiso_eliminar": p_d
                }
                guardar_estado(db)
                st.success(f"Usuario '{n_u}' registrado correctamente.")
                st.rerun()

    st.markdown("---")
    st.subheader("📜 Usuarios Registrados")
    
    for u_k in list(db["usuarios"].keys()):
        u_d = db["usuarios"][u_k]
        with st.container():
            col_u_inf, col_u_e, col_u_d, col_u_del = st.columns([2, 1, 1, 1])
            with col_u_inf:
                st.markdown(f"**{u_d.get('nombre', '')}** (`{u_k}`) — Rol: *{u_d.get('rol', '')}*")
            
            if u_k == "admin1":
                st.caption("Administrador Principal (Protegido)")
            else:
                with col_u_e:
                    val_e = st.checkbox("Editar", value=u_d.get("permiso_editar", False), key=f"pe_{u_k}")
                with col_u_d:
                    val_d = st.checkbox("Eliminar", value=u_d.get("permiso_eliminar", False), key=f"pd_{u_k}")
                with col_u_del:
                    if st.button("🗑️ Eliminar", key=f"del_user_{u_k}"):
                        del db["usuarios"][u_k]
                        guardar_estado(db)
                        st.success(f"Usuario {u_k} eliminado.")
                        st.rerun()
                        
                if val_e != u_d.get("permiso_editar") or val_d != u_d.get("permiso_eliminar"):
                    db["usuarios"][u_k]["permiso_editar"] = val_e
                    db["usuarios"][u_k]["permiso_eliminar"] = val_d
                    guardar_estado(db)
        st.markdown("---")
