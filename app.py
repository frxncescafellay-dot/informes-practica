import os
import io
import json
import shutil
import pandas as pd
import streamlit as st
from PIL import Image
from google import genai
import pypdf
from docx import Document
from pptx import Presentation

# --- CONFIGURACIÓN DE PÁGINA Y LÍMITES ---
st.set_page_config(
    page_title="Plataforma Integral de Intervención IA",
    page_icon="📁",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CREACIÓN DE DIRECTORIOS PARA PERSISTENCIA TOTAL ---
DIR_BASE = "almacen_datos"
DIR_DATASETS = os.path.join(DIR_BASE, "datasets")
DIR_INTERVENCION = os.path.join(DIR_BASE, "intervencion")
FILE_DB = os.path.join(DIR_BASE, "base_datos.json")

for d in [DIR_BASE, DIR_DATASETS, DIR_INTERVENCION]:
    os.makedirs(d, exist_ok=True)

# --- ESTILOS VISUALES ---
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #0b132b 0%, #1c2541 100%);
        color: #f1f5f9;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    h1, h2, h3, h4 {
        color: #48cae4 !important;
        font-weight: 700;
    }
    section[data-testid="stSidebar"] {
        background-color: #080d1a !important;
        border-right: 1px solid #1e293b;
    }
    .stButton>button {
        background: linear-gradient(90deg, #0077b6 0%, #023e8a 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 8px 20px;
        font-weight: 600;
    }
    .badge-user {
        background-color: #0096c7;
        color: white;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# --- GESTOR DE PERSISTENCIA EN DISCO (JSON) ---
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
                },
                "analista1": {
                    "nombre": "Juan Pérez",
                    "rol": "Analista",
                    "pin": "5678",
                    "permiso_editar": False,
                    "permiso_eliminar": False
                }
            },
            "datasets": {},
            "intervencion": [],
            "informes": []
        }
        guardar_estado(data_inicial)
        return data_inicial
    with open(FILE_DB, "r", encoding="utf-8") as f:
        return json.load(f)

def guardar_estado(data):
    with open(FILE_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

db = cargar_estado()

# --- CLIENTE DE INTELIGENCIA ARTIFICIAL ---
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
            texto = f"Resumen estadístico del Excel:\n{df_tmp.describe(include='all').to_string()}\nPrimeras filas:\n{df_tmp.head(10).to_string()}"
    except Exception as e:
        texto = f"Error extrayendo texto: {e}"
    return texto[:8000]

# --- MANEJO DE SESIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None

if not st.session_state.autenticado:
    st.markdown("<h1 style='text-align:center;'>🔐 Acceso al Sistema</h1>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login"):
            usr = st.text_input("Usuario")
            pin = st.text_input("PIN / Contraseña", type="password")
            if st.form_submit_button("Entrar", use_container_width=True):
                if usr in db["usuarios"] and db["usuarios"][usr]["pin"] == pin:
                    st.session_state.autenticado = True
                    st.session_state.usuario_clave = usr
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
    st.stop()

usr_actual_key = st.session_state.usuario_clave
usr_actual = db["usuarios"].get(usr_actual_key, {"nombre": "Desconocido", "rol": "Invitado", "permiso_editar": False, "permiso_eliminar": False})

# --- SIDEBAR DE NAVEGACIÓN ---
st.sidebar.markdown(f"**Usuario:** {usr_actual['nombre']}")
st.sidebar.markdown(f"<span class='badge-user'>{usr_actual['rol']}</span>", unsafe_allow_html=True)
st.sidebar.markdown("---")

opciones = ["📊 Datasets & Archivos", "📂 Intervención", "📄 Informes Compartidos"]
if usr_actual_key == "admin1" or usr_actual.get("rol") == "Admin":
    opciones.append("👥 Gestión de Usuarios")

seleccion = st.sidebar.radio("Navegación:", opciones)

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None
    st.rerun()

# ==========================================
# SECCIÓN 1: DATASETS EXCEL
# ==========================================
if seleccion == "📊 Datasets & Archivos":
    st.title("📊 Gestión Persistente de Datasets")
    
    with st.expander("➕ Subir Nuevo Dataset Excel con Ventana Propia", expanded=True):
        col_t, col_f = st.columns([1, 1])
        with col_t:
            titulo_dataset = st.text_input("Título descriptivo del dataset:")
        with col_f:
            archivo_excel = st.file_uploader("Archivo Excel (.xlsx, .xls)", type=["xlsx", "xls"])
        
        if st.button("Guardar Dataset"):
            if not titulo_dataset.strip() or archivo_excel is None:
                st.error("Debe ingresar un título y seleccionar un archivo.")
            else:
                nombre_archivo_disco = f"{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}_{archivo_excel.name}"
                ruta_disco = os.path.join(DIR_DATASETS, nombre_archivo_disco)
                
                with open(ruta_disco, "wb") as f:
                    f.write(archivo_excel.getbuffer())
                
                db["datasets"][titulo_dataset] = {
                    "titulo": titulo_dataset,
                    "autor": usr_actual["nombre"],
                    "fecha": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                    "ruta": ruta_disco
                }
                guardar_estado(db)
                st.success("Dataset guardado permanentemente.")
                st.rerun()

    st.markdown("---")
    
    if not db["datasets"]:
        st.info("No hay datasets subidos.")
    else:
        titulos = list(db["datasets"].keys())
        pestanas = st.tabs(titulos)
        
        for idx, tab in enumerate(pestanas):
            t_actual = titulos[idx]
            info = db["datasets"][t_actual]
            
            with tab:
                st.subheader(f"Dataset: {info['titulo']}")
                st.caption(f"Subido por: **{info['autor']}** el {info['fecha']}")
                
                if os.path.exists(info["ruta"]):
                    df_cargado = pd.read_excel(info["ruta"])
                    
                    if usr_actual.get("permiso_editar") or usr_actual_key == "admin1":
                        st.write("✏️ **Edición de datos habilitada:**")
                        df_modificado = st.data_editor(df_cargado, key=f"edit_{t_actual}", use_container_width=True)
                        
                        col_s, col_d = st.columns([1, 1])
                        with col_s:
                            if st.button("💾 Guardar Modificaciones", key=f"btn_save_{t_actual}"):
                                df_modificado.to_excel(info["ruta"], index=False)
                                st.success("Cambios persistidos en el archivo original.")
                                st.rerun()
                    else:
                        st.dataframe(df_cargado, use_container_width=True)
                    
                    if usr_actual.get("permiso_eliminar") or usr_actual_key == "admin1":
                        st.markdown("---")
                        if st.button("🗑️ Eliminar Dataset", key=f"btn_del_{t_actual}"):
                            if os.path.exists(info["ruta"]):
                                os.remove(info["ruta"])
                            del db["datasets"][t_actual]
                            guardar_estado(db)
                            st.success("Dataset eliminado.")
                            st.rerun()
                else:
                    st.error("El archivo físico no se encuentra en el servidor.")

# ==========================================
# SECCIÓN 2: MÓDULO INTERVENCIÓN
# ==========================================
elif seleccion == "📂 Intervención":
    st.title("📂 Módulo de Intervención Multi-Formato")
    st.markdown("Formatos soportados: **Excel, Word, PDF, PPT, JPG, PNG, MP4, M4A** (Límite 500 MB).")
    
    with st.form("form_intervencion"):
        tit_int = st.text_input("Título del registro de Intervención:")
        archivos = st.file_uploader(
            "Seleccionar archivos:",
            type=["xlsx", "xls", "docx", "pdf", "pptx", "jpg", "jpeg", "png", "mp4", "m4a"],
            accept_multiple_files=True
        )
        subir_int = st.form_submit_button("Subir y Procesar Archivos")
        
        if subir_int:
            if not tit_int.strip() or not archivos:
                st.error("Complete el título y cargue al menos un archivo.")
            else:
                client = obtener_cliente_ia()
                for a in archivos:
                    ext = a.name.split(".")[-1].lower()
                    nombre_guardado = f"{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}_{a.name}"
                    ruta_int = os.path.join(DIR_INTERVENCION, nombre_guardado)
                    
                    # Guardar archivo físico
                    with open(ruta_int, "wb") as f:
                        f.write(a.getbuffer())
                    
                    resumen_txt = "Archivo de audio/video almacenado (sin resumen de texto automático)."
                    
                    # Generar resumen con IA si NO es video ni audio
                    if ext not in ["mp4", "m4a"] and client:
                        with st.spinner(f"Analizando {a.name}..."):
                            try:
                                if ext in ["jpg", "jpeg", "png"]:
                                    img = Image.open(ruta_int)
                                    resp = client.models.generate_content(
                                        model="gemini-3.6-flash",
                                        contents=["Realiza un diagnóstico y resumen visual detallado de esta imagen para un informe de intervención:", img]
                                    )
                                    resumen_txt = resp.text
                                else:
                                    texto_doc = extraer_texto_archivo(ruta_int, ext)
                                    resp = client.models.generate_content(
                                        model="gemini-2.5-flash",
                                        contents=f"Genera un resumen analítico y diagnóstico del siguiente documento '{a.name}':\n\n{texto_doc}"
                                    )
                                    resumen_txt = resp.text
                            except Exception as e:
                                resumen_txt = f"Archivo guardado. No se pudo generar resumen automático: {e}"
                    
                    db["intervencion"].append({
                        "titulo_registro": tit_int,
                        "nombre_original": a.name,
                        "tipo": ext,
                        "autor": usr_actual["nombre"],
                        "fecha": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                        "ruta": ruta_int,
                        "resumen": resumen_txt
                    })
                guardar_estado(db)
                st.success("Archivos procesados y guardados permanentemente.")
                st.rerun()

    st.markdown("---")
    st.subheader("📋 Registros de Intervención Almacenados")
    
    if not db["intervencion"]:
        st.info("No hay archivos en el módulo de intervención.")
    else:
        for idx, item in enumerate(db["intervencion"]):
            with st.container():
                st.markdown(f"### 📂 {item['titulo_registro']} — `{item['nombre_original']}`")
                st.caption(f"Autor: **{item['autor']}** | Fecha: {item['fecha']} | Formato: **{item['tipo'].upper()}**")
                
                # Pestañas separadas para el archivo y para el resumen
                tab_visual, tab_resumen = st.tabs(["👁️ Archivo Multimedia / Descarga", "📝 Diagnóstico y Resumen Visual"])
                
                with tab_visual:
                    if os.path.exists(item["ruta"]):
                        if item["tipo"] == "mp4":
                            st.video(item["ruta"])
                        elif item["tipo"] == "m4a":
                            st.audio(item["ruta"])
                        elif item["tipo"] in ["jpg", "jpeg", "png"]:
                            st.image(item["ruta"], use_container_width=True)
                        else:
                            with open(item["ruta"], "rb") as f_descarga:
                                st.download_button(
                                    label=f"📥 Descargar {item['nombre_original']}",
                                    data=f_descarga.read(),
                                    file_name=item["nombre_original"],
                                    key=f"dl_file_{idx}"
                                )
                    else:
                        st.error("El archivo no se encuentra disponible en disco.")
                
                with tab_resumen:
                    st.markdown("#### Resumen Analítico Generado")
                    st.info(item["resumen"])
                
                st.markdown("---")

# ==========================================
# SECCIÓN 3: INFORMES COMPARTIDOS
# ==========================================
elif seleccion == "📄 Informes Compartidos":
    st.title("📄 Informes Compartidos del Equipo")
    
    with st.expander("🤖 Redactar Nuevo Informe con IA", expanded=False):
        nom_inf = st.text_input("Título del Informe:")
        enfoque_inf = st.selectbox("Enfoque:", ["Resumen Ejecutivo", "Diagnóstico Técnico", "Evaluación de Riesgos"])
        extra_inst = st.text_area("Instrucciones complementarias:")
        
        if st.button("🚀 Generar Informe"):
            if not nom_inf.strip():
                st.error("Ingrese un título para el informe.")
            else:
                client = obtener_cliente_ia()
                if not client:
                    st.error("API Key de Gemini no configurada.")
                else:
                    with st.spinner("Compilando materiales y generando informe..."):
                        resumenes = "\n".join([f"- Archivo {i['nombre_original']}: {i['resumen']}" for i in db["intervencion"]])
                        
                        prompt = f"""
                        Actúa como un consultor y analista senior.
                        Genera un informe analítico con enfoque '{enfoque_inf}'.
                        Título: {nom_inf}
                        Autor solicitante: {usr_actual['nombre']}
                        
                        Instrucciones:
                        {extra_inst}
                        
                        Materiales de Intervención disponibles:
                        {resumenes if resumenes else 'Sin materiales adjuntos.'}
                        
                        Estructura requerida:
                        1. Diagnóstico General
                        2. Hallazgos Clave
                        3. Plan de Acción Recomendado
                        """
                        try:
                            resp = client.models.generate_content(
                                model="gemini-2.5-flash",
                                contents=prompt
                            )
                            db["informes"].append({
                                "titulo": nom_inf,
                                "autor": usr_actual["nombre"],
                                "fecha": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                                "contenido": resp.text
                            })
                            guardar_estado(db)
                            st.success("Informe guardado y publicado para todos los usuarios.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error generando informe: {e}")

    st.markdown("---")
    st.subheader("📚 Repositorio General de Informes")
    
    if not db["informes"]:
        st.info("No hay informes generados.")
    else:
        for idx_inf, inf in enumerate(db["informes"]):
            with st.container():
                st.markdown(f"### 📄 [{inf['autor']}] - {inf['titulo']}")
                st.caption(f"Generado el: {inf['fecha']}")
                st.markdown(inf["contenido"])
                
                col_d1, col_d2 = st.columns([1, 1])
                with col_d1:
                    df_out = pd.DataFrame([{"Autor": inf["autor"], "Título": inf["titulo"], "Fecha": inf["fecha"], "Contenido": inf["contenido"]}])
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_out.to_excel(writer, index=False)
                    st.download_button(
                        label="📥 Descargar en Excel (.xlsx)",
                        data=buf.getvalue(),
                        file_name=f"[{inf['autor']}] - {inf['titulo']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"xls_{idx_inf}"
                    )
                with col_d2:
                    html_carta = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>{inf['titulo']}</title>
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
                            <div class="title">{inf['titulo']}</div>
                            <div class="meta">Autor: {inf['autor']} | Fecha: {inf['fecha']}</div>
                        </div>
                        <div class="body">{inf['contenido']}</div>
                    </body>
                    </html>
                    """
                    st.download_button(
                        label="🖼️ Descargar Formato Carta (Imprimible / PDF)",
                        data=html_carta,
                        file_name=f"[{inf['autor']}] - {inf['titulo']}_Carta.html",
                        mime="text/html",
                        key=f"doc_{idx_inf}"
                    )
                st.markdown("---")

# ==========================================
# SECCIÓN 4: GESTIÓN Y ELIMINACIÓN DE USUARIOS
# ==========================================
elif seleccion == "👥 Gestión de Usuarios":
    st.title("👥 Gestión y Control de Usuarios")
    
    with st.expander("➕ Registrar Nuevo Usuario", expanded=True):
        col_u1, col_u2, col_u3 = st.columns([1, 1, 1])
        with col_u1:
            n_usr = st.text_input("Usuario (Login):")
            n_nom = st.text_input("Nombre Completo:")
        with col_u2:
            n_pin = st.text_input("PIN:", type="password")
            n_rol = st.selectbox("Rol:", ["Analista", "Especialista", "Gerencia", "Admin"])
        with col_u3:
            st.markdown("**Permisos:**")
            p_e = st.checkbox("Editar Datasets")
            p_d = st.checkbox("Eliminar Datasets")
        
        if st.button("Guardar Usuario"):
            if not n_usr or not n_pin or not n_nom:
                st.error("Complete todos los campos.")
            elif n_usr in db["usuarios"]:
                st.error("El usuario ya existe.")
            else:
                db["usuarios"][n_usr] = {
                    "nombre": n_nom,
                    "rol": n_rol,
                    "pin": n_pin,
                    "permiso_editar": p_e,
                    "permiso_eliminar": p_d
                }
                guardar_estado(db)
                st.success(f"Usuario '{n_usr}' registrado.")
                st.rerun()

    st.markdown("---")
    st.subheader("📜 Usuarios Registrados")
    
    usuarios_claves = list(db["usuarios"].keys())
    for u_key in usuarios_claves:
        u_data = db["usuarios"][u_key]
        with st.container():
            col_inf, col_p1, col_p2, col_del = st.columns([2, 1, 1, 1])
            with col_inf:
                st.markdown(f"**{u_data['nombre']}** (`{u_key}`) — *{u_data['rol']}*")
            
            if u_key == "admin1":
                st.caption("Administrador Principal (Protegido)")
            else:
                with col_p1:
                    e_val = st.checkbox("Editar", value=u_data.get("permiso_editar", False), key=f"pe_{u_key}")
                with col_p2:
                    d_val = st.checkbox("Eliminar", value=u_data.get("permiso_eliminar", False), key=f"pd_{u_key}")
                with col_del:
                    if st.button("🗑️ Eliminar Usuario", key=f"del_u_{u_key}"):
                        del db["usuarios"][u_key]
                        guardar_estado(db)
                        st.success(f"Usuario {u_key} eliminado.")
                        st.rerun()
                
                # Actualizar permisos
                if e_val != u_data.get("permiso_editar") or d_val != u_data.get("permiso_eliminar"):
                    db["usuarios"][u_key]["permiso_editar"] = e_val
                    db["usuarios"][u_key]["permiso_eliminar"] = d_val
                    guardar_estado(db)
        st.markdown("---")
