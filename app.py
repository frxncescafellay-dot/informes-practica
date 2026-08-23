import os
import io
import json
import pandas as pd
import streamlit as st
from PIL import Image
from google import genai

# --- CONFIGURACIÓN PRINCIPAL DE LA PÁGINA ---
st.set_page_config(
    page_title="Plataforma de Gestión & Intervención IA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS PERSONALIZADOS (MEJORA VISUAL) ---
st.markdown("""
<style>
    /* Fondo principal y tipografía */
    .stApp {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #f8fafc;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Encabezados y títulos */
    h1, h2, h3, h4 {
        color: #38bdf8 !important;
        font-weight: 700;
    }
    
    /* Estilizado de Tarjetas y Contenedores */
    .custom-card {
        background-color: rgba(30, 41, 59, 0.7);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
    }
    
    /* Botones primarios */
    .stButton>button {
        background: linear-gradient(90deg, #0284c7 0%, #0369a1 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #0369a1 0%, #075985 100%);
        box-shadow: 0 0 12px rgba(56, 189, 248, 0.4);
    }
    
    /* Menú Lateral */
    section[data-testid="stSidebar"] {
        background-color: #0b0f19 !important;
        border-right: 1px solid #1e293b;
    }
    
    /* Badges / Etiquetas */
    .badge-user {
        background-color: #0369a1;
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    .badge-perm {
        background-color: #059669;
        color: white;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
    }
    
    /* Vista previa Formato Carta */
    .letter-preview {
        background-color: #ffffff;
        color: #1e293b;
        padding: 40px;
        border-radius: 8px;
        width: 100%;
        max-width: 800px;
        min-height: 1000px;
        margin: 20px auto;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
        font-family: 'Georgia', serif;
    }
</style>
""", unsafe_allow_html=True)

# --- CLIENTE GEMINI IA ---
API_KEY_GEMINI = os.environ.get("GEMINI_API_KEY", "")

def obtener_cliente_gemini():
    if not API_KEY_GEMINI:
        st.warning("⚠️ No se ha detectado GEMINI_API_KEY en las variables de entorno.")
        return None
    return genai.Client(api_key=API_KEY_GEMINI)

# --- INICIALIZACIÓN DE ESTADO (BASE DE DATOS EN MEMORIA) ---
if "usuarios_db" not in st.session_state:
    st.session_state.usuarios_db = {
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
    }

if "datasets_db" not in st.session_state:
    st.session_state.datasets_db = {}  # {id_dataset: {titulo, autor, df, fecha}}

if "intervencion_db" not in st.session_state:
    st.session_state.intervencion_db = []  # [{titulo, tipo, autor, resumen, bytes, filename}]

if "informes_db" not in st.session_state:
    st.session_state.informes_db = []  # [{titulo_informe, autor, fecha, contenido}]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None

# --- PANTALLA DE AUTENTICACIÓN (LOGIN) ---
if not st.session_state.autenticado:
    st.markdown("<h1 style='text-align: center;'>🔐 Acceso a la Plataforma</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Ingrese sus credenciales para continuar</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("form_login"):
            usr = st.text_input("Usuario", placeholder="Ej: admin1")
            pin = st.text_input("PIN / Contraseña", type="password")
            btn_login = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            
            if btn_login:
                db_usr = st.session_state.usuarios_db.get(usr)
                if db_usr and db_usr["pin"] == pin:
                    st.session_state.autenticado = True
                    st.session_state.usuario_clave = usr
                    st.success("¡Bienvenido/a!")
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
    st.stop()

# --- DATOS DEL USUARIO ACTUAL ---
usuario_actual_key = st.session_state.usuario_clave
usuario_actual = st.session_state.usuarios_db[usuario_actual_key]

# --- MENÚ LATERAL DE NAVEGACIÓN ---
st.sidebar.title("📌 Menú Principal")
st.sidebar.markdown(f"**Usuario:** {usuario_actual['nombre']}")
st.sidebar.markdown(f"<span class='badge-user'>{usuario_actual['rol']}</span>", unsafe_allow_html=True)

# Mostrar permisos activos en sidebar
permisos_str = []
if usuario_actual.get("permiso_editar"): permisos_str.append("Editar")
if usuario_actual.get("permiso_eliminar"): permisos_str.append("Eliminar")
if permisos_str:
    st.sidebar.caption("Permisos: " + ", ".join(permisos_str))

st.sidebar.markdown("---")

opciones_menu = ["📊 Datasets & Archivos", "📂 Intervención", "📄 Informes Compartidos"]
if usuario_actual_key == "admin1" or usuario_actual["rol"] == "Admin":
    opciones_menu.append("👥 Gestión de Usuarios")

opcion_seleccionada = st.sidebar.radio("Ir a:", opciones_menu)

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario_clave = None
    st.rerun()

# ==========================================
# SECCIÓN 1: DATASETS & ARCHIVOS (EXCEL)
# ==========================================
if opcion_seleccionada == "📊 Datasets & Archivos":
    st.title("📊 Gestión de Datasets Excel")
    st.markdown("Suba archivos Excel asignándoles un título. Cada archivo se desplegará en una ventana/pestana independiente.")
    
    # Subida de archivo con Título obligatorio
    with st.expander("➕ Cargar Nuevo Dataset Excel", expanded=True):
        col_t, col_f = st.columns([1, 1])
        with col_t:
            titulo_dataset = st.text_input("Título del Archivo / Dataset:", placeholder="Ej: Registro Ventas Q3 2026")
        with col_f:
            archivo_excel = st.file_uploader("Seleccione un archivo Excel:", type=["xlsx", "xls"])
        
        if st.button("Guardar Dataset", type="primary"):
            if not titulo_dataset.strip():
                st.error("Por favor ingrese un título para el dataset.")
            elif archivo_excel is None:
                st.error("Por favor seleccione un archivo Excel.")
            else:
                try:
                    df = pd.read_excel(archivo_excel)
                    id_dataset = f"{titulo_dataset}_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}"
                    st.session_state.datasets_db[id_dataset] = {
                        "titulo": titulo_dataset,
                        "autor": usuario_actual["nombre"],
                        "fecha": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                        "df": df
                    }
                    st.success(f"Dataset '{titulo_dataset}' guardado con éxito.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al leer el archivo Excel: {e}")

    st.markdown("---")
    
    # Visualización por Ventanas / Pestañas
    if not st.session_state.datasets_db:
        st.info("No hay datasets cargados actualmente.")
    else:
        titulos_datasets = [data["titulo"] for data in st.session_state.datasets_db.values()]
        tabs = st.tabs(titulos_datasets)
        
        keys_list = list(st.session_state.datasets_db.keys())
        for idx, tab in enumerate(tabs):
            dataset_key = keys_list[idx]
            dataset_info = st.session_state.datasets_db[dataset_key]
            
            with tab:
                st.subheader(f"📌 {dataset_info['titulo']}")
                st.caption(f"Subido por: **{dataset_info['autor']}** el {dataset_info['fecha']}")
                
                df_actual = dataset_info["df"]
                
                # Control de permisos de Edición
                if usuario_actual.get("permiso_editar") or usuario_actual_key == "admin1":
                    st.markdown("✏️ **Modo Edición Activado** (Puede modificar directamente las celdas):")
                    df_editado = st.data_editor(df_actual, key=f"editor_{dataset_key}", use_container_width=True)
                    
                    col_sav, col_del = st.columns([1, 1])
                    with col_sav:
                        if st.button("💾 Guardar Cambios en Datos", key=f"save_{dataset_key}"):
                            st.session_state.datasets_db[dataset_key]["df"] = df_editado
                            st.success("Cambios guardados correctamente.")
                            st.rerun()
                else:
                    st.dataframe(df_actual, use_container_width=True)
                
                # Control de permisos de Eliminación
                if usuario_actual.get("permiso_eliminar") or usuario_actual_key == "admin1":
                    st.markdown("---")
                    if st.button("🗑️ Eliminar este Dataset", key=f"del_{dataset_key}", type="secondary"):
                        del st.session_state.datasets_db[dataset_key]
                        st.success("Dataset eliminado.")
                        st.rerun()

# ==========================================
# SECCIÓN 2: MÓDULO DE INTERVENCIÓN
# ==========================================
elif opcion_seleccionada == "📂 Intervención":
    st.title("📂 Módulo de Intervención Multi-Formato")
    st.markdown("Suba archivos de soporte (Excel, Word, PDF, PPT, Imágenes, Videos). La IA generará automáticamente un resumen de los documentos e imágenes.")
    
    with st.form("form_intervencion", clear_on_submit=True):
        titulo_intervencion = st.text_input("Título del Material / Intervención:", placeholder="Ej: Registro Fotográfico y Reporte de Campo")
        archivos_subidos = st.file_uploader(
            "Cargar archivos (Excel, Word, PDF, PPT, JPG, PNG, MP4):",
            type=["xlsx", "xls", "docx", "pdf", "pptx", "jpg", "jpeg", "png", "mp4"],
            accept_multiple_files=True
        )
        btn_subir_intervencion = st.form_submit_button("Procesar y Guardar en Intervención")
        
        if btn_subir_intervencion:
            if not titulo_intervencion.strip():
                st.error("Por favor ingrese un título.")
            elif not archivos_subidos:
                st.error("Seleccione al menos un archivo.")
            else:
                client = obtener_cliente_gemini()
                for arc in archivos_subidos:
                    ext = arc.name.split(".")[-1].lower()
                    bytes_data = arc.read()
                    
                    resumen_generado = "Video cargado (sin resumen de texto automático)."
                    
                    # Generar resumen con IA si NO es video
                    if ext != "mp4" and client:
                        with st.spinner(f"Analizando y resumiendo {arc.name} con IA..."):
                            try:
                                if ext in ["jpg", "jpeg", "png"]:
                                    img = Image.open(io.BytesIO(bytes_data))
                                    response = client.models.generate_content(
                                        model="gemini-2.5-flash",
                                        contents=["Describe y resume detalladamente el contenido de esta imagen para un informe de intervención:", img]
                                    )
                                    resumen_generado = response.text
                                else:
                                    # Documentos de texto o tablas
                                    prompt_doc = f"Sintetiza y resume los aspectos clave del archivo '{arc.name}' para un informe técnico de intervención."
                                    response = client.models.generate_content(
                                        model="gemini-2.5-flash",
                                        contents=prompt_doc
                                    )
                                    resumen_generado = response.text
                            except Exception as err:
                                resumen_generado = f"Archivo guardado. No se pudo auto-resumir: {err}"
                    
                    st.session_state.intervencion_db.append({
                        "titulo": titulo_intervencion,
                        "filename": arc.name,
                        "tipo": ext,
                        "autor": usuario_actual["nombre"],
                        "resumen": resumen_generado,
                        "bytes": bytes_data
                    })
                st.success("Archivos procesados y añadidos al módulo de Intervención.")
                st.rerun()

    st.markdown("---")
    st.subheader("📋 Materiales de Intervención Registrados")
    
    if not st.session_state.intervencion_db:
        st.info("Aún no hay archivos de intervención registrados.")
    else:
        for idx, item in enumerate(st.session_state.intervencion_db):
            with st.container():
                st.markdown(f"### {item['titulo']} — `{item['filename']}`")
                st.caption(f"Cargado por: **{item['autor']}** | Tipo: **{item['tipo'].upper()}**")
                
                col_m, col_r = st.columns([1, 1])
                
                with col_m:
                    if item["tipo"] == "mp4":
                        st.video(item["bytes"])
                    elif item["tipo"] in ["jpg", "jpeg", "png"]:
                        st.image(item["bytes"], use_container_width=True)
                    else:
                        st.info(f"📄 Archivo de documento ({item['tipo'].upper()}) disponible para descarga/revisión.")
                        st.download_button(
                            label=f"📥 Descargar {item['filename']}",
                            data=item["bytes"],
                            file_name=item["filename"],
                            key=f"dl_int_{idx}"
                        )
                
                with col_r:
                    st.markdown("**Resumen / Diagnóstico:**")
                    st.write(item["resumen"])
                
                st.markdown("---")

# ==========================================
# SECCIÓN 3: INFORMES COMPARTIDOS & GENERADOR
# ==========================================
elif opcion_seleccionada == "📄 Informes Compartidos":
    st.title("📄 Informes Colaborativos e Inteligentes")
    st.markdown("Genera informes consolidados con IA utilizando los datasets y materiales disponibles. **Todos los usuarios pueden visualizar los informes generados.**")
    
    # Generador de Informes
    with st.expander("🤖 Generar Nuevo Informe con IA", expanded=False):
        nombre_informe = st.text_input("Título del Informe:", placeholder="Ej: Informe Consolidado de Intervención Q3")
        enfoque = st.selectbox("Enfoque del Informe:", ["Resumen Ejecutivo", "Diagnóstico Técnico", "Evaluación de Riesgos y Recomendaciones"])
        instrucciones = st.text_area("Instrucciones o contexto adicional:")
        
        if st.button("🚀 Generar y Compartir Informe", type="primary"):
            if not nombre_informe.strip():
                st.error("Ingrese un nombre para el informe.")
            else:
                client = obtener_cliente_gemini()
                if not client:
                    st.error("API Key de Gemini no configurada.")
                else:
                    with st.spinner("Compilando datos y redactando informe..."):
                        # Contexto consolidado
                        resumenes_int = "\n".join([f"- {i['filename']}: {i['resumen']}" for i in st.session_state.intervencion_db])
                        data_keys = list(st.session_state.datasets_db.keys())
                        
                        prompt = f"""
                        Actúa como un consultor experto.
                        Genera un informe con enfoque '{enfoque}'.
                        Título del Informe: {nombre_informe}
                        Autor que lo genera: {usuario_actual['nombre']}
                        
                        Instrucciones del usuario: {instrucciones}
                        
                        Resumen de materiales de Intervención disponibles:
                        {resumenes_int if resumenes_int else 'Sin materiales adicionales.'}
                        
                        Estructura requerida:
                        1. Resumen Ejecutivo
                        2. Hallazgos y Análisis General
                        3. Conclusiones y Plan de Acción
                        """
                        try:
                            resp = client.models.generate_content(
                                model="gemini-2.5-flash",
                                contents=prompt
                            )
                            
                            st.session_state.informes_db.append({
                                "titulo_informe": nombre_informe,
                                "autor": usuario_actual["nombre"],
                                "fecha": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                                "contenido": resp.text
                            })
                            st.success("Informe generado y publicado en el repositorio compartido.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error generando informe: {e}")

    st.markdown("---")
    st.subheader("📚 Repositorio de Informes Compartidos")
    
    if not st.session_state.informes_db:
        st.info("No hay informes guardados aún.")
    else:
        for idx_inf, inf in enumerate(st.session_state.informes_db):
            with st.container():
                st.markdown(f"### 📄 [{inf['autor']}] - {inf['titulo_informe']}")
                st.caption(f"Fecha de creación: {inf['fecha']}")
                
                st.markdown(inf["contenido"])
                
                # OPCIONES DE DESCARGA
                col_d1, col_d2 = st.columns([1, 1])
                
                with col_d1:
                    # Exportar a Excel
                    df_exp = pd.DataFrame([{"Título": inf["titulo_informe"], "Autor": inf["autor"], "Fecha": inf["fecha"], "Informe": inf["contenido"]}])
                    buffer_excel = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
                        df_exp.to_excel(writer, index=False, sheet_name="Informe")
                    
                    st.download_button(
                        label="📥 Descargar en Excel (.xlsx)",
                        data=buffer_excel.getvalue(),
                        file_name=f"[{inf['autor']}] - {inf['titulo_informe']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"excel_{idx_inf}"
                    )
                
                with col_d2:
                    # Exportar a Formato Especial Tamaño Carta (HTML Imprimible / Guardar como PDF)
                    html_carta = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>{inf['titulo_informe']}</title>
                        <style>
                            @page {{
                                size: letter portrait;
                                margin: 20mm;
                            }}
                            body {{
                                font-family: 'Arial', sans-serif;
                                color: #111827;
                                line-height: 1.6;
                                padding: 20px;
                            }}
                            .header {{
                                border-bottom: 3px solid #0284c7;
                                padding-bottom: 10px;
                                margin-bottom: 20px;
                            }}
                            .title {{ font-size: 24px; color: #0369a1; font-weight: bold; margin: 0; }}
                            .meta {{ font-size: 12px; color: #4b5563; margin-top: 5px; }}
                            .content {{ font-size: 14px; white-space: pre-wrap; }}
                            .footer {{ margin-top: 40px; border-top: 1px solid #e5e7eb; padding-top: 10px; font-size: 10px; text-align: center; color: #9ca3af; }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <div class="title">{inf['titulo_informe']}</div>
                            <div class="meta">Autor: {inf['autor']} | Fecha: {inf['fecha']}</div>
                        </div>
                        <div class="content">{inf['contenido']}</div>
                        <div class="footer">Documento Oficial generado por Plataforma IA — Formato Tamaño Carta</div>
                    </body>
                    </html>
                    """
                    
                    st.download_button(
                        label="🖼️ Descargar Formato Especial (Carta / PDF)",
                        data=html_carta,
                        file_name=f"[{inf['autor']}] - {inf['titulo_informe']}_Carta.html",
                        mime="text/html",
                        key=f"carta_{idx_inf}"
                    )
                st.markdown("---")

# ==========================================
# SECCIÓN 4: GESTIÓN DE USUARIOS Y PERMISOS
# ==========================================
elif opcion_seleccionada == "👥 Gestión de Usuarios":
    st.title("👥 Panel de Gestión de Usuarios y Permisos")
    st.markdown("Administrado exclusivamente por el usuario **admin1**.")
    
    # Crear Nuevo Usuario
    with st.expander("➕ Registrar Nuevo Usuario", expanded=True):
        col_u1, col_u2, col_u3 = st.columns([1, 1, 1])
        with col_u1:
            nuevo_usr = st.text_input("Nombre de Usuario (Login):")
            nuevo_nombre = st.text_input("Nombre Completo:")
        with col_u2:
            nuevo_pin = st.text_input("PIN / Contraseña:", type="password")
            nuevo_rol = st.selectbox("Rol:", ["Analista", "Especialista", "Gerencia", "Admin"])
        with col_u3:
            st.markdown("**Otorgar Permisos:**")
            p_edit = st.checkbox("Permiso para Modificar / Editar Datasets")
            p_del = st.checkbox("Permiso para Eliminar Datasets")
            
        if st.button("Guardar Usuario", type="primary"):
            if not nuevo_usr or not nuevo_pin or not nuevo_nombre:
                st.error("Complete todos los campos requeridos.")
            elif nuevo_usr in st.session_state.usuarios_db:
                st.error("El usuario ya existe.")
            else:
                st.session_state.usuarios_db[nuevo_usr] = {
                    "nombre": nuevo_nombre,
                    "rol": nuevo_rol,
                    "pin": nuevo_pin,
                    "permiso_editar": p_edit,
                    "permiso_eliminar": p_del
                }
                st.success(f"Usuario '{nuevo_usr}' registrado con éxito.")
                st.rerun()

    st.markdown("---")
    st.subheader("📜 Usuarios Registrados y Modificación de Permisos")
    
    for u_key, u_data in st.session_state.usuarios_db.items():
        with st.container():
            col_info, col_p1, col_p2 = st.columns([2, 1, 1])
            with col_info:
                st.markdown(f"**{u_data['nombre']}** (`{u_key}`) — Rol: *{u_data['rol']}*")
            
            # No modificar al admin principal
            if u_key == "admin1":
                st.caption("Administrador Principal (Permisos Totales)")
            else:
                with col_p1:
                    edit_val = st.checkbox("Editar Datasets", value=u_data["permiso_editar"], key=f"p_e_{u_key}")
                with col_p2:
                    del_val = st.checkbox("Eliminar Datasets", value=u_data["permiso_eliminar"], key=f"p_d_{u_key}")
                
                # Actualizar permisos inmediatamente al cambiar
                st.session_state.usuarios_db[u_key]["permiso_editar"] = edit_val
                st.session_state.usuarios_db[u_key]["permiso_eliminar"] = del_val
        st.markdown("---")
