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

# --- PALETA TEAL + NARANJA TERRACOTA (CERO ELEMENTOS NEGROS) ---
st.markdown("""
<style>
    /* Tipografía uniforme */
    html, body, [class*="css"], .stApp {
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
    }

    /* Fondo principal */
    .stApp {
        background: #73b5ae !important;
        color: #061e1b;
    }
    
    /* Encabezados y títulos */
    h1, h2, h3, h4, h5, h6 {
        color: #061e1b !important;
        font-weight: 800;
        letter-spacing: -0.5px;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
    }
    
    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #63a59e !important;
        border-right: 2px solid #52948d !important;
    }
    section[data-testid="stSidebar"] * {
        color: #061e1b !important;
        font-weight: 600;
    }

    /* Tarjetas y Contenedores */
    .modern-card {
        background: #89c7c0 !important;
        border: 2px solid #63a59e !important;
        border-radius: 14px;
        padding: 22px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(13, 44, 41, 0.12);
    }

    /* Desplegables (Expanders) */
    div[data-testid="stExpander"] {
        background-color: #89c7c0 !important;
        border: 2px solid #63a59e !important;
        border-radius: 10px !important;
    }
    div[data-testid="stExpander"] summary {
        background-color: #7ab8b1 !important;
        color: #061e1b !important;
        border-radius: 8px !important;
        font-weight: 750 !important;
    }
    div[data-testid="stExpander"] summary * {
        color: #061e1b !important;
        font-weight: 750 !important;
    }
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #89c7c0 !important;
        color: #061e1b !important;
    }

    /* Cuadros de texto */
    input[type="text"], 
    input[type="password"], 
    textarea {
        background-color: #a2d2cc !important;
        color: #061e1b !important;
        border: 2px solid #52948d !important;
        border-radius: 8px !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
    }
    input[type="text"]:focus, 
    input[type="password"]:focus, 
    textarea:focus {
        background-color: #bfe3de !important;
        color: #000000 !important;
        border-color: #c84b1e !important;
        box-shadow: 0 0 0 3px rgba(200, 75, 30, 0.25) !important;
    }

    /* Quitar fondo negro del botón de ver contraseña en input password */
    div[data-baseweb="input"] button,
    div[data-baseweb="input"] div {
        background-color: #a2d2cc !important;
        border: none !important;
        color: #061e1b !important;
    }
    div[data-baseweb="input"] svg {
        fill: #061e1b !important;
        color: #061e1b !important;
    }

    /* Selectores desplegables (Selectbox) */
    div[data-baseweb="select"] > div {
        background-color: #a2d2cc !important;
        color: #061e1b !important;
        border: 2px solid #52948d !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] * {
        background-color: #a2d2cc !important;
        color: #061e1b !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="select"] svg {
        fill: #061e1b !important;
        color: #061e1b !important;
    }

    /* CUADROS PARA MARCAR (Checkboxes) */
    div[data-testid="stCheckbox"] span[role="checkbox"] {
        background-color: #a2d2cc !important;
        border: 2px solid #52948d !important;
        border-radius: 5px !important;
    }
    div[data-testid="stCheckbox"] span[role="checkbox"][aria-checked="true"] {
        background-color: #c84b1e !important;
        border-color: #9e3610 !important;
    }
    div[data-testid="stCheckbox"] svg {
        color: #ffffff !important;
        fill: #ffffff !important;
    }
    div[data-testid="stCheckbox"] label span {
        color: #061e1b !important;
        font-weight: 600 !important;
    }

    /* Zona de carga de archivos (Uploader) */
    div[data-testid="stFileUploader"] {
        background-color: #89c7c0 !important;
        border: 2px dashed #c84b1e !important;
        border-radius: 12px !important;
        padding: 14px !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: #89c7c0 !important;
    }
    div[data-testid="stFileUploader"] section * {
        color: #061e1b !important;
    }
    div[data-testid="stFileUploader"] button {
        background-color: #c84b1e !important;
        border: 2px solid #9e3610 !important;
        border-radius: 8px !important;
        font-weight: 750 !important;
    }
    div[data-testid="stFileUploader"] button p {
        color: #ffffff !important;
    }

    /* BOTONES DE DESCARGA */
    div[data-testid="stDownloadButton"]>button {
        background: #a2d2cc !important;
        color: #061e1b !important;
        border: 2px solid #52948d !important;
        border-radius: 9px !important;
        padding: 9px 20px !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 3px 8px rgba(13, 44, 41, 0.15) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stDownloadButton"]>button:hover {
        background: #bfe3de !important;
        color: #000000 !important;
        border-color: #061e1b !important;
        transform: translateY(-1px);
    }
    div[data-testid="stDownloadButton"]>button p {
        color: #061e1b !important;
        font-weight: 800 !important;
    }

    /* BOTONES PRINCIPALES Y FORMULARIOS */
    .stButton>button, 
    div[data-testid="stFormSubmitButton"]>button {
        background: linear-gradient(135deg, #c84b1e 0%, #b23b14 100%) !important;
        color: #ffffff !important;
        border: 2px solid #9e3610 !important;
        border-radius: 9px !important;
        padding: 9px 22px !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 12px rgba(178, 59, 20, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button:hover, 
    div[data-testid="stFormSubmitButton"]>button:hover {
        background: linear-gradient(135deg, #db5928 0%, #c84b1e 100%) !important;
        border-color: #db5928 !important;
        box-shadow: 0 6px 16px rgba(219, 89, 40, 0.45) !important;
        transform: translateY(-1px);
    }
    .stButton>button p,
    div[data-testid="stFormSubmitButton"]>button p {
        color: #ffffff !important;
        font-weight: 800 !important;
    }

    /* Tablas de datos Excel */
    div[data-testid="stDataFrame"], 
    div[data-testid="stDataEditor"],
    div[data-testid="stDataFrame"] > div,
    div[data-testid="stDataEditor"] > div {
        background-color: #89c7c0 !important;
        border: 2px solid #52948d !important;
        border-radius: 10px !important;
    }
    
    .glide-data-grid,
    .gdg-container,
    div[data-testid="stDataFrame"] canvas,
    div[data-testid="stDataEditor"] canvas {
        filter: invert(0.85) sepia(0.3) saturate(2.5) hue-rotate(130deg) brightness(1.05) !important;
        border-radius: 8px !important;
    }

    /* ESTILO PROFESIONAL PARA TABLAS MARKDOWN DENTRO DE INFORMES */
    table {
        width: 100% !important;
        border-collapse: collapse !important;
        margin: 16px 0 !important;
        background-color: #89c7c0 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
        border: 2px solid #52948d !important;
    }
    th {
        background-color: #63a59e !important;
        color: #061e1b !important;
        font-weight: 800 !important;
        padding: 12px 14px !important;
        border: 1px solid #52948d !important;
        text-align: left !important;
    }
    td {
        padding: 10px 14px !important;
        border: 1px solid #52948d !important;
        color: #061e1b !important;
        font-size: 0.95rem !important;
    }
    tr:nth-child(even) {
        background-color: #7ab8b1 !important;
    }

    /* Textos con fondo destacado para nombres de usuarios y archivos */
    .highlight-tag {
        background: #a2d2cc;
        border: 1.5px solid #52948d;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 700;
        color: #061e1b;
        font-size: 0.92rem;
        display: inline-block;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
    }

    code, pre {
        background-color: #a2d2cc !important;
        color: #061e1b !important;
        border: 1.5px solid #52948d !important;
        border-radius: 6px !important;
        padding: 3px 8px !important;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
        font-size: 0.92rem !important;
        font-weight: 700 !important;
    }
    
    /* Botones de eliminación / secundarios */
    div[data-testid="stBaseButton-secondary"] button {
        background: #e29b9b !important;
        border: 2px solid #c96b6b !important;
        box-shadow: none !important;
    }
    div[data-testid="stBaseButton-secondary"] button:hover {
        background: #d67a7a !important;
    }
    div[data-testid="stBaseButton-secondary"] button p {
        color: #5c0f0f !important;
    }

    /* Pestañas de navegación */
    button[data-baseweb="tab"] {
        background-color: #63a59e !important;
        color: #061e1b !important;
        border-radius: 8px 8px 0 0 !important;
        font-weight: 750 !important;
        margin-right: 4px !important;
        padding: 9px 18px !important;
        border: 1px solid #52948d !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #89c7c0 !important;
        color: #061e1b !important;
        border-bottom: 3.5px solid #c84b1e !important;
    }

    label, p, span {
        color: #061e1b;
        font-weight: 600;
    }

    .badge-role {
        display: inline-block;
        background: #c84b1e;
        color: #ffffff !important;
        padding: 4px 12px;
        border-radius: 14px;
        font-size: 0.82rem;
        font-weight: 800;
        border: 1.5px solid #9e3610;
    }
</style>
""", unsafe_allow_html=True)

# --- RELOJ DUAL NARANJA OSCURO / TERRACOTA ---
def renderizar_reloj_chile():
    html_reloj = """
    <div style="
        background: linear-gradient(135deg, #c84b1e 0%, #b23b14 100%);
        border: 2px solid #9e3610;
        border-radius: 16px;
        padding: 12px 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        box-shadow: 0 6px 18px rgba(178, 59, 20, 0.35);
        max-width: 480px;
        margin-left: auto;
        margin-bottom: 15px;
        font-family: 'Segoe UI', system-ui, sans-serif;
    ">
        <!-- Reloj Análogo -->
        <div style="position: relative; width: 60px; height: 60px;">
            <svg id="analog-clock" width="60" height="60" viewBox="0 0 100 100">
                <circle cx="50
