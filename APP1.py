import base64
import html
import mimetypes
import unicodedata
from collections import deque
from io import BytesIO
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps


# -----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Asociación de Fútbol Nacimiento",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

BUILD_ID = "MENU_MOVIL_VISIBLE_SIN_CAMBIOS_ESCRITORIO_V16"
print(f"[ANFA] Build: {BUILD_ID} | Archivo ejecutado: {Path(__file__).resolve()}")

PREFERRED_EXCEL_NAMES = [
    "Planilla_Maestra_Campeonato_AF_Nacimiento_2026.xlsx",
    "Planilla_Maestra_Campeonato_AF_Nacimiento.xlsx",
]

DEFAULT_SERIES = [
    "Segunda Adulta",
    "Primera Adulta",
    "Super Senior",
    "Senior",
    "Dorados",
    "Honor",
    "Tercera Infantil",
    "Segunda Infantil",
    "Primera Infantil",
]

DEFAULT_CLUBS = [
    "Villa Alegre",
    "U.J. Lautaro",
    "Forestal",
    "Grinvasul",
    "Jota Montt",
    "Inca",
    "Colo Colo",
    "Maestranza",
]

SHEET_DEFAULTS = {
    "Clubes": ["club", "estadio", "fundacion", "presidente", "estado"],
    "Series": ["serie", "orden", "categoria", "estado"],
    "Inscripciones": ["serie", "club", "participa"],
    "Canchas": ["nombre", "direccion", "comuna", "estado"],
    "Jugadores": [
        "id_jugador", "serie", "club", "nombres", "apellido_paterno",
        "apellido_materno", "rut", "fecha_nacimiento", "edad",
        "numero_camiseta", "posicion", "estado_jugador",
    ],
    "Cuerpo_Tecnico": ["serie", "club", "cargo", "nombres", "estado"],
    "Partidos": [
        "serie", "jornada", "fecha", "hora", "estadio", "local", "visita",
        "goles_local", "goles_visita", "id_partido", "temporada",
        "campeonato", "estado", "arbitro", "asistente_1", "asistente_2",
        "publicar_web", "observaciones", "enlace_acta",
    ],
    "Tabla_Posiciones": [
        "serie", "club", "participa", "descuento_puntos", "observaciones",
    ],
    "Goles": [
        "id_gol", "id_partido", "fecha", "serie", "id_jugador", "jugador",
        "club", "minuto", "tipo_gol", "asistencia", "observaciones",
    ],
    "Goleadores": ["id_jugador", "jugador", "club", "serie", "goles", "estado"],
    "Disciplina": [
        "id_evento", "id_partido", "fecha", "serie", "jugador", "id_jugador",
        "club", "tipo_evento", "minuto", "motivo", "arbitro_informante",
        "pasa_a_tribunal", "enlace_informe", "observaciones",
    ],
    "Castigados": [
        "jugador", "club", "serie", "motivo", "fechas_pendientes",
        "id_sancion", "estado", "numero_resolucion", "enlace_resolucion",
        "observaciones",
    ],
    "Multas": [
        "club", "serie", "motivo", "monto", "estado", "id_multa",
        "fecha_emision", "fecha_vencimiento", "monto_pagado", "saldo",
        "fecha_pago", "numero_resolucion", "observaciones",
    ],
    "Arbitros": [
        "id_arbitro", "nombres", "apellido_paterno", "apellido_materno",
        "categoria", "telefono", "correo", "estado",
    ],
    "Noticias": [
        "titulo", "fecha", "resumen", "id_noticia", "tipo", "contenido",
        "autor", "imagen_url", "documento_url", "destacada", "estado",
        "fecha_publicacion",
    ],
    "Documentos": [
        "id_documento", "tipo", "titulo", "temporada", "serie", "club",
        "fecha_documento", "numero_documento", "estado", "enlace_archivo",
        "publicar_web", "observaciones",
    ],
}

TEAM_ALIASES = {
    "VILLA ALEGRE": "Villa Alegre",
    "U J LAUTARO": "U.J. Lautaro",
    "UJ LAUTARO": "U.J. Lautaro",
    "FORESTAL": "Forestal",
    "GRINVASUL": "Grinvasul",
    "JOTA MONTT": "Jota Montt",
    "INCA": "Inca",
    "COLO COLO": "Colo Colo",
    "COLO COLO NACIMIENTO": "Colo Colo",
    "MAESTRANZA": "Maestranza",
}


# -----------------------------------------------------------------------------
# ESTILOS
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --azul: #092a55;
        --azul2: #0d3f7a;
        --dorado: #e6b643;
        --claro: #f4f7fb;
        --borde: #dce4ef;
        --sidebar-fixed-width: 15.5rem;
    }
    html, body, [class*="css"] {
        font-family: Arial, Helvetica, sans-serif;
    }
    .stApp { background: #f6f8fc; }

    /*
       Oculta la cabecera, la franja blanca y la barra de herramientas
       de Streamlit (Share, GitHub, menú, estrella y controles superiores).
       Se incluyen selectores de distintas versiones de Streamlit para
       mantener compatibilidad cuando cambia la estructura interna.
    */
    header[data-testid="stHeader"],
    [data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stAppToolbar"],
    [data-testid="stHeaderActionElements"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    [data-testid="stMainMenu"],
    .stAppHeader,
    .stAppToolbar,
    div[class*="viewerBadge"],
    div[class*="styles_viewerBadge"],
    #MainMenu {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 0 !important;
        height: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        border: 0 !important;
    }
    html, body, #root, .stApp {
        margin: 0 !important;
        padding: 0 !important;
    }
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewBlockContainer"],
    [data-testid="stMain"],
    section.main {
        top: 0 !important;
        padding-top: 0 !important;
        margin-top: 0 !important;
    }
    [data-testid="stAppViewContainer"] > .main,
    [data-testid="stAppViewContainer"] > section {
        top: 0 !important;
        margin-top: 0 !important;
        padding-top: 0 !important;
    }

    /*
       El ancho lo administra el contenedor principal de Streamlit.
       No se fuerza un ancho mínimo global, porque eso hacía que la página
       se extendiera por debajo del menú lateral.
    */
    [data-testid="stAppViewContainer"] {
        box-sizing: border-box !important;
        width: 100vw !important;
        max-width: 100vw !important;
        overflow-x: hidden !important;
    }
    [data-testid="stMain"] {
        box-sizing: border-box !important;
        min-width: 0 !important;
        width: calc(100vw - var(--sidebar-fixed-width)) !important;
        max-width: calc(100vw - var(--sidebar-fixed-width)) !important;
        margin-left: var(--sidebar-fixed-width) !important;
        overflow-x: hidden !important;
    }
    [data-testid="stMain"] > div {
        box-sizing: border-box !important;
        min-width: 0 !important;
        width: 100% !important;
        max-width: 100% !important;
    }
    .block-container {
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        padding: 0 2rem 2.5rem !important;
        margin: 0 !important;
    }
    [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        align-items: stretch !important;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        min-width: 0 !important;
    }
    [data-testid="stDataFrame"] {
        width: 100% !important;
        max-width: 100% !important;
    }
    .hero {
        background: linear-gradient(120deg, #071d3a, #0d4a8d);
        border-radius: 14px;
        min-height: 58px;
        padding: 10px 18px 9px;
        color: white;
        margin-bottom: 10px;
        box-shadow: 0 6px 18px rgba(9,42,85,.14);
        display: flex;
        align-items: center;
    }
    .hero-content {
        width: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .hero h1 {
        margin: 0;
        font-size: clamp(1.25rem, 1.8vw, 1.9rem);
        line-height: 1;
    }
    .hero-footer {
        width: 100%;
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 14px;
        margin-top: 6px;
    }
    .hero p {
        margin: 0;
        opacity: .9;
        font-size: .78rem;
        line-height: 1.1;
    }
    .hero-year {
        margin-left: auto;
        color: #f4c542;
        font-size: .82rem;
        line-height: 1;
        font-weight: 900;
        white-space: nowrap;
    }
    .section-title {
        color: #092a55;
        font-size: 1.6rem;
        font-weight: 800;
        margin: 8px 0 14px;
    }
    .match-card, .club-card, .notice-card, .document-card {
        background: white;
        border: 1px solid #dce4ef;
        border-radius: 15px;
        padding: 17px;
        margin-bottom: 12px;
        box-shadow: 0 5px 16px rgba(20,45,80,.06);
    }
    .clubs-single-row {
        display: flex;
        flex-direction: column;
        gap: 10px;
        width: 100%;
        overflow: visible;
        padding: 2px 0 12px;
    }
    .clubs-single-row .club-card {
        display: grid;
        grid-template-columns: 48px minmax(180px, .8fr) minmax(170px, 1fr) minmax(190px, 1fr) minmax(190px, 1fr);
        align-items: center;
        column-gap: 16px;
        width: 100%;
        min-height: 74px;
        margin-bottom: 0;
        padding: 13px 18px;
    }
    .clubs-single-row .club-card > div:first-child {
        font-size: 1.7rem !important;
        line-height: 1;
        text-align: center;
    }
    .clubs-single-row .club-card h3 {
        margin: 0 !important;
        font-size: 1.08rem;
        line-height: 1.2;
    }
    .clubs-single-row .club-card > div:not(:first-child) {
        color: #3f4856;
        line-height: 1.35;
        overflow-wrap: anywhere;
    }
    .featured-news-grid {
        display: grid;
        grid-template-columns: minmax(0, 312px) minmax(0, 440px);
        justify-content: center;
        align-items: start;
        gap: 18px;
        width: 100%;
        margin: 2px auto 16px;
    }
    .featured-news-card {
        width: 100%;
        margin: 0 auto;
        background: white;
        border: 1px solid #dce4ef;
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 6px 18px rgba(20,45,80,.09);
    }
    .featured-news-card--honor {
        max-width: 312px;
    }
    .featured-news-card--ssenior {
        max-width: 440px;
    }
    .featured-news-card img {
        display: block;
        width: 100%;
        height: auto;
        object-position: center center;
    }
    .featured-news-card--honor img {
        aspect-ratio: 6 / 7;
        object-fit: cover;
    }
    .featured-news-card--ssenior img {
        aspect-ratio: auto;
        object-fit: contain;
        background: white;
    }
    .featured-news-caption {
        padding: 11px 16px 12px;
        color: #092a55;
        font-size: 1.05rem;
        line-height: 1.15;
        font-weight: 900;
        text-align: center;
    }
    .home-standings-wrap {
        width: 100%;
        overflow: hidden;
        background: rgba(255,255,255,.97);
        border: 1px solid #dce4ef;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(20,45,80,.05);
    }
    .home-standings-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        color: #29384c;
        font-size: clamp(.65rem, .74vw, .78rem);
    }
    .home-standings-table th,
    .home-standings-table td {
        padding: 6px 2px;
        text-align: center !important;
        vertical-align: middle !important;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        border-right: 1px solid #e7ebf1;
        border-bottom: 1px solid #e7ebf1;
        line-height: 1.15;
    }
    .home-standings-table th {
        background: #f4f6f9;
        color: #738198;
        font-weight: 700;
    }
    .home-standings-table th:last-child,
    .home-standings-table td:last-child {
        border-right: 0;
    }
    .home-standings-table tbody tr:last-child td {
        border-bottom: 0;
    }
    .home-standings-table th:nth-child(1),
    .home-standings-table td:nth-child(1) { width: 7%; }
    .home-standings-table th:nth-child(2),
    .home-standings-table td:nth-child(2) { width: 22%; }
    .home-standings-table th:nth-child(n+3),
    .home-standings-table td:nth-child(n+3) { width: 8.875%; }
    .match-row {
        display: grid;
        grid-template-columns: 1fr auto 1fr;
        gap: 14px;
        align-items: center;
    }
    .team-local {
        text-align: right;
        font-weight: 750;
        color: #0b2b54;
    }
    .team-away {
        text-align: left;
        font-weight: 750;
        color: #0b2b54;
    }
    .score {
        min-width: 86px;
        text-align: center;
        font-weight: 900;
        background: #092a55;
        color: white;
        border-radius: 10px;
        padding: 8px 10px;
    }
    .meta {
        color: #65758a;
        text-align: center;
        font-size: .88rem;
        padding-top: 10px;
    }
    .kpi {
        background: white;
        border: 1px solid #dce4ef;
        border-radius: 16px;
        padding: 18px;
        min-height: 122px;
        box-shadow: 0 5px 16px rgba(20,45,80,.06);
    }
    .kpi-label {
        color: #65758a;
        font-size: .88rem;
        font-weight: 700;
    }
    .kpi-value {
        color: #092a55;
        font-size: 1.85rem;
        font-weight: 900;
        margin-top: 8px;
        overflow-wrap: anywhere;
    }
    .series-summary {
        background: white;
        border: 1px solid #dce4ef;
        border-radius: 15px;
        padding: 14px 18px;
        margin-bottom: 16px;
    }
    .series-summary b { color: #092a55; }
    /*
       Menú lateral fijo: permanece visible mientras se desplaza
       el contenido principal.
    */
    [data-testid="stSidebar"] {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        bottom: 0 !important;
        width: var(--sidebar-fixed-width) !important;
        min-width: var(--sidebar-fixed-width) !important;
        max-width: var(--sidebar-fixed-width) !important;
        background: #071d3a;
        height: 100vh !important;
        overflow: hidden !important;
        z-index: 1000 !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        height: 100vh !important;
        min-height: 100vh !important;
        overflow-y: auto !important;
        overscroll-behavior: contain;
        padding: 0 1rem 145px !important;
    }
    [data-testid="stSidebarHeader"] {
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        border: 0 !important;
        overflow: visible !important;
    }
    [data-testid="stSidebar"] * { color: white; }
    [data-testid="stSidebar"] div[role="radiogroup"] {
        gap: .48rem !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] > label {
        min-height: 32px !important;
        margin: 0 !important;
        padding: .18rem 0 !important;
        align-items: center !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] > label p {
        line-height: 1.35 !important;
    }
    .sidebar-season {
        position: fixed;
        left: 22px;
        bottom: 24px;
        width: 181px;
        border-top: 1px solid rgba(255,255,255,.22);
        padding-top: 20px;
        color: #9fb0c7;
        font-size: .83rem;
        line-height: 1.55;
        z-index: 1001;
    }
    .sidebar-season strong {
        display: block;
        color: #c3cfdf;
        font-weight: 700;
    }
    [data-testid="stSidebarCollapseButton"],
    [data-testid="collapsedControl"] { display: none !important; }
    .sidebar-logo-wrap {
        position: sticky;
        top: 0;
        z-index: 1003;
        display: flex;
        justify-content: center;
        margin: 0 0 6px;
        padding: 2px 0 4px;
        background: #071d3a;
    }
    .sidebar-logo {
        width: 160px;
        height: 160px;
        object-fit: contain;
        border-radius: 50%;
        background: transparent;
    }
    [data-testid="stMain"],
    [data-testid="stAppViewContainer"] {
        background-color: transparent !important;
    }
    footer { visibility: hidden; }

    /* -------------------------------------------------------------
       FIXTURE: navegación horizontal de jornadas
       ------------------------------------------------------------- */
    .fixture-page-title {
        color: #092a55;
        font-size: 1.9rem;
        line-height: 1.15;
        font-weight: 900;
        margin: 4px 0 14px;
    }
    .fixture-round-header {
        background: transparent;
        border: 0;
        border-radius: 0;
        padding: 8px 0;
        margin: 0;
        min-height: 64px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        box-shadow: none;
    }
    .fixture-round-kicker {
        color: #7a8798;
        font-size: .78rem;
        line-height: 1;
        font-weight: 850;
        letter-spacing: .08em;
        text-transform: uppercase;
        margin-bottom: 7px;
    }
    .fixture-round-title {
        color: #092a55;
        font-size: 1.75rem;
        line-height: 1;
        font-weight: 950;
        margin: 0;
    }
    .fixture-round-title span {
        color: #e13b2f;
    }
    .fixture-program-title {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        background: rgba(255,255,255,.96);
        border: 1px solid #dce4ef;
        border-radius: 16px;
        padding: 16px 20px;
        margin: 22px 0 14px;
        box-shadow: 0 6px 18px rgba(20,45,80,.07);
    }
    .fixture-program-name {
        color: #092a55;
        font-size: 1.28rem;
        line-height: 1.2;
        font-weight: 900;
    }
    .fixture-program-count {
        color: #738198;
        font-size: .9rem;
        font-weight: 800;
        white-space: nowrap;
    }

    /* Botones de fecha: compactos y legibles */
    [data-testid="stMain"] [data-testid="stButton"] button {
        min-height: 48px;
        border-radius: 14px;
        font-weight: 850;
    }

    @media (max-width: 1200px) {
        :root {
            --sidebar-fixed-width: 14rem;
        }
        .block-container {
            padding-left: 1.25rem !important;
            padding-right: 1.25rem !important;
        }
        .hero {
            padding-left: 24px;
            padding-right: 24px;
        }
    }

    /* Navegación alternativa exclusiva para teléfonos. */
    .mobile-primary-nav {
        display: none;
    }

    /* -------------------------------------------------------------
       VISTA MÓVIL
       Estas reglas se activan solo en teléfonos y no modifican PC.
       ------------------------------------------------------------- */
    @media (max-width: 768px) {
        :root {
            --sidebar-fixed-width: min(86vw, 18rem);
        }

        html, body, #root, .stApp,
        [data-testid="stAppViewContainer"] {
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }

        /*
           En teléfonos se usa una navegación propia y siempre visible.
           Así no dependemos del botón lateral interno de Streamlit, que puede
           cambiar o desaparecer según la versión o el navegador integrado.
        */
        header[data-testid="stHeader"],
        [data-testid="stHeader"],
        .stAppHeader,
        [data-testid="stToolbar"],
        [data-testid="stAppToolbar"],
        [data-testid="stHeaderActionElements"],
        [data-testid="stMainMenu"],
        [data-testid="stSidebar"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="stSidebarCollapseButton"],
        [data-testid="collapsedControl"] {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }

        /* El contenido ocupa todo el ancho del teléfono. */
        [data-testid="stMain"] {
            width: 100vw !important;
            max-width: 100vw !important;
            min-width: 0 !important;
            margin-left: 0 !important;
            padding-left: 0 !important;
            overflow-x: hidden !important;
        }
        [data-testid="stMain"] > div,
        [data-testid="stMain"] .block-container {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
        }
        .block-container {
            padding: .7rem .75rem 2rem !important;
        }

        /* Menú principal visible únicamente en la versión móvil. */
        .mobile-primary-nav {
            display: grid !important;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: .45rem;
            width: 100%;
            margin: 0 0 .9rem;
        }
        .mobile-primary-nav .mobile-nav-item {
            grid-column: span 2;
            display: flex;
            min-width: 0;
            min-height: 43px;
            align-items: center;
            justify-content: center;
            padding: .55rem .25rem;
            border: 1px solid #cfd9e7;
            border-radius: 11px;
            background: rgba(255,255,255,.94);
            color: #092a55 !important;
            text-align: center;
            text-decoration: none !important;
            font-size: .76rem;
            line-height: 1.1;
            font-weight: 850;
            box-shadow: 0 3px 9px rgba(20,45,80,.07);
            -webkit-tap-highlight-color: transparent;
        }
        .mobile-primary-nav .mobile-nav-item:nth-child(4),
        .mobile-primary-nav .mobile-nav-item:nth-child(5) {
            grid-column: span 3;
        }
        .mobile-primary-nav .mobile-nav-item.active {
            background: #092a55 !important;
            border-color: #092a55 !important;
            color: white !important;
            box-shadow: 0 4px 12px rgba(9,42,85,.20);
        }
        .mobile-primary-nav .mobile-nav-item:active {
            transform: translateY(1px);
        }

        /* Encabezado compacto y legible. */
        .hero {
            min-height: auto !important;
            padding: 12px 14px 11px !important;
            border-radius: 12px !important;
            margin-bottom: 10px !important;
        }
        .hero h1 {
            font-size: 1.15rem !important;
            line-height: 1.15 !important;
        }
        .hero-footer {
            flex-direction: column !important;
            align-items: stretch !important;
            gap: 5px !important;
            margin-top: 6px !important;
        }
        .hero p {
            font-size: .72rem !important;
            line-height: 1.3 !important;
        }
        .hero-year {
            align-self: flex-end !important;
            margin-left: 0 !important;
            font-size: .78rem !important;
        }

        .section-title,
        .fixture-page-title {
            font-size: 1.32rem !important;
            margin: 8px 0 11px !important;
        }
        .series-summary {
            padding: 11px 13px !important;
            margin-bottom: 12px !important;
        }
        .featured-news-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: .65rem !important;
            margin: 2px auto 13px !important;
        }
        .featured-news-card {
            width: 100% !important;
            max-width: 100% !important;
            margin: 0 auto !important;
            border-radius: 13px !important;
        }
        .featured-news-caption {
            padding: 10px 12px !important;
            font-size: .95rem !important;
        }
        .home-standings-table {
            font-size: .61rem !important;
        }
        .home-standings-table th,
        .home-standings-table td {
            padding: 5px 1px !important;
        }

        /* Las columnas de contenido se apilan en móvil. */
        [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: .65rem !important;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"] {
            flex: 1 1 100% !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
        }

        /* Navegación anterior/título/siguiente del fixture en una sola fila. */
        [data-testid="stHorizontalBlock"]:has([data-testid="stButton"]):not(:has(> [data-testid="column"]:nth-child(4))) {
            display: grid !important;
            grid-template-columns: 48px minmax(0, 1fr) 48px !important;
            align-items: center !important;
            gap: .4rem !important;
        }
        [data-testid="stHorizontalBlock"]:has([data-testid="stButton"]):not(:has(> [data-testid="column"]:nth-child(4))) > [data-testid="column"] {
            width: auto !important;
            max-width: none !important;
            min-width: 0 !important;
        }

        /* Fechas del fixture desplazables horizontalmente, sin comprimirlas. */
        [data-testid="stHorizontalBlock"]:has(> [data-testid="column"]:nth-child(4)):has([data-testid="stButton"]) {
            display: flex !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            gap: .45rem !important;
            padding: 2px 1px 7px !important;
            scroll-snap-type: x proximity;
            scrollbar-width: thin;
        }
        [data-testid="stHorizontalBlock"]:has(> [data-testid="column"]:nth-child(4)):has([data-testid="stButton"]) > [data-testid="column"] {
            flex: 0 0 60px !important;
            width: 60px !important;
            min-width: 60px !important;
            max-width: 60px !important;
            scroll-snap-align: start;
        }
        [data-testid="stMain"] [data-testid="stButton"] button {
            min-height: 43px !important;
            border-radius: 11px !important;
            padding-left: .45rem !important;
            padding-right: .45rem !important;
        }

        .fixture-round-header {
            min-height: 54px !important;
            padding: 5px 0 !important;
        }
        .fixture-round-kicker {
            font-size: .68rem !important;
            margin-bottom: 5px !important;
        }
        .fixture-round-title {
            font-size: 1.35rem !important;
        }
        .fixture-program-title {
            flex-direction: column !important;
            align-items: flex-start !important;
            gap: 6px !important;
            padding: 13px 14px !important;
            margin: 16px 0 11px !important;
            border-radius: 13px !important;
        }
        .fixture-program-name {
            font-size: 1rem !important;
        }
        .fixture-program-count {
            font-size: .78rem !important;
        }

        .match-card, .club-card, .notice-card, .document-card {
            border-radius: 13px !important;
            padding: 13px !important;
            margin-bottom: 10px !important;
        }
        .clubs-single-row {
            display: flex !important;
            flex-direction: column !important;
            gap: .65rem !important;
            padding-bottom: 10px !important;
        }
        .clubs-single-row .club-card {
            display: grid !important;
            grid-template-columns: 42px minmax(0, 1fr) !important;
            gap: 7px 10px !important;
            min-height: 0 !important;
            margin-bottom: 0 !important;
            padding: 12px 13px !important;
        }
        .clubs-single-row .club-card > div:first-child {
            grid-row: 1 / span 4;
            align-self: start;
            font-size: 1.45rem !important;
        }
        .clubs-single-row .club-card h3 {
            grid-column: 2;
            font-size: 1rem !important;
        }
        .clubs-single-row .club-card > div:not(:first-child) {
            grid-column: 2;
            font-size: .84rem;
        }
        .match-row {
            grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr) !important;
            gap: 7px !important;
        }
        .team-local,
        .team-away {
            font-size: .84rem !important;
            line-height: 1.2 !important;
            overflow-wrap: anywhere !important;
        }
        .score {
            min-width: 58px !important;
            padding: 7px 6px !important;
            border-radius: 9px !important;
            font-size: .88rem !important;
        }
        .meta {
            font-size: .74rem !important;
            line-height: 1.45 !important;
            padding-top: 9px !important;
            overflow-wrap: anywhere !important;
        }

        /* Formularios, tablas y avisos ajustados al ancho del teléfono. */
        [data-testid="stSelectbox"],
        [data-testid="stTextInput"],
        [data-testid="stDataFrame"],
        [data-testid="stAlert"] {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
        }
        [data-testid="stDataFrame"] {
            overflow-x: auto !important;
        }
        [data-testid="stDataFrame"] iframe,
        [data-testid="stDataFrame"] canvas {
            max-width: 100% !important;
        }
        .kpi {
            min-height: 98px !important;
            padding: 14px !important;
        }
        .kpi-value {
            font-size: 1.5rem !important;
        }
    }

    @media (max-width: 390px) {
        .block-container {
            padding-left: .55rem !important;
            padding-right: .55rem !important;
        }
        .hero h1 {
            font-size: 1.03rem !important;
        }
        .hero p {
            font-size: .68rem !important;
        }
        .team-local,
        .team-away {
            font-size: .78rem !important;
        }
        .score {
            min-width: 52px !important;
            font-size: .8rem !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# UTILIDADES DE ARCHIVOS
# -----------------------------------------------------------------------------
def find_excel_file() -> Path | None:
    """Busca la planilla primero en la carpeta principal y luego en data."""
    search_dirs = [BASE_DIR, DATA_DIR]

    for folder in search_dirs:
        for filename in PREFERRED_EXCEL_NAMES:
            candidate = folder / filename
            if candidate.exists() and candidate.is_file():
                return candidate

    for folder in search_dirs:
        candidates = sorted(
            file
            for file in folder.glob("*.xlsx")
            if file.is_file() and not file.name.startswith("~$")
        )
        if candidates:
            return candidates[0]

    return None


def find_logo_file() -> Path | None:
    """Detecta la imagen logoanfa más reciente, sin depender de su extensión."""
    valid_extensions = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
    candidates = [
        file
        for file in BASE_DIR.glob("logoanfa*")
        if file.is_file() and file.suffix.lower() in valid_extensions
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda file: file.stat().st_mtime)


def find_sello_file() -> Path | None:
    """Detecta la imagen sello sin depender de una extensión específica."""
    preferred_extensions = [".png", ".jpg", ".jpeg", ".webp", ".gif"]
    for extension in preferred_extensions:
        candidate = BASE_DIR / f"sello{extension}"
        if candidate.exists():
            return candidate
    candidates = sorted(file for file in BASE_DIR.glob("sello*") if file.is_file())
    return candidates[0] if candidates else None


def find_honor_file() -> Path | None:
    """Detecta la fotografía de la Selección Serie Honor en la carpeta principal."""
    valid_extensions = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
    candidates = [
        file
        for file in BASE_DIR.glob("honor*")
        if file.is_file() and file.suffix.lower() in valid_extensions
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda file: file.stat().st_mtime)


def find_ssenior_file() -> Path | None:
    """Detecta la fotografía de la Selección Super Senior en la carpeta principal."""
    valid_extensions = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
    candidates = [
        file
        for file in BASE_DIR.glob("ssenior*")
        if file.is_file() and file.suffix.lower() in valid_extensions
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda file: file.stat().st_mtime)


def image_data_uri(path: Path | None) -> str:
    if path is None or not path.exists():
        return ""
    mime_type, _ = mimetypes.guess_type(path.name)
    mime_type = mime_type or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"


def photo_data_uri(path: Path | None, max_width: int = 900) -> str:
    """Convierte fotografías a un JPEG liviano para evitar que Streamlit rompa el HTML."""
    if path is None or not path.exists():
        return ""

    try:
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")

            if image.width > max_width:
                new_height = max(1, round(image.height * max_width / image.width))
                image = image.resize((max_width, new_height), Image.Resampling.LANCZOS)

            buffer = BytesIO()
            image.save(
                buffer,
                format="JPEG",
                quality=84,
                optimize=True,
                progressive=True,
            )
            encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
            return f"data:image/jpeg;base64,{encoded}"
    except Exception:
        return image_data_uri(path)


def logo_data_uri_without_white_border(path: Path | None) -> str:
    """Elimina solo el blanco exterior conectado a los bordes y conserva el blanco interior."""
    if path is None or not path.exists():
        return ""

    try:
        image = Image.open(path).convert("RGBA")
        width, height = image.size
        pixels = image.load()
        visited = set()
        pending = deque()

        def is_exterior_white(x: int, y: int) -> bool:
            red, green, blue, alpha = pixels[x, y]
            return alpha > 0 and red >= 238 and green >= 238 and blue >= 238

        for x in range(width):
            if is_exterior_white(x, 0):
                pending.append((x, 0))
            if is_exterior_white(x, height - 1):
                pending.append((x, height - 1))
        for y in range(height):
            if is_exterior_white(0, y):
                pending.append((0, y))
            if is_exterior_white(width - 1, y):
                pending.append((width - 1, y))

        while pending:
            x, y = pending.popleft()
            if (x, y) in visited or not (0 <= x < width and 0 <= y < height):
                continue
            if not is_exterior_white(x, y):
                continue

            visited.add((x, y))
            red, green, blue, _ = pixels[x, y]
            pixels[x, y] = (red, green, blue, 0)

            pending.extend(((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)))

        alpha = image.getchannel("A")
        bbox = alpha.getbbox()
        if bbox:
            image = image.crop(bbox)

        buffer = BytesIO()
        image.save(buffer, format="PNG", optimize=True)
        encoded = base64.b64encode(buffer.getvalue()).decode("utf-8")
        return f"data:image/png;base64,{encoded}"
    except Exception:
        return image_data_uri(path)


def finalized_mask(series: pd.Series) -> pd.Series:
    """Entrega una máscara booleana estable, incluso con tipos string[pyarrow]."""
    normalized = (
        series.astype("string")
        .fillna("")
        .str.strip()
        .str.lower()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("ascii")
    )
    return normalized.isin(["finalizado", "jugado", "terminado"]).astype(bool)


def normalize_column_name(value: object) -> str:
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("/", "_").replace("-", "_").replace(" ", "_")
    while "__" in text:
        text = text.replace("__", "_")
    return text.strip("_")


def clean_text(value: object, default: str = "") -> str:
    if pd.isna(value):
        return default
    text = str(value).strip()
    return default if text.lower() in {"nan", "nat", "none"} else text


def safe_html(value: object, default: str = "") -> str:
    return html.escape(clean_text(value, default))


def safe_sort_by_date(
    dataframe: pd.DataFrame,
    column: str,
    ascending: bool = False,
) -> pd.DataFrame:
    """Ordena por fecha sin fallar cuando la hoja está vacía o no contiene la columna."""
    result = dataframe.copy()
    if column not in result.columns:
        result[column] = pd.NaT
    result[column] = pd.to_datetime(result[column], errors="coerce")
    return result.sort_values(column, ascending=ascending, na_position="last")


def canonical_team_name(value: object) -> object:
    if pd.isna(value):
        return value
    original = clean_text(value)
    cleaned = " ".join(original.upper().replace(".", " ").split())
    return TEAM_ALIASES.get(cleaned, original)


def is_yes(value: object) -> bool:
    normalized = clean_text(value).lower()
    normalized = unicodedata.normalize("NFKD", normalized)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    return normalized in {"si", "yes", "1", "true"}


def is_finalized(value: object) -> bool:
    normalized = clean_text(value).lower()
    normalized = unicodedata.normalize("NFKD", normalized)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    return normalized in {"finalizado", "jugado", "terminado"}


def empty_sheet(sheet_name: str) -> pd.DataFrame:
    return pd.DataFrame(columns=SHEET_DEFAULTS.get(sheet_name, []))


def clean_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    if df is None or df.empty:
        return empty_sheet(sheet_name)

    cleaned = df.copy()
    cleaned.columns = [normalize_column_name(column) for column in cleaned.columns]
    cleaned = cleaned.dropna(how="all")

    for required_column in SHEET_DEFAULTS.get(sheet_name, []):
        if required_column not in cleaned.columns:
            cleaned[required_column] = pd.NA

    key_columns = {
        "Clubes": ["club"],
        "Series": ["serie"],
        "Inscripciones": ["serie", "club"],
        "Canchas": ["nombre"],
        "Jugadores": ["id_jugador", "nombres", "rut"],
        "Cuerpo_Tecnico": ["club", "nombres"],
        "Partidos": ["local", "visita"],
        "Tabla_Posiciones": ["serie", "club"],
        "Goles": ["id_gol", "jugador", "id_jugador"],
        "Goleadores": ["jugador"],
        "Disciplina": ["id_evento", "jugador"],
        "Castigados": ["jugador"],
        "Multas": ["club", "motivo"],
        "Arbitros": ["id_arbitro", "nombres"],
        "Noticias": ["titulo"],
        "Documentos": ["titulo", "id_documento"],
    }

    keys = [column for column in key_columns.get(sheet_name, []) if column in cleaned.columns]
    if keys:
        cleaned = cleaned[cleaned[keys].notna().any(axis=1)]

    return cleaned.reset_index(drop=True)


def read_sheet_with_detected_header(
    workbook: pd.ExcelFile,
    sheet_name: str,
) -> pd.DataFrame:
    """Detecta la fila de encabezados aunque la hoja tenga título y subtítulo."""
    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
    if raw.empty:
        return empty_sheet(sheet_name)

    expected = set(SHEET_DEFAULTS.get(sheet_name, []))
    best_row = None
    best_score = 0

    for row_index in range(min(12, len(raw))):
        normalized_values = {
            normalize_column_name(value)
            for value in raw.iloc[row_index].tolist()
            if pd.notna(value) and clean_text(value)
        }
        score = len(normalized_values & expected)
        if score > best_score:
            best_score = score
            best_row = row_index

    if best_row is None or best_score == 0:
        return empty_sheet(sheet_name)

    headers = []
    for position, value in enumerate(raw.iloc[best_row].tolist(), start=1):
        header = normalize_column_name(value) if pd.notna(value) else ""
        headers.append(header or f"columna_{position}")

    dataframe = raw.iloc[best_row + 1:].copy()
    dataframe.columns = headers
    return dataframe.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_workbook_data(excel_path: str, modified_time: float) -> dict[str, pd.DataFrame]:
    """Carga solo las hojas necesarias y se invalida al cambiar el Excel."""
    del modified_time

    workbook = pd.ExcelFile(excel_path, engine="openpyxl")
    available_sheets = set(workbook.sheet_names)
    data: dict[str, pd.DataFrame] = {}

    for sheet_name in SHEET_DEFAULTS:
        if sheet_name in available_sheets:
            raw = read_sheet_with_detected_header(workbook, sheet_name)
            data[sheet_name] = clean_sheet(raw, sheet_name)
        else:
            data[sheet_name] = empty_sheet(sheet_name)

    return data


# -----------------------------------------------------------------------------
# CARGA Y NORMALIZACIÓN DE DATOS
# -----------------------------------------------------------------------------
excel_file = find_excel_file()
logo_file = find_logo_file()
sello_file = find_sello_file()
honor_file = find_honor_file()
ssenior_file = find_ssenior_file()

if excel_file is None:
    st.error(
        "No se encontró la planilla Excel en la carpeta principal ni dentro de **data**. "
        "Guarda el archivo `Planilla_Maestra_Campeonato_AF_Nacimiento_2026.xlsx` junto a `app.py`."
    )
    st.stop()

try:
    workbook_data = load_workbook_data(str(excel_file), excel_file.stat().st_mtime)
except Exception as error:
    st.error("No fue posible leer la planilla Excel.")
    st.exception(error)
    st.stop()

clubs = workbook_data["Clubes"].copy()
series_df = workbook_data["Series"].copy()
inscriptions = workbook_data["Inscripciones"].copy()
players = workbook_data["Jugadores"].copy()
matches = workbook_data["Partidos"].copy()
positions_master = workbook_data["Tabla_Posiciones"].copy()
goals = workbook_data["Goles"].copy()
sanctions = workbook_data["Castigados"].copy()
fines = workbook_data["Multas"].copy()
news = workbook_data["Noticias"].copy()
documents = workbook_data["Documentos"].copy()

# Clubes
if clubs.empty:
    clubs = pd.DataFrame(
        {
            "club": DEFAULT_CLUBS,
            "estadio": ["Por definir"] * len(DEFAULT_CLUBS),
            "fundacion": ["Por definir"] * len(DEFAULT_CLUBS),
            "presidente": ["Por definir"] * len(DEFAULT_CLUBS),
            "estado": ["Activo"] * len(DEFAULT_CLUBS),
        }
    )
else:
    clubs["club"] = clubs["club"].apply(canonical_team_name)
    clubs = clubs.drop_duplicates(subset="club", keep="first")

for column in ["estadio", "fundacion", "presidente", "estado"]:
    if column not in clubs.columns:
        clubs[column] = "Por definir"

# Series
if series_df.empty:
    series_order = DEFAULT_SERIES.copy()
else:
    series_df["serie"] = series_df["serie"].apply(clean_text)
    if "estado" in series_df.columns:
        series_df = series_df[
            ~series_df["estado"].astype(str).str.strip().str.lower().eq("inactivo")
        ]
    if "orden" in series_df.columns:
        series_df["orden"] = pd.to_numeric(series_df["orden"], errors="coerce")
        series_df = series_df.sort_values(["orden", "serie"], na_position="last")
    series_order = [value for value in series_df["serie"].tolist() if value]
    series_order = list(dict.fromkeys(series_order)) or DEFAULT_SERIES.copy()

# Partidos
for column in ["local", "visita"]:
    matches[column] = matches[column].apply(canonical_team_name)

matches["serie"] = matches["serie"].apply(clean_text)
matches["fecha"] = pd.to_datetime(matches["fecha"], errors="coerce")
matches["jornada"] = pd.to_numeric(matches["jornada"], errors="coerce")
matches["goles_local"] = pd.to_numeric(matches["goles_local"], errors="coerce")
matches["goles_visita"] = pd.to_numeric(matches["goles_visita"], errors="coerce")
matches["estado"] = matches["estado"].apply(clean_text)

# Otras hojas
for frame in [inscriptions, players, positions_master, goals, sanctions, fines, news, documents]:
    if "serie" in frame.columns:
        frame["serie"] = frame["serie"].apply(clean_text)
    if "club" in frame.columns:
        frame["club"] = frame["club"].apply(canonical_team_name)

for frame, date_columns in [
    (players, ["fecha_nacimiento", "fecha_inscripcion"]),
    (goals, ["fecha"]),
    (news, ["fecha", "fecha_publicacion"]),
    (documents, ["fecha_documento"]),
    (fines, ["fecha_emision", "fecha_vencimiento", "fecha_pago"]),
]:
    for column in date_columns:
        if column in frame.columns:
            frame[column] = pd.to_datetime(frame[column], errors="coerce")

for column in ["monto", "monto_pagado", "saldo"]:
    if column in fines.columns:
        fines[column] = pd.to_numeric(fines[column], errors="coerce").fillna(0)

extra_series = [
    value
    for value in matches["serie"].dropna().astype(str).unique().tolist()
    if value and value not in series_order
]
available_series = series_order + extra_series


# -----------------------------------------------------------------------------
# CÁLCULOS
# -----------------------------------------------------------------------------
def registered_teams_for_series(series_name: str) -> list[str]:
    if not inscriptions.empty:
        subset = inscriptions[inscriptions["serie"] == series_name].copy()
        if "participa" in subset.columns:
            subset = subset[subset["participa"].apply(is_yes)]
        teams = [clean_text(value) for value in subset["club"].dropna().tolist()]
        teams = [canonical_team_name(value) for value in teams if value]
        if teams:
            return list(dict.fromkeys(teams))

    subset_matches = matches[matches["serie"] == series_name]
    teams = list(subset_matches["local"].dropna()) + list(subset_matches["visita"].dropna())
    return sorted(set(clean_text(team) for team in teams if clean_text(team)))


def point_deductions_for_series(series_name: str) -> dict[str, int]:
    if positions_master.empty or "descuento_puntos" not in positions_master.columns:
        return {}

    subset = positions_master[positions_master["serie"] == series_name].copy()
    subset["descuento_puntos"] = pd.to_numeric(
        subset["descuento_puntos"], errors="coerce"
    ).fillna(0)

    return {
        clean_text(row["club"]): int(row["descuento_puntos"])
        for _, row in subset.iterrows()
        if clean_text(row.get("club"))
    }


def standings_table(series_name: str) -> pd.DataFrame:
    teams = registered_teams_for_series(series_name)
    stats = {
        team: {
            "Club": team,
            "PJ": 0,
            "PG": 0,
            "PE": 0,
            "PP": 0,
            "GF": 0,
            "GC": 0,
            "DG": 0,
            "DESC": 0,
            "PTS": 0,
        }
        for team in teams
    }

    series_matches = matches[matches["serie"] == series_name].copy()
    played = series_matches[
        finalized_mask(series_matches["estado"])
        & series_matches["goles_local"].notna()
        & series_matches["goles_visita"].notna()
    ]

    for _, row in played.iterrows():
        local = clean_text(row["local"])
        visita = clean_text(row["visita"])
        if not local or not visita:
            continue

        stats.setdefault(
            local,
            {"Club": local, "PJ": 0, "PG": 0, "PE": 0, "PP": 0, "GF": 0, "GC": 0, "DG": 0, "DESC": 0, "PTS": 0},
        )
        stats.setdefault(
            visita,
            {"Club": visita, "PJ": 0, "PG": 0, "PE": 0, "PP": 0, "GF": 0, "GC": 0, "DG": 0, "DESC": 0, "PTS": 0},
        )

        goals_local = int(row["goles_local"])
        goals_away = int(row["goles_visita"])

        stats[local]["PJ"] += 1
        stats[visita]["PJ"] += 1
        stats[local]["GF"] += goals_local
        stats[local]["GC"] += goals_away
        stats[visita]["GF"] += goals_away
        stats[visita]["GC"] += goals_local

        if goals_local > goals_away:
            stats[local]["PG"] += 1
            stats[local]["PTS"] += 3
            stats[visita]["PP"] += 1
        elif goals_away > goals_local:
            stats[visita]["PG"] += 1
            stats[visita]["PTS"] += 3
            stats[local]["PP"] += 1
        else:
            stats[local]["PE"] += 1
            stats[visita]["PE"] += 1
            stats[local]["PTS"] += 1
            stats[visita]["PTS"] += 1

    deductions = point_deductions_for_series(series_name)
    for team, values in stats.items():
        values["DG"] = values["GF"] - values["GC"]
        values["DESC"] = deductions.get(team, 0)
        values["PTS"] -= values["DESC"]

    table = pd.DataFrame(stats.values())
    if table.empty:
        return table

    table = table.sort_values(
        ["PTS", "DG", "GF", "Club"],
        ascending=[False, False, False, True],
    ).reset_index(drop=True)
    table.insert(0, "Pos.", range(1, len(table) + 1))
    return table


def render_match_card(row: pd.Series) -> None:
    played = (
        pd.notna(row.get("goles_local"))
        and pd.notna(row.get("goles_visita"))
        and is_finalized(row.get("estado"))
    )
    score = (
        f'{int(row["goles_local"])} - {int(row["goles_visita"])}'
        if played
        else "VS"
    )
    fecha = row.get("fecha")
    fecha_text = fecha.strftime("%d-%m-%Y") if pd.notna(fecha) else "Por confirmar"
    hora_text = safe_html(row.get("hora"), "Por confirmar")
    stadium_text = safe_html(row.get("estadio"), "Por definir")
    status_text = safe_html(row.get("estado"), "Programado")

    st.markdown(
        f"""
        <div class="match-card">
          <div class="match-row">
            <div class="team-local">{safe_html(row.get('local'))}</div>
            <div class="score">{score}</div>
            <div class="team-away">{safe_html(row.get('visita'))}</div>
          </div>
          <div class="meta">
            📅 {fecha_text} · 🕒 {hora_text} · 📍 {stadium_text} · {status_text}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def format_currency(value: object) -> str:
    try:
        return f"${float(value):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "$0"


# -----------------------------------------------------------------------------
# NAVEGACIÓN Y ENCABEZADO
# -----------------------------------------------------------------------------
logo_uri = logo_data_uri_without_white_border(logo_file)
sello_uri = image_data_uri(sello_file)
honor_uri = photo_data_uri(honor_file)
ssenior_uri = photo_data_uri(ssenior_file)

if sello_uri:
    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            background-image:
                linear-gradient(rgba(246, 248, 252, 0.72), rgba(246, 248, 252, 0.72)),
                url("{sello_uri}") !important;
            background-repeat: no-repeat !important;
            background-position: center center !important;
            background-size: 100% 100% !important;
            background-attachment: fixed !important;
        }}
        [data-testid="stMain"],
        [data-testid="stMain"] > div,
        .stApp {{
            background-color: transparent !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

MENU_OPTIONS = [
    "Noticias",
    "Fixture",
    "Posiciones",
    "Clubes",
    "Goleadores",
]

# Los botones móviles navegan mediante el parámetro ?menu=.
# El valor se comparte con el radio del escritorio sin alterar su diseño.
requested_menu = st.query_params.get("menu")
if isinstance(requested_menu, list):
    requested_menu = requested_menu[0] if requested_menu else None

if requested_menu in MENU_OPTIONS:
    last_request = st.session_state.get("_last_mobile_menu_request")
    if requested_menu != last_request:
        st.session_state["main_navigation"] = requested_menu
        st.session_state["_last_mobile_menu_request"] = requested_menu

with st.sidebar:
    if logo_uri:
        st.markdown(
            f'<div class="sidebar-logo-wrap"><img class="sidebar-logo" src="{logo_uri}" alt="Logo ANFA Nacimiento"></div>',
            unsafe_allow_html=True,
        )

    st.markdown("## ⚽ ANFA Nacimiento")
    st.caption("Asociación de Fútbol Nacimiento")

    menu = st.radio(
        "Navegación",
        MENU_OPTIONS,
        label_visibility="collapsed",
        key="main_navigation",
    )

    st.markdown(
        """
        <div class="sidebar-season">
            <strong>Temporada 2026 · Nacimiento</strong>
            Región del Biobío
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="hero">
      <div class="hero-content">
        <h1>Asociación de Fútbol Nacimiento</h1>
        <div class="hero-footer">
          <p>Fixture, resultados, posiciones y gestión de las competencias locales.</p>
          <div class="hero-year">2026</div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

mobile_nav_items = []
for option in MENU_OPTIONS:
    active_class = " active" if option == menu else ""
    current_page = ' aria-current="page"' if option == menu else ""
    mobile_nav_items.append(
        f'<a class="mobile-nav-item{active_class}" '
        f'href="?menu={option}" target="_self"{current_page}>{option}</a>'
    )

st.markdown(
    '<nav class="mobile-primary-nav" aria-label="Menú principal">'
    + ''.join(mobile_nav_items)
    + '</nav>',
    unsafe_allow_html=True,
)

if not available_series:
    available_series = DEFAULT_SERIES.copy()

if menu == "Fixture":
    st.markdown(
        '<div class="fixture-page-title">Fixture oficial</div>',
        unsafe_allow_html=True,
    )

default_index = available_series.index("Primera Adulta") if "Primera Adulta" in available_series else 0
serie = None

# En Noticias el filtro se muestra después de la fotografía de la Serie Honor.
if menu not in {"Noticias", "Clubes"}:
    serie = st.selectbox(
        "Serie",
        available_series,
        index=default_index,
        format_func=lambda value: value.upper(),
        key="series_filter",
    )

if menu not in {"Fixture", "Noticias", "Clubes"}:
    st.markdown(
        f"""
        <div class="series-summary">
            <b>Serie seleccionada:</b> {safe_html(serie).upper()}
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# PÁGINAS
# -----------------------------------------------------------------------------
if menu == "Noticias":
    # El HTML se construye sin sangrías ni saltos de línea iniciales.
    # Streamlit/Markdown interpretaba la segunda tarjeta como bloque de código.
    featured_cards = []

    if honor_uri:
        featured_cards.append(
            '<div class="featured-news-card featured-news-card--honor">'
            f'<img src="{honor_uri}" alt="Selección Serie Honor 2026">'
            '<div class="featured-news-caption">Selección Serie Honor 2026</div>'
            '</div>'
        )

    if ssenior_uri:
        featured_cards.append(
            '<div class="featured-news-card featured-news-card--ssenior">'
            f'<img src="{ssenior_uri}" alt="Selección Super Senior 2025">'
            '<div class="featured-news-caption">Selección Super Senior 2025</div>'
            '</div>'
        )

    if featured_cards:
        featured_news_html = (
            '<div class="featured-news-grid">'
            + ''.join(featured_cards)
            + '</div>'
        )
        st.markdown(featured_news_html, unsafe_allow_html=True)

    serie = st.selectbox(
        "Serie",
        available_series,
        index=default_index,
        format_func=lambda value: value.upper(),
        key="series_filter",
    )

    filtered = matches[matches["serie"] == serie].copy()
    played_mask = (
        finalized_mask(filtered["estado"])
        & filtered["goles_local"].notna()
        & filtered["goles_visita"].notna()
    )
    standings = standings_table(serie)
    left, right = st.columns([1.0, 1.65])

    with left:
        st.markdown('<div class="section-title">Próximos partidos</div>', unsafe_allow_html=True)
        upcoming = filtered[~played_mask].copy()
        upcoming = upcoming.sort_values(["fecha", "jornada"], na_position="last").head(5)

        if upcoming.empty:
            st.info("No existen partidos pendientes cargados para esta serie.")
        else:
            for _, match in upcoming.iterrows():
                render_match_card(match)

    with right:
        st.markdown('<div class="section-title">Tabla de posiciones</div>', unsafe_allow_html=True)
        if standings.empty:
            st.info("Aún no existen clubes inscritos o partidos asignados a esta serie.")
        else:
            home_table = standings.drop(columns=["DESC"], errors="ignore").head(8)
            home_columns = ["Pos.", "Club", "PJ", "PG", "PE", "PP", "GF", "GC", "DG", "PTS"]
            home_table = home_table[[column for column in home_columns if column in home_table.columns]]

            header_html = "".join(
                f"<th>{safe_html(column)}</th>" for column in home_table.columns
            )
            rows_html = []
            for _, position_row in home_table.iterrows():
                cells = []
                for column in home_table.columns:
                    value = position_row[column]
                    if column != "Club" and pd.notna(value):
                        try:
                            value = int(value)
                        except (TypeError, ValueError):
                            pass
                    cells.append(f"<td>{safe_html(value)}</td>")
                rows_html.append("<tr>" + "".join(cells) + "</tr>")

            standings_html = (
                '<div class="home-standings-wrap">'
                '<table class="home-standings-table">'
                f'<thead><tr>{header_html}</tr></thead>'
                f'<tbody>{"".join(rows_html)}</tbody>'
                '</table>'
                '</div>'
            )
            st.markdown(standings_html, unsafe_allow_html=True)

    st.markdown('<div class="section-title">Últimas noticias</div>', unsafe_allow_html=True)
    public_news = news.copy()
    if "estado" in public_news.columns:
        public_news = public_news[
            public_news["estado"].apply(lambda value: clean_text(value).lower() in {"publicada", "", "nan"})
        ]
    public_news = safe_sort_by_date(public_news, "fecha", ascending=False)

    if public_news.empty:
        st.info("Todavía no existen noticias publicadas.")
    else:
        for _, item in public_news.head(3).iterrows():
            news_date = item.get("fecha")
            date_text = news_date.strftime("%d-%m-%Y") if pd.notna(news_date) else "Fecha por definir"
            st.markdown(
                f"""
                <div class="notice-card">
                    <b>{safe_html(item.get('titulo'))}</b><br>
                    <span style="color:#65758a">{date_text}</span>
                    <p>{safe_html(item.get('resumen'))}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

elif menu == "Fixture":
    dataframe = matches[matches["serie"] == serie].copy()

    if dataframe.empty:
        st.info(
            "No existen partidos asignados a esta serie. Completa la columna `serie` "
            "en la hoja Partidos del Excel."
        )
    else:
        dataframe = dataframe.sort_values(
            ["jornada", "fecha", "hora"],
            na_position="last",
        )

        valid_rounds = [
            int(value)
            for value in dataframe["jornada"].dropna().sort_values().unique().tolist()
        ]

        if not valid_rounds:
            st.info("Los partidos cargados no tienen una jornada asignada.")
        else:
            state_key = f"fixture_round_{normalize_column_name(serie)}"
            selected_round = st.session_state.get(state_key, valid_rounds[0])

            if selected_round not in valid_rounds:
                selected_round = valid_rounds[0]
                st.session_state[state_key] = selected_round

            selected_position = valid_rounds.index(selected_round)
            previous_round = valid_rounds[max(0, selected_position - 1)]
            next_round = valid_rounds[min(len(valid_rounds) - 1, selected_position + 1)]

            previous_col, title_col, next_col = st.columns([1, 8, 1], gap="small")

            with previous_col:
                previous_disabled = selected_position == 0
                if st.button(
                    "❮",
                    key=f"fixture_previous_{normalize_column_name(serie)}",
                    disabled=previous_disabled,
                    width="stretch",
                    help="Fecha anterior",
                ):
                    st.session_state[state_key] = previous_round
                    st.rerun()

            with title_col:
                st.markdown(
                    f"""
                    <div class="fixture-round-header">
                        <div class="fixture-round-kicker">Jornada</div>
                        <div class="fixture-round-title">
                            Fecha <span>{selected_round}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with next_col:
                next_disabled = selected_position == len(valid_rounds) - 1
                if st.button(
                    "❯",
                    key=f"fixture_next_{normalize_column_name(serie)}",
                    disabled=next_disabled,
                    width="stretch",
                    help="Fecha siguiente",
                ):
                    st.session_state[state_key] = next_round
                    st.rerun()

            # Todas las fechas permanecen visibles en una sola fila horizontal.
            round_columns = st.columns(len(valid_rounds), gap="small")
            for column, round_number in zip(round_columns, valid_rounds):
                with column:
                    is_selected = round_number == selected_round
                    if st.button(
                        f"F{round_number}",
                        key=f"fixture_round_button_{normalize_column_name(serie)}_{round_number}",
                        type="primary" if is_selected else "secondary",
                        width="stretch",
                        help=f"Ver programación de la Fecha {round_number}",
                    ):
                        st.session_state[state_key] = round_number
                        st.rerun()

            selected_matches = dataframe[
                dataframe["jornada"] == selected_round
            ].copy()
            selected_matches = selected_matches.sort_values(
                ["fecha", "hora", "estadio"],
                na_position="last",
            )

            st.markdown(
                f"""
                <div class="fixture-program-title">
                    <div class="fixture-program-name">
                        🏆 {safe_html(serie).upper()} · FECHA {selected_round}
                    </div>
                    <div class="fixture-program-count">
                        {len(selected_matches)}
                        {"PARTIDO" if len(selected_matches) == 1 else "PARTIDOS"}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if selected_matches.empty:
                st.info("No existen partidos programados para esta fecha.")
            else:
                for _, match in selected_matches.iterrows():
                    render_match_card(match)

elif menu == "Posiciones":
    st.markdown('<div class="section-title">Tabla de posiciones</div>', unsafe_allow_html=True)
    table = standings_table(serie)

    if table.empty:
        st.info(
            "No hay datos para generar la tabla. Marca `Sí` en la columna participa "
            "de la hoja Inscripciones o asigna partidos a esta serie."
        )
    else:
        st.dataframe(
            table,
            hide_index=True,
            use_container_width=True,
            height=38 + (len(table) * 35),
            column_config={
                "Pos.": st.column_config.NumberColumn(width="small"),
                "Club": st.column_config.TextColumn(width="large"),
                "DESC": st.column_config.NumberColumn("Desc.", help="Descuento de puntos"),
                "PTS": st.column_config.NumberColumn("PTS", help="Puntos finales"),
            },
        )
        st.caption(
            "Criterios: puntos, diferencia de gol y goles a favor. "
            "Los descuentos se leen desde la hoja Tabla_Posiciones."
        )

elif menu == "Clubes":
    st.markdown('<div class="section-title">Clubes asociados</div>', unsafe_allow_html=True)

    dataframe = clubs.copy()
    dataframe["_orden_club"] = (
        dataframe["club"]
        .astype("string")
        .fillna("")
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("ascii")
        .str.upper()
    )
    dataframe = dataframe.sort_values("_orden_club").drop(columns="_orden_club")

    # Se genera el HTML sin sangrías ni saltos iniciales. Markdown interpretaba
    # las tarjetas siguientes como bloques de código y mostraba las etiquetas.
    club_cards = []
    for _, club in dataframe.iterrows():
        club_cards.append(
            '<div class="club-card">'
            '<div style="font-size:2rem">🛡️</div>'
            f'<h3 style="color:#092a55;margin:7px 0">{safe_html(club.get("club"))}</h3>'
            f'<div>📍 {safe_html(club.get("estadio"), "Por definir")}</div>'
            f'<div>📅 Fundación: {safe_html(club.get("fundacion"), "Por definir")}</div>'
            f'<div>👤 Presidente: {safe_html(club.get("presidente"), "Por definir")}</div>'
            '</div>'
        )

    clubs_html = '<div class="clubs-single-row">' + ''.join(club_cards) + '</div>'
    st.markdown(clubs_html, unsafe_allow_html=True)

elif menu == "Jugadores":
    st.markdown('<div class="section-title">Nómina de jugadores</div>', unsafe_allow_html=True)
    dataframe = players[players["serie"] == serie].copy()

    club_options = ["Todos"] + sorted(dataframe["club"].dropna().astype(str).unique().tolist())
    selected_club = st.selectbox("Club", club_options)
    if selected_club != "Todos":
        dataframe = dataframe[dataframe["club"] == selected_club]

    if dataframe.empty:
        st.info("No existen jugadores registrados para esta serie.")
    else:
        dataframe["jugador"] = (
            dataframe.get("nombres", "").fillna("").astype(str).str.strip()
            + " "
            + dataframe.get("apellido_paterno", "").fillna("").astype(str).str.strip()
            + " "
            + dataframe.get("apellido_materno", "").fillna("").astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()

        display_columns = [
            column for column in [
                "id_jugador", "jugador", "club", "numero_camiseta", "posicion",
                "fecha_nacimiento", "edad", "estado_jugador",
            ] if column in dataframe.columns
        ]
        st.dataframe(dataframe[display_columns], hide_index=True, use_container_width=True)

elif menu == "Goleadores":
    st.markdown('<div class="section-title">Tabla de goleadores</div>', unsafe_allow_html=True)
    dataframe = goals[goals["serie"] == serie].copy()

    if dataframe.empty:
        st.info("No existen goles individuales registrados para esta serie.")
    else:
        dataframe["jugador"] = dataframe["jugador"].apply(clean_text)
        dataframe["club"] = dataframe["club"].apply(clean_text)
        summary = (
            dataframe.groupby(["jugador", "club"], dropna=False)
            .size()
            .reset_index(name="Goles")
            .sort_values(["Goles", "jugador"], ascending=[False, True])
            .reset_index(drop=True)
        )
        summary.insert(0, "Pos.", range(1, len(summary) + 1))
        st.dataframe(summary, hide_index=True, use_container_width=True)

st.markdown("---")
st.caption("© 2026 Asociación de Fútbol Nacimiento")
