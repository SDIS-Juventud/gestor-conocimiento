# -*- coding: utf-8 -*-
"""
Genera el archivo gestion_conocimiento_jco_2025.html
para el gestor de conocimiento de Jóvenes con Oportunidades.

Cada sección del HTML se almacena como variable de Python
para facilitar la edición de contenido sin tocar HTML crudo.
"""

import os
import sys
from urllib.parse import quote
import pandas as pd

# CSS y datos compartidos con los demás generadores
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _comun.estilos import css_para
from _comun.aliados import seccion_jco as seccion_aliados_jco
from _comun.diagramas_flujo import svg_diagrama_jco

# ============================================================
# Inventario de documentación oficial de JCO.
# Los metadatos y URLs viven en `enlaces/enlaces.xlsx`, hoja
# `jco_documentacion` (columnas: CATEGORIA, LABEL, FECHA, URL).
# Para agregar/quitar/renombrar documentos basta editar el Excel.
# Las categorías se definen aquí porque incluyen estilos visuales
# (ícono Lucide y color) que no es necesario versionar en Excel.
# ============================================================

CATEGORIAS_DOCS = {
    "manual-jco":      {"label": "Manual JCO",        "icon": "book-open",      "color": "#663a93"},
    "instructivo":     {"label": "Instructivo pago",  "icon": "clipboard-list", "color": "#1e9da3"},
    "convenio":        {"label": "Convenio 1285-2025","icon": "handshake",      "color": "#f58b53"},
    "manual-parceros": {"label": "Manual Parceros",   "icon": "archive",        "color": "#1e7895"},
    "portafolios":     {"label": "Portafolios SDIS",  "icon": "library",        "color": "#f4676e"},
    "normativo":       {"label": "Marco normativo",   "icon": "scale",          "color": "#1eaf76"},
}

_BASE_JCO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_XLSX_JCO_DOCS = os.path.join(_BASE_JCO, "enlaces", "enlaces.xlsx")

def _cargar_documentos_desde_excel():
    """Lee la hoja jco_documentacion de enlaces.xlsx y devuelve una lista de
    tuplas (categoria, label, fecha, url) en el orden del Excel. Si la hoja
    no existe o falla la lectura, devuelve [] y avisa por stdout."""
    if not os.path.exists(_XLSX_JCO_DOCS):
        print(f"enlaces/enlaces.xlsx no encontrado; sección Documentación quedará vacía")
        return []
    try:
        df = pd.read_excel(_XLSX_JCO_DOCS, sheet_name="jco_documentacion")
    except Exception as e:
        print(f"No se pudo leer hoja 'jco_documentacion': {e}")
        return []
    docs = []
    for _, fila in df.iterrows():
        cat = str(fila.get("CATEGORIA", "")).strip()
        label = str(fila.get("LABEL", "")).strip()
        fecha = str(fila.get("FECHA", "")).strip()
        url = str(fila.get("URL", "")).strip()
        if not cat or not label or not url:
            continue
        if cat not in CATEGORIAS_DOCS:
            print(f"  ! Categoría desconocida '{cat}' en fila '{label}' (se ignora)")
            continue
        docs.append((cat, label, fecha, url))
    print(f"Documentación JCO desde enlaces.xlsx: {len(docs)} documentos")
    return docs

DOCUMENTOS_JCO = _cargar_documentos_desde_excel()


def _generar_bloques_documentos():
    """Devuelve el HTML con un bloque por cada categoría: subtítulo de color +
    grid de cards. Cada card es un <a> que abre el PDF en nueva pestaña."""
    bloques = []
    for cat, meta in CATEGORIAS_DOCS.items():
        docs_cat = [d for d in DOCUMENTOS_JCO if d[0] == cat]
        if not docs_cat:
            continue
        cards = []
        for _, label, fecha, ruta in docs_cat:
            # URLs http/https ya vienen codificadas (ej. SharePoint); las locales sí necesitan quote().
            if ruta.startswith(("http://", "https://")):
                url = ruta
            else:
                url = quote(ruta, safe="/")
            cards.append(
                f'                            <a class="docs-card" href="{url}" target="_blank">\n'
                f'                                <div class="docs-icono" style="color:{meta["color"]};"><i data-lucide="{meta["icon"]}"></i></div>\n'
                f'                                <div class="docs-info">\n'
                f'                                    <p class="docs-meta">{fecha}</p>\n'
                f'                                    <p class="docs-name">{label}</p>\n'
                f'                                    <span class="docs-link">Abrir PDF &rarr;</span>\n'
                f'                                </div>\n'
                f'                            </a>'
            )
        bloque = (
            f'                    <h3 class="docs-grupo-titulo" style="color:{meta["color"]};">{meta["label"]} <span class="docs-grupo-conteo">({len(docs_cat)})</span></h3>\n'
            f'                    <div class="docs-grid">\n'
            + "\n".join(cards)
            + '\n                    </div>'
        )
        bloques.append(bloque)
    return "\n\n".join(bloques)
# ============================================================
# 1. CSS — estilos del gestor
# ============================================================

# Las reglas para .modulo-acordeon viven en _comun/estilos.py porque
# ahora se comparten con el generador de Forjar.
#
# EXTRAS_CSS_JCO replica los componentes nuevos del rediseño 2026-05 ya
# aplicados a Forjar: la infografía "A tener en cuenta" con manos halftone
# (.atc-*), el flujo de gestión con íconos Lucide en círculo accent y
# conector punteado (.flujo-*), las cards de Equipo (.equipo-card-*) y el
# override de subtítulos en Antonio Bold mayúsculas (.card-subtitle).

EXTRAS_CSS_JCO = """\
/* Infograf&iacute;a "A tener en cuenta": 5 fichas crema con manos halftone del branding.
   Sin bandas de color en cabecera (eso se siente AI). El acento del punto vive
   en el subt&iacute;tulo en Antonio Bold del color de la paleta oficial. */
.atc-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 22px;
    margin-top: 18px;
}
.atc-card {
    background: #f5efd2;
    border-radius: 14px;
    padding: 28px 24px 26px;
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    transition: transform 0.18s ease, box-shadow 0.18s ease;
}
.atc-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 10px 24px rgba(0,0,0,0.08);
}
.atc-mano {
    margin: 0 0 16px;
    height: 110px;
    display: flex;
    align-items: center;
    justify-content: center;
}
.atc-mano img { height: 100%; object-fit: contain; display: block; }
.atc-titulo {
    font-family: 'Antonio', 'Anton', 'Segoe UI', sans-serif;
    font-weight: 700;
    font-size: 1.05rem;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    margin-bottom: 12px;
    line-height: 1.2;
}
.atc-titulo-1 { color: #f4676e; }
.atc-titulo-2 { color: #1eaf76; }
.atc-titulo-3 { color: #663a93; }
.atc-titulo-4 { color: #f58b53; }
.atc-titulo-5 { color: #1e9da3; }
.atc-titulo-6 { color: #1e7895; }
.atc-texto {
    font-family: 'Figtree', 'Segoe UI', sans-serif;
    font-weight: 500;
    font-size: 0.83rem;
    color: #3a3a3a;
    line-height: 1.6;
}
.atc-texto strong { font-weight: 700; color: #2f3e3c; }
@media (max-width: 900px) {
    .atc-grid { grid-template-columns: 1fr; }
}
@media (min-width: 901px) and (max-width: 1100px) {
    .atc-grid { grid-template-columns: repeat(2, 1fr); }
}

/* Flujo de gesti&oacute;n de la informaci&oacute;n: pasos en una columna con &iacute;cono
   redondo a la izquierda y contenido a la derecha. Conector vertical punteado
   sutil entre &iacute;conos para se&ntilde;alar el flujo. */
.flujo-pasos { display: flex; flex-direction: column; gap: 28px; margin: 22px 0 8px; }
.flujo-paso {
    display: grid;
    grid-template-columns: 64px 1fr;
    gap: 22px;
    align-items: flex-start;
    position: relative;
}
.flujo-paso:not(:last-child)::before {
    content: '';
    position: absolute;
    left: 31px;
    top: 64px;
    bottom: -28px;
    width: 0;
    border-left: 2px dashed var(--accent-border);
}
.flujo-icono {
    width: 56px;
    height: 56px;
    border-radius: 50%;
    background: var(--accent-bg);
    color: var(--accent);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    z-index: 1;
}
.flujo-icono svg { width: 26px; height: 26px; stroke-width: 1.8; }
.flujo-orden {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 0.72rem;
    color: #888;
    margin: 0 0 2px;
}
.flujo-titulo {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    font-size: 1.15rem;
    color: var(--accent);
    margin: 0 0 4px;
}
.flujo-responsable {
    font-size: 0.75rem;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin: 0 0 12px;
    font-weight: 600;
}
.flujo-texto {
    font-size: 0.9rem;
    color: #3a3a3a;
    line-height: 1.7;
    margin: 0 0 8px;
}
.flujo-texto:last-child { margin-bottom: 0; }
.flujo-texto strong { color: #2f3e3c; font-weight: 700; }
.flujo-lista { font-size: 0.9rem; color: #3a3a3a; line-height: 1.75; margin: 4px 0 8px 18px; }
.flujo-lista li { margin-bottom: 4px; }
.flujo-lista strong { color: #2f3e3c; font-weight: 700; }
@media (max-width: 720px) {
    .flujo-paso { grid-template-columns: 48px 1fr; gap: 14px; }
    .flujo-paso:not(:last-child)::before { left: 23px; top: 48px; }
    .flujo-icono { width: 40px; height: 40px; }
    .flujo-icono svg { width: 20px; height: 20px; }
}

/* Organigrama JCO: tree horizontal jer&aacute;rquico, replica el formato que
   comparti&oacute; el equipo. Card de l&iacute;der arriba, debajo seis ramas (una por
   l&iacute;nea de jefatura) en colores de la paleta extendida, y subniveles
   conectados con l&iacute;neas verticales y horizontales en CSS puro. Permite
   scroll horizontal en pantallas peque&ntilde;as. En m&oacute;vil colapsa a vertical
   con conectores punteados a la izquierda. */
.org-tree { overflow-x: auto; padding: 12px 4px 24px; }
.org-tree-inner { display: inline-block; min-width: 100%; }
.org-nodo { display: flex; flex-direction: column; align-items: center; position: relative; }
.org-hijos {
    display: flex;
    justify-content: center;
    gap: 18px;
    margin-top: 26px;
    padding-top: 14px;
    position: relative;
}
/* Conector vertical desde la card del padre hasta la barra horizontal */
.org-hijos::after {
    content: '';
    position: absolute;
    top: -12px;
    left: 50%;
    width: 0;
    height: 12px;
    border-left: 1.5px solid #c4cec9;
}
/* Barra horizontal entre hermanos (se acorta a 0 si solo hay un hijo) */
.org-hijos::before {
    content: '';
    position: absolute;
    top: 0;
    left: 50%;
    right: 50%;
    border-top: 1.5px solid #c4cec9;
}
.org-hijos.org-hijos-multi::before {
    left: calc(var(--first-half, 50%) * 0);
    right: 0;
}
.org-hijos.org-hijos-multi { /* posiciones de la barra: del primer al &uacute;ltimo */ }
.org-hijos.org-hijos-multi > .org-nodo:first-child::after { content: ''; position: absolute; top: -14px; left: 50%; right: 0; border-top: 1.5px solid #c4cec9; }
.org-hijos.org-hijos-multi > .org-nodo:last-child::after { content: ''; position: absolute; top: -14px; left: 0; right: 50%; border-top: 1.5px solid #c4cec9; }
.org-hijos.org-hijos-multi > .org-nodo:not(:first-child):not(:last-child)::after { content: ''; position: absolute; top: -14px; left: 0; right: 0; border-top: 1.5px solid #c4cec9; }
.org-hijos.org-hijos-multi::before { display: none; }
/* Conector vertical desde la barra hasta la card de cada hijo */
.org-hijos > .org-nodo::before {
    content: '';
    position: absolute;
    top: -14px;
    left: 50%;
    width: 0;
    height: 14px;
    border-left: 1.5px solid #c4cec9;
    z-index: 1;
}
/* Card de cada nodo */
.org-card {
    background: #2F3E3C;
    color: #F8F4E1;
    border-radius: 6px;
    padding: 10px 14px;
    text-align: center;
    min-width: 150px;
    max-width: 210px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    line-height: 1.3;
    position: relative;
    z-index: 2;
}
.org-card .org-rol {
    display: block;
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-size: 0.68rem;
    opacity: 0.9;
    margin-bottom: 3px;
}
.org-card .org-persona {
    display: block;
    font-family: 'Figtree', 'Segoe UI', sans-serif;
    font-weight: 700;
    font-size: 0.86rem;
}
.org-card .org-persona em {
    display: block;
    font-style: normal;
    font-weight: 500;
    font-size: 0.72rem;
    opacity: 0.8;
    margin-top: 2px;
}
.org-card-lider { background: #2F3E3C; min-width: 220px; padding: 14px 22px; }
.org-card-lider .org-persona { font-size: 1rem; }
/* Colores por NIVEL jer&aacute;rquico (no por rama): primera l&iacute;nea toda en verde,
   segunda en morado, tercera (leaves) en crema. El verde se eligi&oacute; en lugar
   del rosa coral para evitar la connotaci&oacute;n negativa del rojo en la primera l&iacute;nea. */
.org-card-1, .org-card-2, .org-card-3,
.org-card-4, .org-card-5, .org-card-6 { background: #1eaf76; }
.org-card-sub { background: #663a93; color: #F8F4E1; }
.org-card-leaf { background: #f5efd2; color: #2F3E3C; min-width: 140px; font-size: 0.78rem; padding: 8px 12px; }
.org-card-leaf .org-rol { opacity: 1; color: #666; }
.org-card-leaf .org-persona { font-weight: 600; font-size: 0.8rem; }
/* En m&oacute;vil: tree colapsa a layout vertical con barras a la izquierda */
@media (max-width: 760px) {
    .org-tree { overflow-x: visible; padding: 12px 0 20px; }
    .org-tree-inner { display: block; }
    .org-nodo { align-items: stretch; }
    .org-hijos {
        flex-direction: column;
        gap: 10px;
        margin: 14px 0 0 22px;
        padding: 0 0 0 18px;
        border-left: 1.5px dashed #c4cec9;
    }
    .org-hijos::before, .org-hijos::after { display: none; }
    .org-hijos > .org-nodo::before {
        content: '';
        position: absolute;
        top: 16px;
        left: -18px;
        width: 18px;
        height: 0;
        border-top: 1.5px dashed #c4cec9;
        border-left: none;
    }
    .org-hijos.org-hijos-multi > .org-nodo::after { display: none; }
    .org-card { max-width: 100%; text-align: left; }
}

/* Secci&oacute;n Documentaci&oacute;n: grids por categor&iacute;a con cards linkeadas a los PDFs.
   Cada categor&iacute;a abre con su subt&iacute;tulo de color (Antonio Bold may&uacute;sculas) y
   muestra los documentos como cards crema neutras con &iacute;cono Lucide del color
   de la categor&iacute;a. Pensado para listar y referenciar trazabilidad documental. */
.docs-grupo-titulo {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    font-size: 1.05rem;
    margin: 28px 0 12px;
}
.docs-grupo-conteo { color: #888; font-weight: 500; font-size: 0.85rem; letter-spacing: 0; margin-left: 4px; }
.docs-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 12px; margin-bottom: 8px; }
.docs-card {
    display: flex;
    align-items: flex-start;
    gap: 14px;
    background: #ffffff;
    border: 1.5px solid #e5e0d3;
    border-radius: 12px;
    padding: 18px 20px;
    text-decoration: none;
    color: #2F3E3C;
    transition: transform 0.18s ease, box-shadow 0.18s ease;
}
.docs-card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
.docs-icono { flex-shrink: 0; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; }
.docs-icono svg { width: 28px; height: 28px; stroke-width: 1.7; }
.docs-info { flex: 1; min-width: 0; }
.docs-meta {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    font-size: 0.7rem;
    color: #888;
    margin: 0 0 3px;
}
.docs-name { font-family: 'Figtree', 'Segoe UI', sans-serif; font-weight: 600; font-size: 0.92rem; line-height: 1.35; color: #2F3E3C; margin: 0 0 6px; }
.docs-link { font-family: 'Figtree', 'Segoe UI', sans-serif; font-size: 0.78rem; font-weight: 600; color: var(--accent); }

/* Override JCO: TODOS los subt&iacute;tulos de tarjeta van en Antonio Bold
   may&uacute;sculas, consistente con .lt-titulo, .atc-titulo, .flujo-titulo
   y dem&aacute;s subt&iacute;tulos especializados de la secci&oacute;n. */
.card-subtitle {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    font-size: 1.05rem;
    opacity: 1;
}

/* Tabla comparativa para "A tener en cuenta" de JCO: 3 programas (RETO,
   Parceros, JcO) que son etapas evolutivas del mismo programa. JcO es el
   "actual" y va en card s&oacute;lida morado oficial; los dos hist&oacute;ricos quedan
   en cards neutras con badge de color. Mismo lenguaje que la tabla de
   Casas pero con 4 columnas (1 caracter&iacute;stica + 3 programas). */
.comp-tabla { position: relative; background: #f5efd2; border-radius: 14px; padding: 24px 24px 18px; margin: 32px 0 28px; box-shadow: 0 2px 10px rgba(0,0,0,0.04); overflow: visible; }
.comp-mano-decor { position: absolute; top: -38px; left: -32px; width: 120px; height: auto; transform: rotate(-14deg); pointer-events: none; z-index: 2; }
.comp-tabla-grid { display: grid; grid-template-columns: 1.1fr 1.4fr 1.4fr; gap: 14px 22px; align-items: center; }
.comp-header { display: flex; align-items: center; justify-content: center; padding: 10px 16px; border-radius: 22px; font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.8rem; line-height: 1.2; text-align: center; color: #F8F4E1; }
.comp-header-1 { background: #2F3E3C; }
.comp-header-2 { background: #1e7895; }
.comp-header-3 { background: #663a93; }
.comp-row-sep { grid-column: 1 / -1; border-top: 1px solid rgba(47, 62, 60, 0.12); margin: 6px 0; }
.comp-caracteristica { font-family: 'Figtree', 'Segoe UI', sans-serif; font-weight: 700; color: #2F3E3C; font-size: 0.95rem; line-height: 1.3; padding: 10px 4px; }
.comp-celda { display: flex; align-items: flex-start; gap: 14px; padding: 10px 4px; font-family: 'Figtree', 'Segoe UI', sans-serif; font-size: 0.88rem; color: #2F3E3C; line-height: 1.5; }
.comp-icono { flex-shrink: 0; width: 52px; height: 52px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-top: 2px; }
.comp-icono svg { width: 26px; height: 26px; stroke-width: 1.7; }
.comp-parceros .comp-icono { background: rgba(30, 120, 149, 0.18); color: #1e7895; }
.comp-jco .comp-icono { background: rgba(102, 58, 147, 0.15); color: #663a93; }
.comp-celda-texto { display: flex; flex-direction: column; gap: 2px; }
.comp-celda-badge { display: none; }
@media (max-width: 760px) {
    .comp-tabla { padding: 30px 16px 18px; margin-top: 42px; }
    .comp-mano-decor { width: 88px; top: -28px; left: -18px; }
    .comp-tabla-grid { grid-template-columns: 1fr; gap: 0; }
    .comp-header { display: none; }
    .comp-row-sep { display: none; }
    .comp-caracteristica { padding: 16px 0 8px; font-size: 0.78rem; color: #2F3E3C; text-transform: uppercase; letter-spacing: 0.06em; font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; border-top: 1px solid rgba(47, 62, 60, 0.18); }
    .comp-caracteristica:first-of-type { border-top: none; padding-top: 4px; }
    .comp-celda { padding: 12px 14px; margin-bottom: 8px; border-radius: 10px; }
    .comp-parceros { background: #ffffff; color: #2F3E3C; }
    .comp-jco { background: #663a93; color: #F8F4E1; }
    .comp-jco .comp-icono { background: rgba(248, 244, 225, 0.15); color: #F8F4E1; }
    .comp-celda-badge { display: inline-block; font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.62rem; padding: 2px 8px; border-radius: 10px; color: #F8F4E1; margin-bottom: 4px; }
    .comp-parceros .comp-celda-badge { background: #1e7895; }
    .comp-jco .comp-celda-badge { background: rgba(248, 244, 225, 0.22); color: #F8F4E1; }
    .comp-icono { width: 40px; height: 40px; }
    .comp-icono svg { width: 20px; height: 20px; }
}
"""

CSS = css_para("jco", extras=EXTRAS_CSS_JCO)

# ============================================================
# 2. Encabezado (header)
# ============================================================

HEADER = """\
    <header class="header">
        <div>
            <h1>Gestor de conocimiento - J&oacute;venes con Oportunidades</h1>
            <div class="subtitle">Subdirecci&oacute;n para la Juventud | SDIS</div>
        </div>
        <div class="header-btns">
            <a class="home-btn" href="index.html" title="Todos los servicios">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#F8F4E1" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </a>
            <div class="home-btn" onclick="showContent('welcome')" title="Inicio J&oacute;venes con Oportunidades">
                <img src="imagenes/servicios/jovenes-con-oportunidades.png" alt="J&oacute;venes con Oportunidades" style="height:32px; border-radius:16px; object-fit:contain; vertical-align:middle;">
            </div>
        </div>
    </header>"""

# ============================================================
# 3. Barra lateral (sidebar)
# ============================================================

SIDEBAR = """\
        <nav class="sidebar">
            <div class="sidebar-title" onclick="showContent('welcome')" style="cursor:pointer;"><span>Inicio</span></div>

            <div class="sidebar-section">
                <div class="sidebar-title active" onclick="toggleSection(this)">
                    <span>Contexto</span><span class="arrow">&#9654;</span>
                </div>
                <div class="sidebar-items show">
                    <div class="sidebar-item" onclick="showContent('linea_tiempo')">L&iacute;nea de tiempo</div>
                    <div class="sidebar-item" onclick="showContent('a_tener_en_cuenta')">A tener en cuenta</div>
                    <div class="sidebar-item" onclick="showContent('equipo')">Equipo</div>
                    <div class="sidebar-item" onclick="showContent('pilares')">Pilares del programa</div>
                    <div class="sidebar-item" onclick="showContent('modulos_proyecto_vida')">M&oacute;dulos de Proyecto de Vida</div>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="showContent('documentacion')" style="cursor:pointer;">
                    <span>Documentaci&oacute;n</span>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="toggleSection(this)">
                    <span>Gesti&oacute;n de datos</span><span class="arrow">&#9654;</span>
                </div>
                <div class="sidebar-items">
                    <div class="sidebar-item" onclick="showContent('flujo_datos')">Flujo de gesti&oacute;n de la informaci&oacute;n</div>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="showContent('aliados')" style="cursor:pointer;">
                    <span>Aliados</span>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="showContent('estadisticas')" style="cursor:pointer;">
                    <span>Estad&iacute;sticas</span>
                </div>
            </div>
        </nav>"""

# ============================================================
# 4. Secciones de contenido
#    Cada sección es una variable independiente para editar
#    sin afectar las demás.
# ============================================================

# --- Bienvenida ---
SECCION_WELCOME = """\
            <div class="content-section active" id="welcome">
                <div class="welcome-section">
                    <div style="font-family:'Anton','Figtree',sans-serif; font-weight:400; font-size:1.9rem; line-height:1.05; letter-spacing:1px; text-transform:uppercase; background:#2d2a28; color:#f4f5de; padding:14px 24px 11px; margin:0 auto 28px; display:block; width:fit-content; max-width:100%; text-align:center;">J&oacute;venes con Oportunidades</div>
                    <p>J&oacute;venes con Oportunidades (SJO) es un servicio de la Secretar&iacute;a Distrital de Integraci&oacute;n Social, operado por la Subdirecci&oacute;n para la Juventud, en el que se desarrolla una estrategia intersectorial junto a la Secretar&iacute;a de Educaci&oacute;n del Distrito, la Secretar&iacute;a Distrital de Desarrollo Econ&oacute;mico y la Agencia Distrital para la Educaci&oacute;n Superior, la Ciencia y la Tecnolog&iacute;a &mdash; Atenea, orientada a la inclusi&oacute;n social y productiva de j&oacute;venes entre 14 y 28 a&ntilde;os en situaci&oacute;n de pobreza extrema, moderada o vulnerabilidad por inseguridad alimentaria. J&oacute;venes con Oportunidades articula una oferta con acompa&ntilde;amiento psicosocial, acceso a formaci&oacute;n, intermediaci&oacute;n laboral y transferencias monetarias condicionadas en un proceso continuo.</p>
                    <div style="margin:30px auto 0; max-width:450px;">
                        <img src="imagenes/Jovenes con oportunidades.png" alt="J&oacute;venes con Oportunidades" style="width:100%; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.1);">
                    </div>
                </div>
            </div>"""

# --- A tener en cuenta ---
# Sección comparativa: el programa nació como Estrategia RETO (2020-21), pasó
# a llamarse Parceros por Bogotá (2021-23) y actualmente es Jóvenes con
# Oportunidades (2024-presente). Son tres etapas evolutivas del mismo programa.
# La tabla comparativa resume 5 características clave; los 5 párrafos debajo
# profundizan aspectos transversales (financiación, condicionalidad, etc.).
SECCION_A_TENER_EN_CUENTA = """\
            <div class="content-section" id="a_tener_en_cuenta">
                <div class="card">
                    <h2 class="card-title">A tener en cuenta</h2>
                    <p style="line-height:1.7; margin-bottom:8px;">El programa actual <strong>J&oacute;venes con Oportunidades</strong> (2024&ndash;presente) sucede a <strong>Parceros por Bogot&aacute;</strong> (2021&ndash;2023) en el componente de transferencias monetarias condicionadas para j&oacute;venes vulnerables. La <strong>Estrategia RETO</strong> fue la &ldquo;sombrilla m&aacute;s amplia&rdquo; o estrategia interinstitucional de la Secretar&iacute;a Distrital de Integraci&oacute;n Social de donde naci&oacute; Parceros.</p>
                    <p style="line-height:1.7; margin-bottom:8px;">La tabla compara las dos etapas del componente de transferencias y formaci&oacute;n para enganche laboral; m&aacute;s abajo se profundiza en aspectos transversales (financiaci&oacute;n, condicionalidad, intermediaci&oacute;n).</p>

                    <div class="comp-tabla">
                        <img class="comp-mano-decor" src="imagenes/manos/3.png" alt="">
                        <div class="comp-tabla-grid">
                            <div class="comp-header comp-header-1">Caracter&iacute;stica</div>
                            <div class="comp-header comp-header-2">Parceros por Bogot&aacute;<br>2021&ndash;2023</div>
                            <div class="comp-header comp-header-3">J&oacute;venes con Oportunidades<br>2024&ndash;Presente</div>

                            <div class="comp-row-sep"></div>
                            <div class="comp-caracteristica">Enfoque principal</div>
                            <div class="comp-celda comp-parceros"><div class="comp-icono"><i data-lucide="shield"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Parceros</span><span>Modelo de mitigaci&oacute;n de riesgos a trav&eacute;s del servicio comunitario temporal.</span></div></div>
                            <div class="comp-celda comp-jco"><div class="comp-icono"><i data-lucide="briefcase"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">JcO</span><span>Modelo de inclusi&oacute;n productiva a trav&eacute;s de la educaci&oacute;n y la conexi&oacute;n directa con el mercado laboral.</span></div></div>

                            <div class="comp-row-sep"></div>
                            <div class="comp-caracteristica">Duraci&oacute;n de la oferta</div>
                            <div class="comp-celda comp-parceros"><div class="comp-icono"><i data-lucide="calendar-clock"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Parceros</span><span>6 meses: 2 de capacitaci&oacute;n y 4 de servicio a la ciudad o de formaci&oacute;n educativa e inclusi&oacute;n laboral.</span></div></div>
                            <div class="comp-celda comp-jco"><div class="comp-icono"><i data-lucide="calendar-range"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">JcO</span><span>Variable seg&uacute;n ruta: 40 a 160 horas en cursos cortos; entre 1 y 2 ciclos (semestres) en educaci&oacute;n flexible y posmedia; entre 2 y 10 semestres en educaci&oacute;n posmedia de ciclo largo.</span></div></div>

                            <div class="comp-row-sep"></div>
                            <div class="comp-caracteristica">Componentes clave</div>
                            <div class="comp-celda comp-parceros"><div class="comp-icono"><i data-lucide="hammer"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Parceros</span><span>i) Componente pedag&oacute;gico y de acompa&ntilde;amiento psicosocial; ii) componente pr&aacute;ctico y de incidencia comunitaria; iii) componente de formaci&oacute;n educativa e inclusi&oacute;n laboral; iv) componente recreativo y de manejo del tiempo.</span></div></div>
                            <div class="comp-celda comp-jco"><div class="comp-icono"><i data-lucide="layers"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">JcO</span><span>i) Acompa&ntilde;amiento psicosocial; ii) acceso a rutas de formaci&oacute;n; iii) intermediaci&oacute;n laboral; iv) transferencias monetarias condicionadas.</span></div></div>

                            <div class="comp-row-sep"></div>
                            <div class="comp-caracteristica">Transferencia monetaria</div>
                            <div class="comp-celda comp-parceros"><div class="comp-icono"><i data-lucide="banknote"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Parceros</span><span>$500.000 COP mensuales durante 6 meses; m&aacute;ximo acumulado <strong>$3.000.000</strong>.</span></div></div>
                            <div class="comp-celda comp-jco"><div class="comp-icono"><i data-lucide="wallet"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">JcO</span><span>Variable seg&uacute;n ruta de formaci&oacute;n, por hitos psicosociales, de formaci&oacute;n y de intermediaci&oacute;n laboral. M&aacute;ximo acumulado: <strong>$1.200.000</strong> (cursos cortos), <strong>$3.200.000</strong> (bachillerato), <strong>$4.300.000</strong> (universidad) o <strong>3&nbsp;SMMLV + $300.000</strong> (educaci&oacute;n y formaci&oacute;n para el trabajo).</span></div></div>

                            <div class="comp-row-sep"></div>
                            <div class="comp-caracteristica">Poblaci&oacute;n objetivo</div>
                            <div class="comp-celda comp-parceros"><div class="comp-icono"><i data-lucide="user-x"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Parceros</span><span>J&oacute;venes entre <strong>18 y 28 a&ntilde;os</strong> altamente vulnerables.</span></div></div>
                            <div class="comp-celda comp-jco"><div class="comp-icono"><i data-lucide="user-check"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">JcO</span><span>J&oacute;venes entre <strong>14 y 28 a&ntilde;os</strong> en pobreza extrema, moderada o vulnerabilidad (<strong>Sisb&eacute;n A, B o hasta C09</strong>; pagadiarios), residentes en Bogot&aacute;.</span></div></div>
                        </div>
                    </div>

                    <h3 class="card-subtitle">Cambio de paradigma en la condicionalidad (componentes)</h3>
                    <p style="line-height:1.7;">En Parceros por Bogot&aacute;, la transferencia monetaria estaba condicionada a la participaci&oacute;n en actividades pedag&oacute;gicas y labores pr&aacute;cticas de servicio a la ciudad, tales como proyectos de embellecimiento, limpieza, huertas urbanas y &ldquo;Nuestras Zonas Seguras&rdquo;. Adem&aacute;s, los participantes se formaban y certificaban como Agentes Comunitarios de Prevenci&oacute;n en temas de salud mental, violencias y consumo de sustancias. Con J&oacute;venes con Oportunidades, la condicionalidad depende directamente del cumplimiento de actividades en tres componentes clave: el acompa&ntilde;amiento psicosocial transversal, el avance en la ruta de formaci&oacute;n asignada, y el proceso de intermediaci&oacute;n laboral.</p>

                    <h3 class="card-subtitle">Intermediaci&oacute;n laboral y conexi&oacute;n con el empleo (componentes)</h3>
                    <p style="line-height:1.7;">En Parceros por Bogot&aacute; exist&iacute;a un componente de <strong>formaci&oacute;n educativa e inclusi&oacute;n laboral</strong> que se establec&iacute;a <strong>desde el inicio de la atenci&oacute;n</strong>, ofertando oportunidades a los participantes y permitiendo el cumplimiento de actividades condicionadas. En el nuevo modelo de J&oacute;venes con Oportunidades, la intermediaci&oacute;n se consolida como un <strong>tercer componente de la ruta</strong>, una etapa formal y estructural. El objetivo es garantizar la integralidad y que el o la joven &ldquo;salga por el &uacute;ltimo eslab&oacute;n&rdquo; con el resultado esperado. Por ello, al finalizar su formaci&oacute;n, la Secretar&iacute;a Distrital de Desarrollo Econ&oacute;mico asume el liderazgo directo de la intermediaci&oacute;n laboral, gestionando el registro en el Servicio P&uacute;blico de Empleo y en las plataformas de la estrategia &ldquo;Talento Capital&rdquo;.</p>

                    <h3 class="card-subtitle">Los montos var&iacute;an seg&uacute;n la ruta de formaci&oacute;n (transferencia)</h3>
                    <p style="line-height:1.7;">Parceros entregaba una cuota fija de <strong>$500.000 mensuales durante 6 meses</strong>. En J&oacute;venes con Oportunidades el apoyo econ&oacute;mico depende de la ruta elegida: hasta <strong>$1.200.000 en cursos cortos</strong>, hasta <strong>$3.200.000</strong> en la ruta de educaci&oacute;n flexible para terminar el bachillerato (j&oacute;venes y adultos), y topes m&aacute;ximos que van desde <strong>$2.700.000</strong> (para nivel t&eacute;cnico y tecn&oacute;logo) hasta <strong>$4.300.000</strong> en educaci&oacute;n universitaria de ciclo largo.</p>
                    <p style="line-height:1.7;">En la ruta de <strong>EFT</strong> en particular, los recursos que pone la SDIS pueden llegar hasta <strong>3 salarios m&iacute;nimos</strong>, m&aacute;s <strong>3 pagos de $400.000</strong> y <strong>$300.000 por intermediaci&oacute;n</strong>.</p>

                    <h3 class="card-subtitle">Focalizaci&oacute;n y requisitos de ingreso (poblaci&oacute;n)</h3>
                    <p style="line-height:1.7;">Los criterios de selecci&oacute;n var&iacute;an entre ambos modelos. Parceros por Bogot&aacute; se enfocaba en j&oacute;venes <strong>que no estudian ni trabajan</strong> de <strong>18 a 28 a&ntilde;os</strong>, identificados directamente en barrios vulnerables mediante una encuesta de 23 preguntas que calculaba su &Iacute;ndice de Vulnerabilidad Juvenil. Por su parte, el nuevo programa J&oacute;venes con Oportunidades atiende a j&oacute;venes de <strong>14 a 28 a&ntilde;os</strong> en situaci&oacute;n de pobreza extrema, moderada o vulnerabilidad por inseguridad alimentaria. Su focalizaci&oacute;n es distinta, ya que exige que los beneficiarios residan en Bogot&aacute; y est&eacute;n registrados oficialmente en el Sisb&eacute;n dentro de las categor&iacute;as <strong>A, B o hasta C09</strong>. Un cambio importante respecto a Parceros es que <strong>en JcO la actividad principal de los j&oacute;venes no es un requisito excluyente</strong>: pueden estar estudiando o trabajando y aun as&iacute; ser beneficiarios, siempre que cumplan los criterios de vulnerabilidad econ&oacute;mica.</p>

                    <h3 class="card-subtitle">Financiaci&oacute;n de los programas</h3>
                    <p style="line-height:1.7;">Una caracter&iacute;stica clave de Parceros por Bogot&aacute; fue que se sustent&oacute;, en gran parte, gracias a la uni&oacute;n institucional con los <strong>Fondos de Desarrollo Local (FDL)</strong>. Esto permiti&oacute; una inversi&oacute;n territorializada, donde los alcaldes locales y ediles destinaban recursos para apoyar directamente a los j&oacute;venes de sus respectivas localidades. En J&oacute;venes con Oportunidades, los FDL tambi&eacute;n han venido aportando recursos: con un rol inicial <strong>desde 2025</strong> y una <strong>presencia muy fuerte en 2026</strong>. A estos aportes se suma la inversi&oacute;n distrital proyectada de <strong>$324.053 millones</strong>, que une los esfuerzos sectoriales de las Secretar&iacute;as de Integraci&oacute;n Social, Educaci&oacute;n, Desarrollo Econ&oacute;mico y la Agencia Atenea.</p>
                </div>
            </div>"""

# --- Línea de tiempo ---
# Componente chevron compartido con Forjar y Casas de Juventud (CSS en _comun/estilos.py).
# JCO usa 5 hitos, por eso --lt-cols: 5 y se usan los 5 colores intercalados de la
# paleta extendida SDIS Juventud (se descartan teal y azul petr&oacute;leo).
SECCION_LINEA_TIEMPO = """\
            <div class="content-section" id="linea_tiempo">
                <div class="card">
                    <h2 class="card-title">L&iacute;nea de tiempo</h2>
                    <p style="line-height:1.7; margin-bottom:24px;">Evoluci&oacute;n de las pol&iacute;ticas de inclusi&oacute;n social juvenil y transferencias monetarias en Bogot&aacute;: del modelo de emergencia <strong>Parceros por Bogot&aacute;</strong> hacia la estrategia integral <strong>J&oacute;venes con Oportunidades</strong>.</p>
                    <div class="linea-tiempo" style="--lt-cols: 5;">
                        <article class="lt-hito">
                            <div class="lt-chevron lt-c1"><i data-lucide="rocket"></i>2020</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i1"><i data-lucide="rocket"></i></div>
                                <div class="lt-titulo lt-t1">Estrategia RETO</div>
                                <div class="lt-texto">La SDIS lanza la Estrategia Retorno de las Oportunidades Juveniles &ndash; RETO, bajo el Proyecto de Inversi&oacute;n <strong>7740 &ldquo;Generaci&oacute;n J&oacute;venes con Derechos en Bogot&aacute;&rdquo;</strong>, para reducir el riesgo social en poblaci&oacute;n joven vulnerable.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c2"><i data-lucide="hand-coins"></i>2021</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i2"><i data-lucide="hand-coins"></i></div>
                                <div class="lt-titulo lt-t2">Parceros por Bogot&aacute;</div>
                                <div class="lt-texto">En el marco de las mesas de di&aacute;logo realizadas durante el estallido social y los efectos de la pandemia por el COVID-19, se crea el <strong>Servicio Social para la Seguridad Econ&oacute;mica para la Juventud</strong>, con la oferta de transferencias de <strong>$500.000 mensuales por seis meses</strong>, condicionadas a participaci&oacute;n en actividades pedag&oacute;gicas, servicio a la ciudad y acompa&ntilde;amiento psicosocial.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c3"><i data-lucide="flag"></i>2023</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i3"><i data-lucide="flag"></i></div>
                                <div class="lt-titulo lt-t3">Modificaci&oacute;n de Parceros</div>
                                <div class="lt-texto">El programa se modifica (septiembre) con <strong>m&aacute;s de 28.000 j&oacute;venes atendidos</strong>. <strong>4 de cada 10 egresados</strong> obtuvieron empleo formal, apoyo para emprendimiento o acceso a educaci&oacute;n superior.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c4"><i data-lucide="sparkles"></i>2024</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i4"><i data-lucide="sparkles"></i></div>
                                <div class="lt-titulo lt-t4">Lanzamiento JCO</div>
                                <div class="lt-texto">Se lanza <strong>J&oacute;venes con Oportunidades</strong> en el <strong>Proyecto de Inversi&oacute;n 7940</strong>. Mantiene acompa&ntilde;amiento psicosocial y transferencias condicionadas, y articula los componentes de acceso a rutas de formaci&oacute;n e intermediaci&oacute;n laboral, a trav&eacute;s de la articulaci&oacute;n intersectorial entre Integraci&oacute;n Social, Educaci&oacute;n, Desarrollo Econ&oacute;mico y Agencia Atenea.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c6"><i data-lucide="target"></i>2024&ndash;27</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i6"><i data-lucide="target"></i></div>
                                <div class="lt-titulo lt-t6">Meta cuatrienio</div>
                                <div class="lt-texto">El reto del cuatrienio es beneficiar a <strong>40.000 j&oacute;venes</strong>, con una inversi&oacute;n distrital proyectada de <strong>$324.053 millones</strong>.</div>
                            </div>
                        </article>
                    </div>
                </div>
            </div>"""

# --- Equipo ---
# Estructura visual replicando el formato que compartió Alejandro: bloque ancho
# de coordinación arriba, bloque grande de equipo territorial al medio, y dos
# bloques 50/50 abajo (metodológico/gestión documental y administrativo).
# Paleta JCO: fondo morado #663a93 + sombra verde #1eaf76 + títulos en verde
# claro #b8f0d4. Listas centradas sin recuadros internos.
SECCION_EQUIPO = """\
            <div class="content-section" id="equipo">
                <div class="card">
                    <h2 class="card-title">Equipo y gesti&oacute;n de J&oacute;venes con Oportunidades</h2>
                    <p style="line-height:1.7; margin-bottom:24px;">El servicio J&oacute;venes con Oportunidades es un esfuerzo articulado entre la <strong>Secretar&iacute;a Distrital de Integraci&oacute;n Social (SDIS)</strong>, la <strong>Secretar&iacute;a de Educaci&oacute;n</strong>, la <strong>Secretar&iacute;a de Desarrollo Econ&oacute;mico</strong> y la <strong>Agencia Atenea</strong>. La gesti&oacute;n operativa recae principalmente en la Subdirecci&oacute;n para la Juventud de la SDIS, cuyo equipo se organiza as&iacute;:</p>

                    <div style="background:#663a93; color:#fff; border-radius:10px; padding:18px 25px; text-align:center; margin-bottom:28px; box-shadow:7px 7px 0 #1eaf76;">
                        <div style="font-size:0.75rem; text-transform:uppercase; letter-spacing:1.5px; color:#b8f0d4; margin-bottom:4px; font-weight:700;">1. Coordinaci&oacute;n general</div>
                        <div style="font-size:1.15rem; font-weight:700; color:#fff;">Ana Catalina Su&aacute;rez</div>
                        <div style="font-size:0.78rem; color:rgba(255,255,255,0.75); margin-top:2px;">L&iacute;der del servicio</div>
                    </div>

                    <div style="background:#663a93; border-radius:14px; padding:26px 26px 22px; box-shadow:7px 7px 0 #1eaf76; color:#fff; margin-bottom:34px; text-align:center;">
                        <div style="font-size:0.75rem; text-transform:uppercase; letter-spacing:1.5px; color:#b8f0d4; font-weight:700; margin-bottom:18px;">2. Equipo territorial</div>
                        <ul style="list-style:none; padding:0; margin:0 0 18px; font-size:0.95rem; line-height:2;">
                            <li><span style="color:#b8f0d4; font-weight:700;">Referente psicosocial y de alertas:</span> Diana Lozano</li>
                            <li><span style="color:#b8f0d4; font-weight:700;">Referente log&iacute;stica y de FDL:</span> Pamela Bar&oacute;n</li>
                            <li><span style="color:#b8f0d4; font-weight:700;">Referente proceso formativo:</span> Alejandra Sosa Aponte</li>
                            <li><span style="color:#b8f0d4; font-weight:700;">Referente inclusi&oacute;n productiva:</span> Edgardo Montes</li>
                        </ul>
                        <div style="font-weight:700; font-size:0.95rem; line-height:1.8; color:#b8f0d4;">
                            <div>Equipo psicosocial</div>
                            <div>Equipo de alertas</div>
                        </div>
                    </div>

                    <div class="equipo-fila-inferior" style="display:grid; grid-template-columns:1fr 1fr; gap:22px;">
                        <div style="background:#663a93; border-radius:14px; padding:26px 26px 22px; box-shadow:7px 7px 0 #1eaf76; color:#fff; text-align:center;">
                            <div style="font-size:0.75rem; text-transform:uppercase; letter-spacing:1.5px; color:#b8f0d4; font-weight:700; margin-bottom:18px;">3. Equipo metodol&oacute;gico y gesti&oacute;n documental</div>
                            <ul style="list-style:none; padding:0; margin:0; font-size:0.95rem; line-height:2;">
                                <li><span style="color:#b8f0d4; font-weight:700;">Referente metodol&oacute;gico:</span> Ana Mar&iacute;a Altamar</li>
                                <li><span style="color:#b8f0d4; font-weight:700;">Referente control pol&iacute;tico:</span> Alejandro Osorio</li>
                            </ul>
                        </div>

                        <div style="background:#663a93; border-radius:14px; padding:26px 26px 22px; box-shadow:7px 7px 0 #1eaf76; color:#fff; text-align:center;">
                            <div style="font-size:0.75rem; text-transform:uppercase; letter-spacing:1.5px; color:#b8f0d4; font-weight:700; margin-bottom:18px;">4. Equipo administrativo</div>
                            <ul style="list-style:none; padding:0; margin:0 0 18px; font-size:0.95rem; line-height:2;">
                                <li><span style="color:#b8f0d4; font-weight:700;">L&iacute;der administrativo:</span> John Garz&oacute;n</li>
                                <li><span style="color:#b8f0d4; font-weight:700;">L&iacute;der financiero:</span> David Quiceno</li>
                                <li><span style="color:#b8f0d4; font-weight:700;">Gesti&oacute;n documental:</span> Andr&eacute;s Rodr&iacute;guez</li>
                            </ul>
                            <div style="font-weight:700; font-size:0.95rem; color:#b8f0d4;">3 apoyos</div>
                        </div>
                    </div>
                </div>
            </div>"""

# --- Pilares ---
SECCION_PILARES = """\
            <div class="content-section" id="pilares">
                <div class="card">
                    <h2 class="card-title">Pilares del programa J&oacute;venes con Oportunidades</h2>

                    <h3 class="card-subtitle">1. Acompa&ntilde;amiento psicosocial</h3>
                    <p style="line-height:1.7;">Orientaci&oacute;n y seguimiento transversal a cargo de la Secretar&iacute;a Distrital de Integraci&oacute;n Social, para ayudar a los j&oacute;venes en la construcci&oacute;n de su proyecto de vida.</p>

                    <h3 class="card-subtitle">2. Rutas de formaci&oacute;n</h3>
                    <p style="line-height:1.7;">Los participantes eligen entre tres opciones:</p>
                    <div class="rutas-formacion-grid" style="display:grid; grid-template-columns:repeat(3, 1fr); gap:16px; margin-top:18px;">
                        <div style="display:flex; flex-direction:column;">
                            <div style="background:#f4676e; color:#fff; padding:10px 16px 10px 14px; font-weight:800; font-size:0.8rem; letter-spacing:0.5px; text-transform:uppercase; clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%);">Ruta 1</div>
                            <div style="padding:14px 4px 0; border-top:1px dashed #ccc; margin-top:-1px;">
                                <h4 style="color:var(--accent); font-size:0.98rem; font-weight:700; margin:0 0 8px 0;">Educaci&oacute;n flexible</h4>
                                <p style="font-size:0.85rem; line-height:1.55; color:#555; margin:0;">Para personas j&oacute;venes y adultas que buscan completar su educaci&oacute;n media.</p>
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column;">
                            <div style="background:#1eaf76; color:#fff; padding:10px 16px 10px 14px; font-weight:800; font-size:0.8rem; letter-spacing:0.5px; text-transform:uppercase; clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%);">Ruta 2</div>
                            <div style="padding:14px 4px 0; border-top:1px dashed #ccc; margin-top:-1px;">
                                <h4 style="color:var(--accent); font-size:0.98rem; font-weight:700; margin:0 0 8px 0;">Cursos cortos certificados</h4>
                                <p style="font-size:0.85rem; line-height:1.55; color:#555; margin:0;">Formaci&oacute;n de hasta 160 horas para adquirir competencias que aumenten las oportunidades de empleo.</p>
                            </div>
                        </div>
                        <div style="display:flex; flex-direction:column;">
                            <div style="background:#663a93; color:#fff; padding:10px 16px 10px 14px; font-weight:800; font-size:0.8rem; letter-spacing:0.5px; text-transform:uppercase; clip-path:polygon(0 0, calc(100% - 14px) 0, 100% 50%, calc(100% - 14px) 100%, 0 100%);">Ruta 3</div>
                            <div style="padding:14px 4px 0; border-top:1px dashed #ccc; margin-top:-1px;">
                                <h4 style="color:var(--accent); font-size:0.98rem; font-weight:700; margin:0 0 8px 0;">Educaci&oacute;n posmedia de ciclo largo</h4>
                                <p style="font-size:0.85rem; line-height:1.55; color:#555; margin:0;">Acceso a Educaci&oacute;n Superior y Educaci&oacute;n para el Trabajo y Desarrollo Humano, en articulaci&oacute;n con iniciativas como J&oacute;venes a la E.</p>
                            </div>
                        </div>
                    </div>

                    <h3 class="card-subtitle">3. Intermediaci&oacute;n laboral</h3>
                    <p style="line-height:1.7;">Una vez finalizada la formaci&oacute;n, la Secretar&iacute;a de Desarrollo Econ&oacute;mico lidera la intermediaci&oacute;n laboral, con acceso a servicios de registro de empleo a trav&eacute;s de plataformas como Talento Capital.</p>

                    <h3 class="card-subtitle">4. Transferencias monetarias condicionadas</h3>
                    <p style="line-height:1.7;">Apoyos econ&oacute;micos que dependen del cumplimiento de las actividades de cada ruta. Los topes m&aacute;ximos son: hasta <strong>$1.200.000</strong> para cursos cortos, hasta <strong>$3.200.000</strong> para educaci&oacute;n flexible, y montos desde <strong>$2.700.000</strong> hasta <strong>$4.300.000</strong> en educaci&oacute;n posmedia y universitaria.</p>
                </div>
            </div>"""


# --- Módulos de Proyecto de Vida ---
# Texto literal redactado por David (equipo psicosocial JCO).
# Diseño: cabecera oscura con número de color en flecha y cuerpo crema con
# el párrafo completo. El número del módulo va en la flecha lateral, así que
# el título solo lleva el nombre del módulo (sin "MÓDULO N:").
SECCION_MODULOS_PROYECTO_VIDA = """\
            <div class="content-section" id="modulos_proyecto_vida">
                <div class="card">
                    <h2 class="card-title">Los 7 m&oacute;dulos de formaci&oacute;n en Proyecto de Vida</h2>
                    <p style="line-height:1.7;">Como parte del proceso de fortalecimiento de capacidades, todos los participantes cursan <strong>siete m&oacute;dulos pedag&oacute;gicos</strong> bajo una metodolog&iacute;a de <strong>talleres presenciales y vivenciales</strong>, estructurados en tres momentos: motivaci&oacute;n, teor&iacute;a y ejercicios pr&aacute;cticos.</p>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#1eaf76;">1</span>
                            <span class="modulo-header-titulo">Sentido de vida</span>
                        </div>
                        <div class="modulo-body">
                            <p>El primer m&oacute;dulo tiene como objetivo aportar al desarrollo del sentido de vida y la visualizaci&oacute;n de metas, con miras al bienestar integral, mediante herramientas de autoconocimiento y proyecci&oacute;n socio-ocupacional. Lo anterior se logra mediante acuerdos previos con los j&oacute;venes para promover espacios seguros para las actividades, donde se recalca que no existen respuestas correctas o incorrectas. Con actividades como el &ldquo;diagrama del Ikigai&rdquo; y &ldquo;Conejos a sus madrigueras&rdquo; se promueve una reflexi&oacute;n sobre la importancia del cuidado propio y cuidado hacia los otros.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#1e7895;">2</span>
                            <span class="modulo-header-titulo">Coaching en finanzas</span>
                        </div>
                        <div class="modulo-body">
                            <p>El segundo m&oacute;dulo busca generar estrategias pr&aacute;cticas que permitan fortalecer habilidades para el entorno financiero y, asimismo, potenciar h&aacute;bitos para un manejo adecuado de las finanzas. A trav&eacute;s del desarrollo de habilidades blandas impartidas durante el m&oacute;dulo, se busca fomentar la asertiva toma de decisiones por medio de h&aacute;bitos financieros saludables. Para lograr este objetivo, el m&oacute;dulo se divide en 2 partes: 1) ense&ntilde;anza de conceptos b&aacute;sicos sobre finanzas personales y 2) la importancia del ahorro.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#1e9da3;">3</span>
                            <span class="modulo-header-titulo">Manejo del estr&eacute;s y la ansiedad</span>
                        </div>
                        <div class="modulo-body">
                            <p>El tercer m&oacute;dulo se plantea entender la diferencia entre estr&eacute;s y ansiedad, identificar sus causas y s&iacute;ntomas, y aprender estrategias efectivas para manejarlos. A trav&eacute;s de t&eacute;cnicas de relajaci&oacute;n, ejercicio f&iacute;sico, planificaci&oacute;n y organizaci&oacute;n para gestionar el tiempo y las cargas de trabajo se busca fortalecer las habilidades para manejar estas emociones de forma asertiva. Por medio de actividades como el &ldquo;medidor emocional&rdquo; y el &ldquo;teatro de la mente&rdquo; se ense&ntilde;a sobre los conceptos b&aacute;sicos necesarios para identificar la ansiedad y el estr&eacute;s, adem&aacute;s de diferenciar sus s&iacute;ntomas e identificar c&oacute;mo manejarlos, ya sea por medio de estrategias de autorregulaci&oacute;n o acudiendo a profesionales cuando los s&iacute;ntomas as&iacute; lo requieran.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#f4676e;">4</span>
                            <span class="modulo-header-titulo">Acoso en el &aacute;mbito educativo, mobbing y resoluci&oacute;n de conflictos</span>
                        </div>
                        <div class="modulo-body">
                            <p>El cuarto m&oacute;dulo se plantea identificar el acoso (Bullying) como una problem&aacute;tica social actual y brindar informaci&oacute;n acerca de los escenarios en los que se presenta, las herramientas existentes para prevenir o afrontar situaciones de acoso, rutas de apoyo y estrategias para la resoluci&oacute;n de conflictos. Por medio de actividades como &ldquo;en tus zapatos&rdquo; y &ldquo;representaciones del acoso&rdquo; se ense&ntilde;an conceptos relevantes sobre el acoso, adem&aacute;s de rutas de atenci&oacute;n para abordar esta problem&aacute;tica cuando se presenta.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#f58b53;">5</span>
                            <span class="modulo-header-titulo">Relaciones saludables y cuidadosas</span>
                        </div>
                        <div class="modulo-body">
                            <p>El quinto m&oacute;dulo tiene como objetivo fortalecer habilidades blandas para la construcci&oacute;n de relaciones saludables y cuidadosas mediante la implementaci&oacute;n de acciones de informaci&oacute;n y formaci&oacute;n. En este m&oacute;dulo se aborda el desarrollo de un &ldquo;botiqu&iacute;n de gesti&oacute;n emocional&rdquo; y &ldquo;la receta de las relaciones sanas y cuidadosas&rdquo;, las cuales buscan ense&ntilde;ar y recopilar las herramientas trabajadas durante la sesi&oacute;n para mejorar el trato de las relaciones saludables en el &aacute;mbito familiar, social, de pareja, laboral, entre otros. Adem&aacute;s, se ense&ntilde;a sobre &ldquo;red flags&rdquo; (banderas rojas), lo cual permite identificar y establecer l&iacute;mites en las relaciones.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#663a93;">6</span>
                            <span class="modulo-header-titulo">Promoci&oacute;n de derechos y habilidades para la vida</span>
                        </div>
                        <div class="modulo-body">
                            <p>El sexto m&oacute;dulo busca incentivar la promoci&oacute;n y protecci&oacute;n de los derechos mediante diferentes estrategias que contribuyan en su conocimiento, difusi&oacute;n y respeto, e identificar las habilidades sociales de cada joven para que estas puedan aportar de manera significativa su desarrollo e interacci&oacute;n con la sociedad en diferentes contextos. Para lograr este objetivo, el m&oacute;dulo aborda tem&aacute;ticas relacionadas con la promoci&oacute;n y la protecci&oacute;n de los derechos, enfocados principalmente en el derecho a la vida, a la salud, al trabajo, a la educaci&oacute;n y los derechos sexuales y reproductivos.</p>
                        </div>
                    </div>

                    <div class="modulo-acordeon">
                        <div class="modulo-header">
                            <span class="modulo-header-num" style="background:#2fa4d4;">7</span>
                            <span class="modulo-header-titulo">Activa tu potencial</span>
                        </div>
                        <div class="modulo-body">
                            <p>El &uacute;ltimo m&oacute;dulo tiene como objetivo orientar e incentivar la apropiaci&oacute;n de elementos relevantes que faciliten la participaci&oacute;n en la ruta de formaci&oacute;n corta e intermediaci&oacute;n laboral. En este m&oacute;dulo se ense&ntilde;a en qu&eacute; consiste la ruta de formaci&oacute;n e intermediaci&oacute;n laboral, adem&aacute;s de ense&ntilde;ar los beneficios de participar en la ruta. En este m&oacute;dulo final, los participantes identifican, con base en sus intereses, los procesos para avanzar en las rutas, adem&aacute;s de los canales de comunicaci&oacute;n en caso de requerir apoyo.</p>
                        </div>
                    </div>
                </div>
            </div>"""

# --- Flujo de gestión de la información ---
# Patrón compartido con Forjar y Casas: cada paso con círculo Lucide en accent
# y conector vertical punteado. Sin riel con gradiente, sin badges de "actor".
# El actor responsable se muestra como meta-línea bajo el título.
SECCION_GESTION_DATOS = """\
            <div class="content-section" id="flujo_datos">
                <div class="card">
                    <h2 class="card-title">Flujo de gesti&oacute;n de la informaci&oacute;n</h2>
                    <p style="color:#666; margin-bottom:20px;">La gesti&oacute;n de la informaci&oacute;n de JCO est&aacute; dise&ntilde;ada para operar con cohortes de <strong>m&aacute;s de 5.000 j&oacute;venes</strong> que el servicio administra internamente y luego carga en SIRBE, en lugar de registrar cada joven uno por uno.</p>

                    <div class="flujo-pasos">
                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="users"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 01</p>
                                <h3 class="flujo-titulo">Ingreso masivo por cohortes</h3>
                                <p class="flujo-responsable">DADE</p>
                                <p class="flujo-texto">En JCO no se llena una ficha SIRBE individual como en Casas de Juventud o Forjar. El ingreso se realiza mediante el <strong>cargue masivo de una base de datos</strong>, porque el servicio opera con cohortes grandes de miles de j&oacute;venes focalizados previamente por la <strong>Direcci&oacute;n de An&aacute;lisis y Dise&ntilde;o Estrat&eacute;gico (DADE)</strong>.</p>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="layers"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 02</p>
                                <h3 class="flujo-titulo">Estructura en SIRBE: servicios sociales con modalidades</h3>
                                <p class="flujo-responsable">Equipo psicosocial</p>
                                <p class="flujo-texto">JCO est&aacute; parametrizado en SIRBE como <strong>&ldquo;servicio social&rdquo;</strong> (no como curso) y tiene configuradas <strong>modalidades que reflejan las rutas del servicio</strong>. Todos los j&oacute;venes inician su proceso por la modalidad de <strong>Proyecto de Vida</strong> y lo finalizan saliendo por <strong>Intermediaci&oacute;n laboral</strong>. La parametrizaci&oacute;n es estricta: no hay campos de texto abierto, los profesionales seleccionan siempre de listas preestablecidas.</p>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="clipboard-list"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 03</p>
                                <h3 class="flujo-titulo">Tres tipos de actuaciones</h3>
                                <p class="flujo-responsable">Equipo psicosocial</p>
                                <p class="flujo-texto">El seguimiento a los j&oacute;venes se registra en SIRBE con tres tipos de actuaci&oacute;n:</p>
                                <ul class="flujo-lista">
                                    <li><strong>Estado:</strong> situaci&oacute;n actual del joven (en atenci&oacute;n, suspendido, transferido o retirado).</li>
                                    <li><strong>Intervenci&oacute;n:</strong> registra si el joven cumpli&oacute; o incumpli&oacute; las condiciones para autorizar el pago de la transferencia monetaria.</li>
                                    <li><strong>Seguimiento:</strong> registra el paso del joven entre modalidades dentro de su ruta.</li>
                                </ul>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="bar-chart-3"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 04</p>
                                <h3 class="flujo-titulo">Anal&iacute;tica propia del servicio</h3>
                                <p class="flujo-responsable">Equipo de anal&iacute;tica</p>
                                <p class="flujo-texto">La informaci&oacute;n oficial del servicio es gestionada directamente por el <strong>equipo de anal&iacute;tica</strong> de JCO mediante bases de datos propias.</p>
                            </div>
                        </div>
                    </div>

                    <h3 class="card-subtitle" style="margin-top:30px;">Diagrama de flujo del proceso</h3>
                    <p style="color:#666; font-size:0.85rem; margin-bottom:12px;">Representaci&oacute;n visual del ciclo de recolecci&oacute;n y digitaci&oacute;n en SIRBE para J&oacute;venes con Oportunidades.</p>
                    <div style="overflow-x:auto; margin-bottom:8px;">%%SVG_DIAGRAMA_JCO%%</div>
                </div>
            </div>"""

# --- Documentación ---
# Lista cronológica y temática de la documentación oficial del servicio JCO y
# su antecesor Parceros. Los 45 PDFs viven en `Documentación JCO/` agrupados
# en 6 carpetas; este bloque los expone como cards linkeadas para que el equipo
# (y la Contraloría) pueda navegarlos directo desde el gestor. Cada categoría
# abre con su subtítulo de color y muestra sus documentos en grid.
SECCION_DOCUMENTACION = """\
            <div class="content-section" id="documentacion">
                <div class="card">
                    <h2 class="card-title">Documentaci&oacute;n</h2>
                    <p style="line-height:1.7;">Documentos oficiales del servicio J&oacute;venes con Oportunidades y su antecesor Parceros por Bogot&aacute;: manuales del servicio, instructivos de pago, documentaci&oacute;n del Convenio 1285-2025, portafolios de servicios SDIS (cronolog&iacute;a completa) y marco normativo.</p>

%%BLOQUES_DOCS%%
                </div>
            </div>"""

SECCION_DOCUMENTACION = SECCION_DOCUMENTACION.replace("%%BLOQUES_DOCS%%", _generar_bloques_documentos())

# --- Estadísticas ---
SECCION_ESTADISTICAS = """\
            <div class="content-section" id="estadisticas">
                <div class="card">
                    <h2 class="card-title">Estad&iacute;sticas</h2>
                    <iframe title="Seguimiento t&eacute;cnico" width="100%" height="600" src="https://app.powerbi.com/view?r=eyJrIjoiMzRiNWRkMDQtNThmNC00Yzk5LThjNTItOWI4MzZkYzYwM2EzIiwidCI6ImIzZTMwODA4LWU5YTgtNGYyYS05YmMxLWE3NjBhZTkxMGNmNSIsImMiOjR9" frameborder="0" allowFullScreen="true" style="border:1px solid #e0e0e0; border-radius:8px;"></iframe>
                </div>
            </div>"""

# ============================================================
# 5. JavaScript — navegación del sidebar
# ============================================================

JAVASCRIPT = """\
function toggleSection(el) {
    el.classList.toggle('active');
    var items = el.nextElementSibling;
    if (items) items.classList.toggle('show');
}
function showContent(id) {
    document.querySelectorAll('.content-section').forEach(function(s) { s.classList.remove('active'); });
    document.querySelectorAll('.sidebar-item').forEach(function(s) { s.classList.remove('active'); });
    var el = document.getElementById(id);
    if (el) el.classList.add('active');
    if (event && event.target && event.target.classList.contains('sidebar-item')) {
        event.target.classList.add('active');
    }
}"""

# ============================================================
# 6. Ensamblaje — une todas las piezas en el HTML final
# ============================================================

def generar_html():
    """Arma el documento HTML completo a partir de las variables."""

    # Se unen todas las secciones de contenido en orden
    seccion_gestion = SECCION_GESTION_DATOS.replace(
        "%%SVG_DIAGRAMA_JCO%%", svg_diagrama_jco()
    )

    secciones = "\n\n".join([
        SECCION_WELCOME,
        SECCION_LINEA_TIEMPO,
        SECCION_A_TENER_EN_CUENTA,
        SECCION_EQUIPO,
        SECCION_PILARES,
        SECCION_MODULOS_PROYECTO_VIDA,
        SECCION_DOCUMENTACION,
        seccion_gestion,
        seccion_aliados_jco(),
        SECCION_ESTADISTICAS,
    ])

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestor de conocimiento - J&oacute;venes con Oportunidades</title>
    <style>
{CSS}</style>
</head>
<body>
{HEADER}
    <div class="container">
{SIDEBAR}
        <main class="main-content">

{secciones}
        </main>
    </div>
    <script>
{JAVASCRIPT}
</script>
    <script src="https://unpkg.com/lucide@latest/dist/umd/lucide.min.js"></script>
    <script>lucide.createIcons();</script>
</body>
</html>
"""
    return html


# ============================================================
# 7. Escritura del archivo
# ============================================================

if __name__ == "__main__":
    # Raíz del proyecto (un nivel arriba de scripts/)
    directorio = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    archivo_salida = os.path.join(directorio, "gestion_conocimiento_jco_2025.html")

    html = generar_html()

    with open(archivo_salida, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Archivo generado: {archivo_salida}")
