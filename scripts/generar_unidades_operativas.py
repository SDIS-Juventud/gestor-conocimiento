# -*- coding: utf-8 -*-
# Genera un mapa unico con las unidades operativas fisicas de la
# Subdireccion para la Juventud: Casas de Juventud y Forjar juntas, para
# ver de un vistazo los espacios de la SDIS en Bogota. Tambien genera la
# pagina unidades_operativas.html que muestra ese mapa con el encabezado
# y pie institucionales.
#
# Reutiliza los mismos datos y estilo de los mapas individuales que ya
# generan generar_juventud.py (Casas, rojo) y generar_gc_forjar.py
# (Forjar, teal). No reemplaza esos mapas: crea uno adicional combinado.

import os
import re
import json
import unicodedata

import pandas as pd
import folium

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(BASE, "datos")
# Asegura que exista la carpeta html/ (ahi viven las paginas generadas).
os.makedirs(os.path.join(BASE, "html"), exist_ok=True)

# Colores por servicio (mismos que en los mapas individuales).
# Casas usa su azul de pagina (#253C5C); se evito el rojo porque suele
# asociarse a alerta o peligro, y no corresponde al servicio.
RELLENO_CASAS = "#253C5C"   # azul institucional de Casas
BORDE_CASAS = "#253C5C"     # azul institucional
COLOR_FORJAR = "#5f9ea0"    # teal


def normalizar(texto):
    """Normaliza nombres de localidad: mayuscula, sin tildes, N por Ñ y
    sin prefijos LA/LOS/EL, para cruzar con el geojson."""
    texto = unicodedata.normalize('NFC', str(texto)).upper().strip()
    texto = texto.replace('Ñ', 'N').replace('ñ', 'n')
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'^(LA |LOS |EL )', '', texto)
    return texto


def _estabilizar_ids(ruta_mapa):
    """Folium asigna IDs hexadecimales aleatorios en cada corrida, lo que
    ensucia el diff de git. Aqui se reemplazan por IDs estables segun el
    orden de aparicion, asi correr el script con los mismos datos da el
    mismo HTML byte a byte."""
    with open(ruta_mapa, "r", encoding="utf-8") as f:
        html_mapa = f.read()
    patron_ids = re.compile(r'(?<![a-zA-Z0-9_])([a-z_]+)_([a-f0-9]{32})(?![a-f0-9])')
    contadores = {}
    mapeo_ids = {}
    def _id_estable(match):
        prefijo = match.group(1)
        original = match.group(0)
        if original not in mapeo_ids:
            contadores[prefijo] = contadores.get(prefijo, 0) + 1
            mapeo_ids[original] = f"{prefijo}_{contadores[prefijo]:03d}"
        return mapeo_ids[original]
    html_mapa = patron_ids.sub(_id_estable, html_mapa)
    with open(ruta_mapa, "w", encoding="utf-8") as f:
        f.write(html_mapa)


def generar_mapa_combinado():
    """Arma un solo mapa Folium con los puntos de Casas de Juventud y de
    Forjar, con leyenda para distinguirlos. Devuelve (n_casas, n_forjar)."""
    geojson_path = os.path.join(DATOS, "localidades_bogota.geojson")
    dir_casas = os.path.join(DATOS, "directorio_casas_juventud.xlsx")
    dir_forjar = os.path.join(DATOS, "directorio_forjar.xlsx")

    df_casas = pd.read_excel(dir_casas) if os.path.exists(dir_casas) else pd.DataFrame()
    df_forjar = pd.read_excel(dir_forjar) if os.path.exists(dir_forjar) else pd.DataFrame()

    # Conjunto de localidades con al menos una unidad (de cualquiera de los
    # dos servicios) para resaltarlas en el mapa.
    locs = set()
    if not df_casas.empty:
        locs |= set(normalizar(l) for l in df_casas["Localidad"].dropna().unique())
    if not df_forjar.empty:
        locs |= set(normalizar(l) for l in df_forjar["Localidad"].dropna().unique())

    m = folium.Map(location=[4.55, -74.15], zoom_start=10,
                   tiles="CartoDB positron", width="100%", height="100%")

    if os.path.exists(geojson_path):
        with open(geojson_path, encoding="utf-8") as f:
            localidades_gj = json.load(f)

        # Base: todas las localidades tenues, para ver el croquis interno de
        # Bogota. Esta capa no se puede apagar (control=False).
        gj_base = folium.GeoJson(
            localidades_gj, name="Localidades",
            style_function=lambda x: {"fillColor": "#f5f5f5", "color": "#d8d8d8", "weight": 0.8, "fillOpacity": 0.3},
            tooltip=folium.GeoJsonTooltip(fields=["nombre"], aliases=["Localidad:"]),
            control=False)
        gj_base.add_to(m)

        # Perimetro de Bogota: une todas las localidades en una sola silueta
        # y la dibuja como linea por encima, para que el croquis se vea
        # siempre (de lejos o de cerca).
        from shapely.ops import unary_union
        from shapely.geometry import shape, mapping
        bogota = unary_union([shape(f["geometry"]) for f in localidades_gj["features"]])
        folium.GeoJson(mapping(bogota), name="Bogotá", control=False,
            style_function=lambda x: {"fillOpacity": 0, "color": "#808080", "weight": 2.5}).add_to(m)

        # Capa que se puede prender o apagar: localidades CON unidades.
        feats_con = {"type": "FeatureCollection",
                     "features": [f for f in localidades_gj["features"]
                                  if normalizar(f["properties"]["nombre"]) in locs]}
        fg_con = folium.FeatureGroup(name="Localidades con unidades", show=True)
        folium.GeoJson(feats_con,
            style_function=lambda x: {"fillColor": "#cdd6e0", "color": "#253C5C", "weight": 1.4, "fillOpacity": 0.55},
            tooltip=folium.GeoJsonTooltip(fields=["nombre"], aliases=["Localidad:"])).add_to(fg_con)
        fg_con.add_to(m)

        # Encuadrar a toda Bogota (incluida Sumapaz).
        m.fit_bounds(gj_base.get_bounds())

    # Marcadores de Casas de Juventud (rojo con borde azul)
    for _, row in df_casas.iterrows():
        if pd.notna(row.get("Latitud")) and pd.notna(row.get("Longitud")):
            popup_html = (
                '<div style="font-family:Arial; min-width:200px;">'
                f'<strong style="color:#253C5C; font-size:14px;">{row["Casa de Juventud"]}</strong><br>'
                f'<span style="color:#666; font-size:12px;">📍 {row["Localidad"]}</span><br>'
                f'<span style="font-size:11px;">{row["Dirección"]}</span><br>'
                f'<span style="font-size:11px; color:#888;">Barrio: {row.get("Barrio", "")}</span></div>'
            )
            folium.CircleMarker(
                location=[row["Latitud"], row["Longitud"]], radius=7,
                color=BORDE_CASAS, fill=True, fill_color=RELLENO_CASAS,
                fill_opacity=0.9, weight=2,
                popup=folium.Popup(popup_html, max_width=250),
                tooltip=str(row["Casa de Juventud"])
            ).add_to(m)

    # Marcadores de Forjar (teal)
    for _, row in df_forjar.iterrows():
        if pd.notna(row.get("Latitud")) and pd.notna(row.get("Longitud")):
            popup_html = (
                '<div style="font-family:Arial; min-width:200px;">'
                f'<strong style="color:#5f9ea0; font-size:14px;">{row["Nombre unidad operativa"]}</strong><br>'
                f'<span style="color:#666; font-size:12px;">📍 {row["Localidad"]}</span><br>'
                f'<span style="font-size:11px;">{row["Dirección"]}</span></div>'
            )
            folium.CircleMarker(
                location=[row["Latitud"], row["Longitud"]], radius=8,
                color=COLOR_FORJAR, fill=True, fill_color=COLOR_FORJAR,
                fill_opacity=0.9, weight=2,
                popup=folium.Popup(popup_html, max_width=250),
                tooltip=str(row["Nombre unidad operativa"])
            ).add_to(m)

    # Control para prender o apagar la capa de localidades con unidades
    folium.LayerControl(collapsed=False).add_to(m)

    # Leyenda fija que distingue los dos servicios por color
    leyenda = folium.Element(
        '<div style="position:fixed; bottom:24px; left:24px; z-index:9999;'
        ' background:#fff; padding:12px 16px; border-radius:8px;'
        ' box-shadow:0 2px 10px rgba(0,0,0,0.12);'
        ' font-family:Figtree,Arial,sans-serif; font-size:13px; color:#2F3E3C;">'
        '<div style="font-weight:700; margin-bottom:8px;">Unidades operativas</div>'
        '<div style="margin-bottom:6px;"><span style="display:inline-block; width:12px;'
        ' height:12px; border-radius:50%; background:#253C5C; border:2px solid #253C5C;'
        ' margin-right:8px; vertical-align:middle;"></span>Casas de Juventud</div>'
        '<div><span style="display:inline-block; width:12px; height:12px;'
        ' border-radius:50%; background:#5f9ea0; border:2px solid #5f9ea0;'
        ' margin-right:8px; vertical-align:middle;"></span>Servicio Forjar Restaurativo</div></div>'
    )
    m.get_root().html.add_child(leyenda)

    ruta_mapa = os.path.join(BASE, "html", "mapa_unidades_operativas.html")
    m.save(ruta_mapa)
    _estabilizar_ids(ruta_mapa)
    return len(df_casas), len(df_forjar)


def _tabla_unidades(df, col_nombre, accent_link):
    """Arma una tabla HTML con las unidades de un servicio (nombre,
    localidad y direccion con enlace a Google Maps), con el mismo formato
    que las tablas de directorio de las paginas de servicio: header oscuro
    y filas alternadas."""
    if df is None or df.empty:
        return '<p style="color:#888;">Sin informaci&oacute;n disponible.</p>'
    filas = ""
    for idx, row in df.iterrows():
        bg = "#fafafa" if idx % 2 == 0 else "#fff"
        nombre = str(row.get(col_nombre, "") or "")
        loc = str(row.get("Localidad", "") or "")
        direccion = str(row.get("Dirección", "") or "")
        link = row.get("Link Google Maps", "")
        if pd.notna(link) and str(link).strip():
            dir_html = f'<a href="{link}" target="_blank" style="color:{accent_link};">{direccion}</a>'
        else:
            dir_html = direccion
        filas += (
            f'<tr style="background:{bg}; border-bottom:1px solid #e0e0e0;">'
            f'<td style="padding:12px 14px; vertical-align:top;"><strong>{nombre}</strong></td>'
            f'<td style="padding:12px 14px; vertical-align:top;">{loc}</td>'
            f'<td style="padding:12px 14px; vertical-align:top;">{dir_html}</td></tr>'
        )
    return (
        '<table style="width:100%; border-collapse:collapse; font-size:0.85rem;">'
        '<thead><tr style="background:#2F3E3C; color:#F8F4E1;">'
        '<th style="padding:12px 14px; text-align:left; font-weight:700;">Unidad operativa</th>'
        '<th style="padding:12px 14px; text-align:left; font-weight:700;">Localidad</th>'
        '<th style="padding:12px 14px; text-align:left; font-weight:700;">Direcci&oacute;n</th>'
        f'</tr></thead><tbody>{filas}</tbody></table>'
    )


def generar_pagina(df_casas=None, df_forjar=None):
    """Crea unidades_operativas.html: header institucional, el mapa
    combinado y, debajo, las tablas de unidades de Casas y de Forjar."""
    if df_casas is None:
        p = os.path.join(DATOS, "directorio_casas_juventud.xlsx")
        df_casas = pd.read_excel(p) if os.path.exists(p) else pd.DataFrame()
    if df_forjar is None:
        p = os.path.join(DATOS, "directorio_forjar.xlsx")
        df_forjar = pd.read_excel(p) if os.path.exists(p) else pd.DataFrame()

    tabla_casas = _tabla_unidades(df_casas, "Casa de Juventud", "#253C5C")
    tabla_forjar = _tabla_unidades(df_forjar, "Nombre unidad operativa", "#5f9ea0")

    html = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Unidades operativas - Subdirecci&oacute;n para la Juventud</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Anton&family=Figtree:wght@400;500;600;700;800&display=swap');
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Figtree', 'Segoe UI', sans-serif; background: #ffffff; color: #2F3E3C; }
        /* Mismo header que las paginas de servicio (barra verde institucional) */
        .header { background: #2F3E3C; color: #F8F4E1; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.15); }
        .header h1 { font-size: 1.5rem; font-weight: 700; }
        .header .subtitle { font-size: 0.9rem; opacity: 0.85; }
        .header-btns { display: flex; align-items: center; gap: 15px; }
        .home-btn { font-size: 1.5rem; cursor: pointer; padding: 5px 12px; border-radius: 8px; transition: background 0.2s; text-decoration: none; color: #F8F4E1; }
        .home-btn:hover { background: rgba(255,255,255,0.15); }
        .main { max-width: 1100px; margin: 0 auto; padding: 40px 30px 50px; }
        /* Mismo formato de tarjeta y titulo de seccion que las demas paginas */
        .card { background: #fff; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); padding: 25px; margin-bottom: 20px; }
        .card-title { font-size: 1.4rem; color: #253C5C; margin-bottom: 15px; }
        .intro { font-size: 0.95rem; color: #555; line-height: 1.6; max-width: 760px; margin-bottom: 20px; }
        .mapa-wrapper { border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden; }
        .mapa-wrapper iframe { width: 100%; height: 650px; border: 0; display: block; }
        @media (max-width: 700px) {
            .main { padding: 30px 18px 40px; }
            .mapa-wrapper iframe { height: 520px; }
        }
    </style>
</head>
<body>
    <header class="header">
        <div>
            <h1>Gestor de conocimiento - Unidades operativas</h1>
            <div class="subtitle">Subdirecci&oacute;n para la Juventud | SDIS</div>
        </div>
        <div class="header-btns">
            <a class="home-btn" href="../index.html" title="Todos los servicios">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#F8F4E1" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </a>
        </div>
    </header>
    <main class="main">
        <div class="card">
            <h2 class="card-title">Unidades operativas</h2>
            <div class="mapa-wrapper">
                <iframe src="mapa_unidades_operativas.html" loading="lazy" title="Mapa de unidades operativas"></iframe>
            </div>
        </div>
        <div class="card">
            <h2 class="card-title" style="color:#253C5C;">Casas de Juventud</h2>
            %%TABLA_CASAS%%
        </div>
        <div class="card">
            <h2 class="card-title" style="color:#5f9ea0;">Servicio Forjar Restaurativo</h2>
            %%TABLA_FORJAR%%
        </div>
    </main>
    <footer style="background:#3a3a3a; padding:18px 30px; display:flex; justify-content:space-between; align-items:center;">
        <img src="../imagenes/Footer1.png" alt="Distrito Joven" style="height:40px; object-fit:contain;">
        <img src="../imagenes/Footer2.png" alt="Secretar&iacute;a de Integraci&oacute;n Social" style="height:40px; object-fit:contain;">
    </footer>
</body>
</html>"""
    html = html.replace("%%TABLA_CASAS%%", tabla_casas).replace("%%TABLA_FORJAR%%", tabla_forjar)
    ruta = os.path.join(BASE, "html", "unidades_operativas.html")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(html)


if __name__ == "__main__":
    print("Generando mapa combinado de unidades operativas...\n")
    n_casas, n_forjar = generar_mapa_combinado()
    generar_pagina()
    print(f"  Mapa: mapa_unidades_operativas.html ({n_casas} Casas + {n_forjar} Forjar)")
    print("  Pagina: unidades_operativas.html")
    print("\nListo.")
