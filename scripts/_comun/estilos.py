"""
CSS base compartido por todos los generadores del gestor de conocimiento.

Antes de abril de 2026 cada generador tenía su propio bloque CSS duplicado
(cuatro copias con pequeñas diferencias). Esto se centralizó aquí para que
cambiar un color o un espaciado afecte a todos los servicios al mismo tiempo.

Uso:

    from _comun.estilos import construir_css

    CSS_BASE = construir_css(accent="#8e6bbf")  # JCO
    CSS_BASE = construir_css(accent="#e67e22", accent_bg="#fdf2e9", accent_border="#f5d9b5")  # Parche seguro

Para agregar reglas específicas de un servicio (como las barras horizontales
de Casas de Juventud), pasar el CSS adicional en el parámetro `extras`, que
se concatena al final y puede hacer override de reglas base.
"""


# Reglas comunes a los cinco generadores (forjar, jco, alertas, home, juventud).
# Usa variables CSS --accent, --accent-bg y --accent-border que se definen
# en el :root generado dinámicamente por construir_css.
_CSS_BASE = """\
@import url('https://fonts.googleapis.com/css2?family=Anton&family=Antonio:wght@400;700&family=Figtree:wght@400;500;600;700;800&display=swap');
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Figtree', 'Segoe UI', sans-serif; background-color: #ffffff; color: #2F3E3C; }
.header { background: #2F3E3C; color: #F8F4E1; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.15); }
.header h1 { font-size: 1.5rem; font-weight: 700; } .header .subtitle { font-size: 0.9rem; opacity: 0.85; }
.header-btns { display: flex; align-items: center; gap: 15px; }
.home-btn { font-size: 1.5rem; cursor: pointer; padding: 5px 12px; border-radius: 8px; transition: background 0.2s; text-decoration: none; color: #F8F4E1; }
.home-btn:hover { background: rgba(255,255,255,0.15); }
.container { display: flex; min-height: calc(100vh - 80px); }
.sidebar { width: 280px; background: #fff; border-right: 1px solid #e0e0e0; padding: 20px 0; overflow-y: auto; }
.sidebar-section { margin-bottom: 10px; }
.sidebar-title { padding: 12px 20px; font-weight: 600; color: var(--accent); cursor: pointer; display: flex; justify-content: space-between; align-items: center; transition: background 0.2s; }
.sidebar-title:hover { background: transparent; opacity: 0.75; }
.sidebar-title .arrow { display: inline-block; width: 8px; height: 8px; border-top: 2px solid var(--accent); border-right: 2px solid var(--accent); transform: rotate(45deg); font-size: 0; color: transparent; transition: transform 0.25s ease; margin-right: 4px; }
.sidebar-title.active .arrow { transform: rotate(135deg); margin-top: -4px; }
.sidebar-items { display: none; padding-left: 20px; }
.sidebar-items.show { display: block; }
.sidebar-item { padding: 10px 20px; cursor: pointer; font-size: 0.9rem; color: #3A3A3A; transition: all 0.2s; }
.sidebar-item:hover { background: transparent; color: var(--accent); }
.sidebar-item.active { background: transparent; color: var(--accent); font-weight: 600; }
.sidebar-link { display: block; padding: 12px 20px; font-weight: 600; color: var(--accent); text-decoration: none; transition: background 0.2s; }
.sidebar-link:hover { background: transparent; opacity: 0.75; }
.main-content { flex: 1; padding: 30px; overflow-y: auto; }
.content-section { display: none; }
.content-section.active { display: block; }
.card { background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); padding: 25px; margin-bottom: 20px; }
.card-title { font-size: 1.4rem; color: var(--accent); margin-bottom: 15px; }
.card-subtitle { font-size: 1.15rem; color: var(--accent); font-weight: 500; margin: 25px 0 10px 0; opacity: 0.85; }
.card p { margin-bottom: 15px; }
.badge { display: inline-block; padding: 5px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 500; margin-right: 8px; margin-bottom: 8px; }
.badge-primary { background: var(--accent-border); color: var(--accent); }
.badge-warning { background: #FFF3E0; color: #F58B53; }
.badge-info { background: #e0f4f5; color: #1E9DA3; }
.welcome-section { text-align: center; padding: 60px 20px; }
.welcome-section h2 { font-size: 2rem; color: var(--accent); margin-bottom: 20px; }
.welcome-section p { font-size: 1.1rem; color: #666; max-width: 700px; margin: 0 auto 20px; line-height: 1.7; }
.methodology-box { background: var(--accent-bg); border-left: 4px solid var(--accent); padding: 15px 20px; margin: 15px 0; border-radius: 0 8px 8px 0; }
.pending-box { background: #FFF3E0; border-left: 4px solid #F58B53; padding: 15px 20px; margin: 15px 0; border-radius: 0 8px 8px 0; color: #8B5E3C; }
/* Línea de tiempo */
.timeline { position: relative; padding-left: 30px; margin: 20px 0; }
.timeline::before { content: ''; position: absolute; left: 8px; top: 0; bottom: 0; width: 3px; background: var(--accent-border); }
.timeline-item { position: relative; margin-bottom: 25px; }
.timeline-item::before { content: ''; position: absolute; left: -26px; top: 4px; width: 12px; height: 12px; border-radius: 50%; background: var(--accent); border: 3px solid var(--accent-border); }
.timeline-year { font-weight: 700; color: var(--accent); font-size: 1rem; }
.timeline-text { color: #555; font-size: 0.9rem; margin-top: 4px; line-height: 1.5; }
/* Footer */
.footer { text-align: center; padding: 35px 30px; color: #bbb; font-size: 0.82rem; }
/* Tarjetas tipo m&oacute;dulo (cabezal con flecha de color y cuerpo crema con texto largo).
   Usado en M&oacute;dulos de Proyecto de Vida (JCO) y Modalidades de atenci&oacute;n (Forjar).
   La variante "-num" lleva el n&uacute;mero del &iacute;tem; "-flag" es solo pesta&ntilde;a de color sin texto. */
.modulo-acordeon { margin-bottom: 14px; }
.modulo-header { width: 100%; display: flex; align-items: stretch; }
.modulo-header-num { color: #fff; min-width: 72px; padding: 16px 24px 16px 18px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 1.6rem; clip-path: polygon(0 0, calc(100% - 20px) 0, 100% 50%, calc(100% - 20px) 100%, 0 100%); z-index: 2; }
.modulo-header-flag { min-width: 36px; padding: 0 18px 0 0; clip-path: polygon(0 0, calc(100% - 18px) 0, 100% 50%, calc(100% - 18px) 100%, 0 100%); z-index: 2; align-self: stretch; }
.modulo-header-titulo { background: #2F3E3C; color: #F8F4E1; flex: 1; padding: 16px 20px 16px 28px; font-size: 0.95rem; font-weight: 700; line-height: 1.4; margin-left: -20px; border-radius: 0 8px 8px 0; display: flex; align-items: center; letter-spacing: 0.02em; }
/* Cuando el cabezal no lleva flecha ni pesta&ntilde;a, el t&iacute;tulo no compensa con margen negativo. */
.modulo-header-titulo:first-child { margin-left: 0; padding-left: 22px; border-radius: 8px 8px 0 0; }
.modulo-body { background: #F8F4E1; color: #2F3E3C; padding: 20px 26px 22px; border-radius: 0 0 8px 8px; }
.modulo-body p { font-size: 0.92rem; line-height: 1.75; margin: 0 0 10px; color: #3A3A3A; }
.modulo-body p:last-child { margin-bottom: 0; }
@media (max-width: 768px) {
    .container { flex-direction: column; }
    .sidebar { width: 100%; border-right: none; border-bottom: 1px solid #e0e0e0; }
    .rutas-formacion-grid { grid-template-columns: 1fr !important; gap: 20px !important; }
    .equipo-fila-inferior { grid-template-columns: 1fr !important; gap: 22px !important; }
}"""


# Línea de tiempo en chevrones encadenados horizontales.
# Componente compartido por Forjar, JCO y Casas de Juventud. El número de hitos es
# flexible: cada wrapper pasa --lt-cols inline (ej. style="--lt-cols: 5"). Los 7
# colores siguen la paleta oficial extendida SDIS Juventud.
# Se mantiene como constante separada (en vez de incrustada en _CSS_BASE) porque
# generar_juventud.py todavía no consume construir_css() y necesita importarla
# por aparte para evitar duplicar el CSS inline.
CSS_LINEA_TIEMPO_CHEVRON = """\
.linea-tiempo { display: grid; grid-template-columns: repeat(var(--lt-cols, 7), 1fr); grid-template-rows: auto auto; row-gap: 22px; }
.lt-hito { display: contents; }
.lt-hito > .lt-chevron { grid-row: 1; }
.lt-hito > .lt-cuerpo { grid-row: 2; }
.lt-chevron { display: flex; align-items: center; justify-content: flex-start; gap: 10px; padding: 16px 22px 16px 18px; font-family: 'Anton','Segoe UI',sans-serif; color: #ffffff; font-size: 1.1rem; letter-spacing: 0.02em; clip-path: polygon(0 0, calc(100% - 18px) 0, 100% 50%, calc(100% - 18px) 100%, 0 100%); }
.lt-hito + .lt-hito > .lt-chevron { margin-left: -18px; }
.lt-chevron svg { width: 20px; height: 20px; flex-shrink: 0; }
.lt-c1 { background: #f4676e; } .lt-c2 { background: #1eaf76; } .lt-c3 { background: #663a93; } .lt-c4 { background: #f58b53; } .lt-c5 { background: #1e9da3; } .lt-c6 { background: #2fa4d4; } .lt-c7 { background: #1e7895; }
.lt-cuerpo { padding: 0 14px; border-left: 1px dashed rgba(47, 62, 60, 0.18); }
.lt-hito:first-child > .lt-cuerpo { border-left: none; }
.lt-icono { display: flex; justify-content: center; margin: 16px 0 12px; }
.lt-icono svg { width: 38px; height: 38px; stroke-width: 1.6; }
.lt-i1 svg { color: #f4676e; } .lt-i2 svg { color: #1eaf76; } .lt-i3 svg { color: #663a93; } .lt-i4 svg { color: #f58b53; } .lt-i5 svg { color: #1e9da3; } .lt-i6 svg { color: #2fa4d4; } .lt-i7 svg { color: #1e7895; }
.lt-titulo { font-family: 'Antonio','Anton','Segoe UI',sans-serif; font-weight: 700; font-size: 1rem; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.02em; text-align: center; }
.lt-t1 { color: #f4676e; } .lt-t2 { color: #1eaf76; } .lt-t3 { color: #663a93; } .lt-t4 { color: #f58b53; } .lt-t5 { color: #1e9da3; } .lt-t6 { color: #2fa4d4; } .lt-t7 { color: #1e7895; }
.lt-texto { font-family: 'Figtree','Segoe UI',sans-serif; font-weight: 500; font-size: 0.8rem; color: #3a3a3a; line-height: 1.6; }
.lt-texto strong { font-weight: 700; color: #2f3e3c; }
@media (max-width: 900px) {
    .linea-tiempo { grid-template-columns: 1fr; grid-template-rows: none; row-gap: 0; }
    .lt-hito { display: block; margin-bottom: 24px; }
    .lt-hito:last-child { margin-bottom: 0; }
    .lt-hito > .lt-chevron, .lt-hito > .lt-cuerpo { grid-row: auto; }
    .lt-chevron { clip-path: none; padding: 14px 20px; border-radius: 6px; font-size: 1.15rem; }
    .lt-hito + .lt-hito > .lt-chevron { margin-left: 0; }
    .lt-cuerpo { border-left: 1px dashed rgba(47, 62, 60, 0.3); padding: 14px 0 4px 18px; margin: 10px 0 0 14px; }
    .lt-hito:first-child > .lt-cuerpo { border-left: 1px dashed rgba(47, 62, 60, 0.3); }
    .lt-icono { justify-content: flex-start; margin: 0 0 8px; }
    .lt-icono svg { width: 30px; height: 30px; }
    .lt-titulo { text-align: left; }
}"""


# Colores por servicio. Cuando se agrega un servicio nuevo, registrarlo aquí.
COLORES = {
    "juventud": {
        "accent": "#663A93",
        "accent_bg": "#f5f0fa",
        "accent_border": "#ede7f6",
    },
    "forjar": {
        "accent": "#5f9ea0",
        "accent_bg": "#edf6f6",
        "accent_border": "#d4eaeb",
    },
    "jco": {
        "accent": "#663a93",
        "accent_bg": "#f1ebf7",
        "accent_border": "#ddd0ea",
    },
    "alertas": {
        "accent": "#f58b53",
        "accent_bg": "#fdeee2",
        "accent_border": "#f5cdb0",
    },
    "home": {
        "accent": "#663A93",
        "accent_bg": "#f5f0fa",
        "accent_border": "#ede7f6",
    },
}


def construir_css(
    accent: str = "#663A93",
    accent_bg: str = "#f5f0fa",
    accent_border: str = "#ede7f6",
    extras: str = "",
) -> str:
    """Devuelve el CSS completo del servicio.

    Construye el bloque `:root` con los colores pedidos, antepone el CSS base
    compartido (que usa las variables --accent, --accent-bg, --accent-border)
    y añade al final las reglas extra del servicio si las hay.
    """
    root = (
        f":root {{ --accent: {accent}; --accent-bg: {accent_bg}; "
        f"--accent-border: {accent_border}; }}\n"
    )
    # Se inserta :root después del @import y del selector universal, para
    # que la cascada CSS funcione igual que antes del refactor.
    lineas = _CSS_BASE.split("\n")
    resultado = []
    root_insertado = False
    for linea in lineas:
        resultado.append(linea)
        if not root_insertado and linea.startswith("* "):
            resultado.append(root.rstrip("\n"))
            root_insertado = True
    base = "\n".join(resultado)
    # La línea de tiempo en chevrones vive en su propia constante para poder
    # ser importada también por generar_juventud.py (que no usa construir_css).
    # Aquí se concatena para que Forjar, JCO y los demás la reciban automáticamente.
    base += "\n" + CSS_LINEA_TIEMPO_CHEVRON
    if extras:
        base += "\n" + extras.strip()
    return base


def css_para(servicio: str, extras: str = "") -> str:
    """Atajo para construir el CSS de un servicio registrado en COLORES."""
    colores = COLORES[servicio]
    return construir_css(
        accent=colores["accent"],
        accent_bg=colores["accent_bg"],
        accent_border=colores["accent_border"],
        extras=extras,
    )
