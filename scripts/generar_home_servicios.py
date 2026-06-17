# Genera el home (index.html) del gestor de conocimiento de la
# Subdirección para la Juventud: la rejilla de tarjetas de servicios.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Raíz del proyecto (un nivel arriba de scripts/)
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(BASE, "datos")

# =====================================================================
# INDEX.HTML — Home de todos los servicios
# =====================================================================
def generar_index():
    # Cada servicio tiene:
    #   alt     → texto alternativo del logo (para accesibilidad)
    #   archivo → HTML destino del gestor
    #   imagen  → logo del servicio
    #   desc    → descripcion que aparece sobre el logo
    servicios = [
        {
            "alt": "Casas de Juventud",
            "archivo": "html/gestion_conocimiento_juventud.html",
            "imagen": "imagenes/servicios/casas-de-juventud.png",
            "desc": "Espacios distritales para j&oacute;venes entre 14 y 28 a&ntilde;os. Oferta integral a trav&eacute;s de 5 ejes: bienestar, cultura, inclusi&oacute;n, liderazgo y SIDICU.",
        },
        {
            "alt": "J&oacute;venes con oportunidades",
            "archivo": "html/gestion_conocimiento_jco.html",
            "imagen": "imagenes/servicios/jovenes-con-oportunidades.png",
            "desc": "Formaci&oacute;n, apoyo psicosocial, transferencias monetarias condicionadas y acompa&ntilde;amiento laboral para j&oacute;venes de 14 a 28 a&ntilde;os en condici&oacute;n de vulnerabilidad.",
        },
        {
            "alt": "Servicio Forjar Restaurativo",
            "archivo": "html/gestion_conocimiento_forjar.html",
            "imagen": "imagenes/servicios/forjar.png",
            "desc": "Servicio de atenci&oacute;n integral, especializada y diferencial que se brinda a adolescentes/j&oacute;venes vinculados al SRPA y sus redes familiares, en el marco de modalidades de atenci&oacute;n no privativas de la libertad, desde un enfoque pedag&oacute;gico y restaurativo.",
        },
        {
            "alt": "Parche seguro",
            "archivo": "html/gestion_conocimiento_alertas.html",
            "imagen": "imagenes/servicios/alertas.png",
            "desc": "Sistema de identificaci&oacute;n y seguimiento de alertas tempranas para la protecci&oacute;n integral de la poblaci&oacute;n joven.",
            # Parche Seguro no es un servicio, es una iniciativa transversal a los tres servicios
            "rotulo": "Iniciativa transversal",
        },
    ]

    # Colores de acento por servicio (borde hover)
    # Azul rey, lila, verde-menta (forjar), naranja (Parche seguro)
    colores = ["#1a237e", "#663A93", "#80cbc4", "#e67e22"]

    tarjetas = ""
    for i, s in enumerate(servicios):
        color = colores[i]
        # Rotulo opcional (ej. Parche Seguro como iniciativa transversal)
        rotulo = s.get("rotulo", "")
        rotulo_html = f'<span class="card-rotulo">{rotulo}</span>' if rotulo else ""
        tarjetas += f"""
            <a class="service-card" href="{s['archivo']}" style="--accent:{color};">
                {rotulo_html}
                <div class="service-desc">{s['desc']}</div>
                <div class="service-logo">
                    <img src="{s['imagen']}" alt="{s['alt']}">
                </div>
            </a>
"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestor de conocimiento - Subdirecci&oacute;n para la Juventud</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Anton&family=Figtree:wght@400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Figtree', 'Segoe UI', sans-serif;
            background-color: #ffffff;
            color: #2F3E3C;
            min-height: 100vh;
        }}

        /* Header banner */
        .header-banner {{
            width: 100%;
            display: block;
        }}
        .header-banner img {{
            width: 100%;
            display: block;
        }}

        /* Contenido */
        .main {{ max-width: 1100px; margin: 0 auto; padding: 60px 30px 60px; }}

        /* Grid */
        .services-grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 72px 36px;
            max-width: 1080px;
            margin: 0 auto;
            padding: 40px 0 40px;
        }}

        /* Tarjetas - flex column, logo sobresale por abajo (mitad dentro / mitad fuera) */
        .service-card {{
            background: rgba(255,255,255,0.75);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-radius: 20px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.04);
            padding: 24px 24px 14px;
            text-align: center;
            text-decoration: none;
            color: inherit;
            transition: transform 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease;
            border: 2px solid rgba(0,0,0,0.04);
            position: relative;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: space-between;
            min-height: 150px;
            overflow: visible;
        }}
        .service-card:hover {{
            transform: translateY(-8px);
            box-shadow: 0 16px 40px rgba(0,0,0,0.1);
            border-color: var(--accent, #663A93);
        }}

        /* Descripci&oacute;n del servicio */
        .service-desc {{
            font-size: 0.88rem;
            color: #555;
            line-height: 1.5;
            text-align: center;
            margin: 0 auto;
            max-width: 340px;
        }}

        /* Logo — sobresale por el borde inferior (mitad dentro, mitad fuera) */
        .service-logo {{
            height: 56px;
            margin: 14px auto -28px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: transform 0.3s;
            position: relative;
            z-index: 2;
        }}
        .service-logo img {{ height: 100%; object-fit: contain; }}
        .service-card:hover .service-logo {{ transform: scale(1.08); }}

        /* Bot&oacute;n de ayuda */
        .help-btn {{
            position: fixed; bottom: 80px; right: 28px; z-index: 100;
            width: 48px; height: 48px; border-radius: 50%;
            background: #2F3E3C; color: #F8F4E1; border: none;
            font-size: 1.4rem; font-weight: 700; cursor: pointer;
            box-shadow: 0 4px 16px rgba(0,0,0,0.18);
            transition: transform 0.2s, box-shadow 0.2s;
            display: flex; align-items: center; justify-content: center;
        }}
        .help-btn:hover {{ transform: scale(1.1); box-shadow: 0 6px 24px rgba(0,0,0,0.25); }}

        /* Modal */
        .modal-overlay {{
            display: none; position: fixed; inset: 0; z-index: 200;
            background: rgba(0,0,0,0.45); align-items: center; justify-content: center;
        }}
        .modal-overlay.show {{ display: flex; }}
        .modal-box {{
            background: #2F3E3C; color: #F8F4E1; border-radius: 14px;
            max-width: 560px; width: 90%; padding: 32px 36px 28px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.35); position: relative;
        }}
        .modal-box h2 {{ font-size: 1.35rem; font-weight: 700; margin-bottom: 6px; }}
        .modal-box .modal-sub {{ font-size: 0.85rem; opacity: 0.75; margin-bottom: 18px; }}
        .modal-box p {{ font-size: 0.92rem; line-height: 1.7; opacity: 0.9; margin-bottom: 12px; }}
        .modal-close {{
            position: absolute; top: 14px; right: 18px;
            background: none; border: none; color: #F8F4E1; font-size: 1.5rem;
            cursor: pointer; opacity: 0.6; transition: opacity 0.2s;
        }}
        .modal-close:hover {{ opacity: 1; }}

        /* Encabezado de seccion: titulo centrado, la jerarquia la da el
           espaciado entre secciones. */
        .seccion {{ margin-bottom: 64px; }}
        .seccion:last-child {{ margin-bottom: 0; }}
        .seccion-eyebrow {{ text-align: center; margin-bottom: 28px; }}
        .eyebrow-titulo {{
            font-family: 'Figtree', sans-serif;
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            color: #2F3E3C;
            font-weight: 700;
            line-height: 1.4;
        }}
        .eyebrow-titulo span {{ color: #888; font-weight: 500; letter-spacing: 0.1em; }}
        .eyebrow-sub {{
            font-family: 'Figtree', sans-serif;
            font-size: 0.8rem;
            color: #888;
            margin-top: 6px;
            letter-spacing: 0.03em;
        }}

        /* Rotulo pequeno dentro de una tarjeta (Parche Seguro: transversal).
           Etiqueta editorial: solo texto, sin caja, en el color de acento
           del servicio. Distingue sin parecer un boton. */
        .card-rotulo {{
            align-self: center;
            font-family: 'Figtree', sans-serif;
            font-size: 0.68rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.13em;
            color: var(--accent, #888);
            margin-bottom: 12px;
        }}

        /* Card del chat (NotebookLM): mismo estilo blanco de las
           service-card. La diferenciacion visual con los servicios reales
           viene del eyebrow "Asistente". */
        .chat-wrapper {{
            display: flex;
            justify-content: center;
            max-width: 1080px;
            margin: 0 auto;
            padding: 20px 0;
        }}
        .chat-card {{
            max-width: 560px;
            width: 100%;
            background: #aacf6a url('imagenes/fondo_verde.png') center/cover no-repeat;
            border: none;
            box-shadow: 0 6px 26px rgba(0,0,0,0.10);
        }}
        .chat-card:hover {{ transform: translateY(-3px); box-shadow: 0 10px 30px rgba(0,0,0,0.16); border-color: transparent; }}
        /* Texto oscuro y en negrilla, para que se lea con fuerza sobre el
           verde texturizado (como las piezas Distrito Joven). */
        .chat-card .service-desc {{ color: #2f3e2a; font-weight: 700; }}
        /* Nota "Vinculado a NotebookLM": resaltado inverso al de Unidades
           operativas (recuadro morado, texto crema), pegado al texto. */
        .geco-nota {{
            display: inline;
            font-family: 'Anton', 'Segoe UI', sans-serif;
            font-weight: 400;
            text-transform: uppercase;
            font-size: 1.4rem;
            letter-spacing: 0.01em;
            line-height: 1.7;
            color: #f4f5de;
            background: #5a4a8c;
            padding: 0.08em 0.3em;
            -webkit-box-decoration-break: clone;
            box-decoration-break: clone;
        }}
        .chat-card .service-desc strong {{
            display: block;
            font-family: 'Anton', 'Segoe UI', sans-serif;
            font-size: 1.55rem;
            color: #253C5C;
            margin-bottom: 8px;
            letter-spacing: 0.01em;
            font-weight: 400;
        }}
        /* El logo GECO CHAT ya viene con su sticker amarillo + texto, asi
           que solo le subimos un poco la altura para que se note mas que
           los logos de los servicios. */
        .chat-card .service-logo {{ height: 80px; margin-bottom: -40px; }}

        /* Tarjeta de Unidades operativas, estilo Distrito Joven: fondo
           morado texturizado (Fondo 1.png) con la ilustracion halftone de
           edificios (iconografia propia SDIS) apoyada sobre una franja
           crema que lleva el enunciado. Pieza con mas caracter que el resto. */
        .uo-card {{
            display: block;
            width: 100%;
            max-width: 520px;
            background: #663A93 url('imagenes/Fondo 1.png') center/cover no-repeat;
            border: none;
            border-radius: 20px;
            box-shadow: 0 6px 26px rgba(0,0,0,0.12);
            padding: 30px 0 0;
            overflow: hidden;
            text-align: center;
        }}
        .uo-card:hover {{ transform: translateY(-3px); box-shadow: 0 10px 30px rgba(0,0,0,0.16); border-color: transparent; }}
        .uo-card .uo-ilustracion {{ padding: 0 24px; }}
        .uo-card .uo-ilustracion img {{ width: 100%; max-width: 380px; height: auto; display: block; margin: 0 auto; position: relative; z-index: 2; }}
        /* El texto va sobre el morado, con un resaltado crema pegado a las
           palabras (highlight), no una franja de lado a lado, como en el
           afiche Distrito Joven. box-decoration-break: clone hace que el
           resaltado abrace cada linea por separado. */
        .uo-banner {{ padding: 2px 26px 30px; text-align: center; }}
        .uo-banner p {{
            display: inline;
            margin: 0;
            font-family: 'Anton', 'Segoe UI', sans-serif;
            font-weight: 400;
            text-transform: uppercase;
            font-size: 1.4rem;
            letter-spacing: 0.01em;
            line-height: 1.7;
            color: #5a4a8c;
            background: #f4f5de;
            padding: 0.08em 0.3em;
            -webkit-box-decoration-break: clone;
            box-decoration-break: clone;
        }}

        @media (max-width: 700px) {{
            .services-grid {{ grid-template-columns: 1fr; gap: 40px; padding: 30px 0 30px; }}
            .main {{ padding: 35px 20px 40px; }}
            .service-desc {{ font-size: 0.9rem; }}
            .service-logo {{ height: 46px; }}
            .chat-wrapper {{ margin-top: 40px; padding: 0 10px; }}
        }}
    </style>
</head>
<body>
    <div class="header-banner">
        <img src="imagenes/Header - gestor.jpeg" alt="Gestor de conocimiento - SDIS Juventud">
    </div>
    <main class="main">
        <section class="seccion">
            <header class="seccion-eyebrow">
                <span class="eyebrow-titulo">Servicios de la Subdirecci&oacute;n <span>para la Juventud</span></span>
                <div class="eyebrow-sub">3 servicios y 1 iniciativa transversal</div>
            </header>
            <div class="services-grid">
{tarjetas}
            </div>
        </section>

        <section class="seccion">
            <header class="seccion-eyebrow">
                <span class="eyebrow-titulo">Unidades operativas <span>espacios f&iacute;sicos en Bogot&aacute;</span></span>
            </header>
            <div class="chat-wrapper">
            <a class="service-card uo-card" href="html/unidades_operativas.html">
                <div class="uo-ilustracion">
                    <img src="imagenes/contexto1.png" alt="Espacios f&iacute;sicos de la Subdirecci&oacute;n">
                </div>
                <div class="uo-banner">
                    <p>D&oacute;nde tiene presencia f&iacute;sica la Subdirecci&oacute;n para la Juventud en Bogot&aacute;</p>
                </div>
            </a>
            </div>
        </section>

        <section class="seccion">
            <header class="seccion-eyebrow">
                <span class="eyebrow-titulo">Asistente <span>para preguntas sobre los servicios</span></span>
            </header>
            <div class="chat-wrapper">
            <a class="service-card chat-card" href="https://notebooklm.google.com/notebook/1b538574-c4c7-4d49-957a-24230c433bc0" target="_blank" rel="noopener" style="--accent:#253C5C;">
                <div class="service-desc">
                    <span class="geco-nota">&iquest;Preguntas? Consulta aqu&iacute;</span>
                    <span style="display:block; margin-top:16px; color:#1a1a1a; font-weight:600; font-size:0.9rem; line-height:1.5;">Servicios, ofertas, datos y normatividad. &Uacute;til para derechos de petici&oacute;n e informes.</span>
                    <span style="display:block; margin-top:10px; color:#5a6a4a; font-weight:500; font-size:0.74rem; line-height:1.45;">Vinculado a NotebookLM (Google). Requiere iniciar sesi&oacute;n con cualquier cuenta de Gmail.</span>
                </div>
                <div class="service-logo">
                    <img src="imagenes/servicios/GECO%20CHAT.png" alt="GECO Chat">
                </div>
            </a>
            </div>
        </section>
    </main>
    <button class="help-btn" onclick="document.getElementById('modal-info').classList.add('show')" title="&iquest;Qu&eacute; es esto?">?</button>

    <div class="modal-overlay" id="modal-info" onclick="if(event.target===this)this.classList.remove('show')">
        <div class="modal-box">
            <button class="modal-close" onclick="document.getElementById('modal-info').classList.remove('show')">&times;</button>
            <h2>Gestor de conocimiento</h2>
            <div class="modal-sub">Subdirecci&oacute;n para la Juventud | SDIS</div>
            <p>Herramienta interna que documenta los procesos, datos, metodolog&iacute;as y aprendizajes de cada servicio de la Subdirecci&oacute;n para la Juventud.</p>
            <p>El objetivo es que cualquier persona del equipo pueda consultar c&oacute;mo funciona cada servicio, qu&eacute; datos se gestionan y cu&aacute;les son los procedimientos vigentes.</p>
        </div>
    </div>

    <footer style="background:#3a3a3a; padding:18px 30px; display:flex; justify-content:space-between; align-items:center;">
        <img src="imagenes/Footer1.png" alt="Distrito Joven" style="height:40px; object-fit:contain;">
        <img src="imagenes/Footer2.png" alt="Secretar&iacute;a de Integraci&oacute;n Social" style="height:40px; object-fit:contain;">
    </footer>
</body>
</html>"""

    ruta = os.path.join(BASE, "index.html")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Generado: index.html ({len(html):,} caracteres)".replace(",", "."))


# =====================================================================
# Ejecutar todo
# =====================================================================
if __name__ == "__main__":
    # Este script solo genera el index (home).
    # Los gestores de cada servicio se generan desde sus scripts individuales:
    # generar_juventud.py, generar_gc_forjar.py, generar_gc_jco.py, generar_gc_alertas.py.
    print("Generando pagina home del gestor de conocimiento...\n")
    generar_index()
    print("\nListo. Index generado.")
