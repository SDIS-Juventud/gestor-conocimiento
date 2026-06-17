# Genera el archivo gestion_conocimiento_forjar_2025.html
# Servicio Forjar Restaurativo - Subdirección para la Juventud, SDIS.
# Separa CSS, sidebar, secciones de contenido y JavaScript en variables
# para facilitar ediciones futuras desde Python.

import os
import re
import sys
import pandas as pd

# Raíz del proyecto (un nivel arriba de scripts/)
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO_SALIDA = os.path.join(BASE, "html", "gestion_conocimiento_forjar.html")
DATOS = os.path.join(BASE, "datos")
# Asegura que exista la carpeta html/ (ahi viven las paginas generadas).
os.makedirs(os.path.join(BASE, "html"), exist_ok=True)

# CSS compartido con los demás generadores
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _comun.estilos import css_para
from _comun.aliados import seccion_forjar as seccion_aliados_forjar
from _comun.diagramas_flujo import svg_diagrama_forjar
from _comun.tablero import powerbi_src
from _comun.reporte_politicas import (
    CSS_REPORTE_POLITICAS,
    leer_politicas,
    html_cards,
)

# CSS extras del servicio. La l&iacute;nea de tiempo en chevrones encadenados
# vive ahora en _comun/estilos.py porque la comparten Forjar, JCO y Casas de
# Juventud. Aqu&iacute; quedan solo las reglas espec&iacute;ficas de Forjar.
EXTRAS_CSS = """\
/* Infograf&iacute;a "A tener en cuenta": 6 fichas crema con manos halftone del branding.
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
   sutil entre &iacute;conos para se&ntilde;alar el flujo. As&iacute; se diferencia visualmente
   de Proceso operativo (cards horizontales clicables con n&uacute;mero grande). */
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
@media (max-width: 720px) {
    .flujo-paso { grid-template-columns: 48px 1fr; gap: 14px; }
    .flujo-paso:not(:last-child)::before { left: 23px; top: 48px; }
    .flujo-icono { width: 40px; height: 40px; }
    .flujo-icono svg { width: 20px; height: 20px; }
}

/* Override Forjar: TODOS los subt&iacute;tulos de tarjeta van en Antonio Bold
   may&uacute;sculas, consistente con .lt-titulo, .atc-titulo, .flujo-titulo,
   .po-stage-title y dem&aacute;s subt&iacute;tulos especializados de la secci&oacute;n.
   Aplica solo a Forjar porque vive en EXTRAS_CSS de este script. */
.card-subtitle {
    font-family: 'Antonio', 'Anton', 'Figtree', sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    font-size: 1.05rem;
    opacity: 1;
}
"""

# CSS del servicio (base compartido + color propio de Forjar + extras)
CSS = css_para("forjar", extras=EXTRAS_CSS + CSS_REPORTE_POLITICAS)

# =====================================================================
# Header
# =====================================================================
HEADER = """\
    <header class="header">
        <div>
            <h1>Gestor de conocimiento - Servicio Forjar Restaurativo</h1>
            <div class="subtitle">Subdirecci&oacute;n para la Juventud | SDIS</div>
        </div>
        <div class="header-btns">
            <a class="home-btn" href="../index.html" title="Todos los servicios">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#F8F4E1" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </a>
            <div class="home-btn" onclick="showContent('welcome')" title="Inicio Servicio Forjar Restaurativo">
                <img src="../imagenes/servicios/forjar.png" alt="Servicio Forjar Restaurativo" style="height:32px; border-radius:16px; object-fit:contain; vertical-align:middle;">
            </div>
        </div>
    </header>"""

# =====================================================================
# Sidebar - secciones de navegación
# =====================================================================
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
                    <div class="sidebar-item" onclick="showContent('ubicacion')">Ubicaci&oacute;n</div>
                    <div class="sidebar-item" onclick="showContent('modalidades')">Modalidades de atenci&oacute;n</div>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="showContent('proceso_operativo')" style="cursor:pointer;">
                    <span>Proceso operativo</span>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="toggleSection(this)">
                    <span>Gesti&oacute;n de datos</span><span class="arrow">&#9654;</span>
                </div>
                <div class="sidebar-items">
                    <div class="sidebar-item" onclick="showContent('flujo_datos')">Flujo de gesti&oacute;n de la informaci&oacute;n</div>
                    <div class="sidebar-item" onclick="showContent('datos_sirbe')">Datos SIRBE</div>
                </div>
            </div>
            <div class="sidebar-section">
                <div class="sidebar-title" onclick="showContent('reporte_politicas')" style="cursor:pointer;">
                    <span>Reporte a pol&iacute;ticas</span>
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

# =====================================================================
# Secciones de contenido - cada una identificada por su id
# =====================================================================

SECCION_WELCOME = """\
            <div class="content-section active" id="welcome">
                <div class="welcome-section">
                    <div style="font-family:'Anton','Figtree',sans-serif; font-weight:400; font-size:1.9rem; line-height:1.05; letter-spacing:1px; text-transform:uppercase; background:#2d2a28; color:#f4f5de; padding:14px 24px 11px; margin:0 auto 28px; display:block; width:fit-content; max-width:100%; text-align:center;">Servicio Forjar Restaurativo</div>
                    <p>Servicio de atenci&oacute;n integral, especializada y diferencial dirigido a adolescentes y j&oacute;venes de 14 a 28 a&ntilde;os vinculados al Sistema de Responsabilidad Penal para Adolescentes (SRPA) y a sus redes familiares, en el marco de modalidades de atenci&oacute;n no privativas de la libertad, desde un enfoque pedag&oacute;gico y restaurativo.</p>
                    <div style="margin:30px auto 0; max-width:450px;">
                        <img src="../imagenes/Forjar.jpg" alt="Servicio Forjar Restaurativo" style="width:100%; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.1);">
                    </div>
                </div>
            </div>"""

SECCION_LINEA_TIEMPO = """\
            <div class="content-section" id="linea_tiempo">
                <div class="card">
                    <h2 class="card-title">L&iacute;nea de tiempo</h2>
                    <div class="linea-tiempo" style="--lt-cols: 7;">
                        <article class="lt-hito">
                            <div class="lt-chevron lt-c1"><i data-lucide="handshake"></i>2009</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i1"><i data-lucide="handshake"></i></div>
                                <div class="lt-titulo lt-t1">Creaci&oacute;n como convenio</div>
                                <div class="lt-texto">El servicio nace en Ciudad Bol&iacute;var como respuesta a la &ldquo;delincuencia juvenil&rdquo; en zonas de calor. Convenio entre <strong>Alcald&iacute;a Local, Artesan&iacute;as de Colombia, CIRCO Ciudad, USAID y OIM</strong>, con enfoque en formaci&oacute;n en artes y oficios.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c2"><i data-lucide="landmark"></i>2012&mdash;13</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i2"><i data-lucide="landmark"></i></div>
                                <div class="lt-titulo lt-t2">Consolidaci&oacute;n en SDIS</div>
                                <div class="lt-texto">La SDIS asume la operaci&oacute;n desde un enfoque de inclusi&oacute;n y pol&iacute;tica p&uacute;blica. Se consolida como servicio especializado bajo Subdirecci&oacute;n para la Infancia. Se abren las unidades de <strong>Suba</strong> y <strong>Rafael Uribe Uribe</strong> para cubrir el 100% del distrito.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c3"><i data-lucide="users"></i>2020</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i3"><i data-lucide="users"></i></div>
                                <div class="lt-titulo lt-t3">Llega a Juventud</div>
                                <div class="lt-texto">El servicio es trasladado de la Subdirecci&oacute;n para la Infancia a la Subdirecci&oacute;n para la Juventud, bajo el proyecto de inversi&oacute;n <strong>7740 &ldquo;Generaci&oacute;n J&oacute;venes con Derechos en Bogot&aacute;&rdquo;</strong>.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c4"><i data-lucide="route"></i>2021</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i4"><i data-lucide="route"></i></div>
                                <div class="lt-titulo lt-t4">Nuevas rutas</div>
                                <div class="lt-texto">Se dise&ntilde;a la <strong>Ruta de Oportunidades Juveniles</strong> y la <strong>Ruta Integral de Atenci&oacute;n para J&oacute;venes (RIAJ)</strong> para orientar a quienes est&aacute;n vinculados al Sistema de Responsabilidad Penal para Adolescentes (SRPA).</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c5"><i data-lucide="scale"></i>2023</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i5"><i data-lucide="scale"></i></div>
                                <div class="lt-titulo lt-t5">Marco normativo</div>
                                <div class="lt-texto">Adopci&oacute;n del Anexo T&eacute;cnico de Est&aacute;ndares de Calidad (<strong>Resoluci&oacute;n 0824</strong>, abril 2023) y publicaci&oacute;n del Documento Metodol&oacute;gico (septiembre 2023). Se armonizan los postulados de justicia restaurativa y juvenil con las pol&iacute;ticas distritales.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c6"><i data-lucide="door-open"></i>2025</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i6"><i data-lucide="door-open"></i></div>
                                <div class="lt-titulo lt-t6">Egreso fortalecido</div>
                                <div class="lt-texto">Se amplía la estrategia de acompa&ntilde;amiento en el egreso para brindar continuidad voluntaria por <strong>6 meses a 1 a&ntilde;o</strong> tras cumplir la sanci&oacute;n, consolidando la inclusi&oacute;n social y productiva.</div>
                            </div>
                        </article>

                        <article class="lt-hito">
                            <div class="lt-chevron lt-c7"><i data-lucide="book-open"></i>2026</div>
                            <div class="lt-cuerpo">
                                <div class="lt-icono lt-i7"><i data-lucide="book-open"></i></div>
                                <div class="lt-titulo lt-t7">Manual operativo</div>
                                <div class="lt-texto">Creaci&oacute;n del <strong>Manual Operativo</strong> que reemplaza el Modelo de Atenci&oacute;n Integral. Actualiza referentes conceptuales y metodol&oacute;gicos articulados con las condiciones espec&iacute;ficas de calidad del servicio.</div>
                            </div>
                        </article>
                    </div>
                </div>
            </div>"""

SECCION_A_TENER_EN_CUENTA = """\
            <div class="content-section" id="a_tener_en_cuenta">
                <div class="card">
                    <h2 class="card-title">A tener en cuenta</h2>

                    <div class="atc-grid">
                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/3.png" alt="Mano se&ntilde;alando"></div>
                            <h3 class="atc-titulo atc-titulo-1">Cambio de paradigma en el nombre</h3>
                            <p class="atc-texto">El servicio antes se conoc&iacute;a como <strong>&ldquo;Centros Forjar&rdquo;</strong>. La palabra &ldquo;centro&rdquo; fue eliminada porque los j&oacute;venes la asociaban con centros de reclusi&oacute;n, lo que iba en contra del enfoque restaurativo. La referencia correcta es <strong>servicio Forjar</strong>. Nunca &ldquo;Centro Forjar&rdquo;.</p>
                        </article>

                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/2.png" alt="Mano love-rock"></div>
                            <h3 class="atc-titulo atc-titulo-2">Justicia restaurativa, no retributiva</h3>
                            <p class="atc-texto">El servicio se fundamenta en tres pilares: <strong>responsabilizaci&oacute;n</strong>, <strong>reparaci&oacute;n del da&ntilde;o</strong> e <strong>inclusi&oacute;n social</strong>. Se distancia del modelo de justicia retributiva centrado en el castigo.</p>
                        </article>

                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/3.png" alt="Mano se&ntilde;alando"></div>
                            <h3 class="atc-titulo atc-titulo-3">Atenci&oacute;n a j&oacute;venes mayores de edad</h3>
                            <p class="atc-texto">Forjar atiende hasta los <strong>28 a&ntilde;os</strong> en dos l&iacute;neas: SRPA (cuando el ingreso ocurri&oacute; antes de los 18) y j&oacute;venes que requieren restablecimiento de derechos. M&aacute;s de la mitad de la poblaci&oacute;n atendida es mayor de edad.</p>
                        </article>

                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/5.png" alt="Mano apuntando"></div>
                            <h3 class="atc-titulo atc-titulo-4">Cobertura distrital desde tres puntos</h3>
                            <p class="atc-texto">Opera f&iacute;sicamente en <strong>3 unidades operativas</strong> (Suba, Ciudad Bol&iacute;var y Rafael Uribe Uribe), pero atiende al <strong>100% del Distrito Capital</strong> recibiendo j&oacute;venes de cualquier localidad.</p>
                        </article>

                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/4.png" alt="Mano OK"></div>
                            <h3 class="atc-titulo atc-titulo-5">Continuidad voluntaria tras la sanci&oacute;n</h3>
                            <p class="atc-texto">El acompa&ntilde;amiento no termina con la orden del juez. Los j&oacute;venes pueden continuar de manera <strong>100% voluntaria por 6 meses a 1 a&ntilde;o</strong> para consolidar su proyecto de vida y evitar el retorno a entornos vulnerables.</p>
                        </article>

                        <article class="atc-card">
                            <div class="atc-mano"><img src="../imagenes/manos/6.png" alt="Pu&ntilde;o cerrado"></div>
                            <h3 class="atc-titulo atc-titulo-6">Cumplimiento de mandatos judiciales</h3>
                            <p class="atc-texto">A diferencia de otros servicios, el cumplimiento en Forjar <strong>no es voluntario</strong>: responde a imposiciones de autoridades del SRPA. La prestaci&oacute;n no puede interrumpirse porque cualquier incumplimiento puede generar consecuencias legales y disciplinarias.</p>
                        </article>
                    </div>
                </div>
            </div>"""

SECCION_EQUIPO = """\
            <div class="content-section" id="equipo">
                <div class="card">
                    <h2 class="card-title">Equipo</h2>

                    <h3 class="card-subtitle">Liderazgo</h3>
                    <p style="line-height:1.7;"><strong>Aura Vanessa Le&oacute;n</strong> - L&iacute;der del Servicio Forjar Restaurativo, Subdirecci&oacute;n para la Juventud, SDIS.</p>
                    <p style="line-height:1.7;">Coordina la atenci&oacute;n especializada, formaci&oacute;n para el trabajo y oportunidades juveniles. El servicio hace parte de la estructura territorial de la SDIS, reportando al Subdirector para la Juventud.</p>

                    <h3 class="card-subtitle">Componentes del equipo de atenci&oacute;n integral</h3>
                    <p style="line-height:1.7; margin-bottom:18px;">El equipo se organiza en ocho componentes que articulan los roles profesionales de la atenci&oacute;n integral. La agrupaci&oacute;n por color identifica el tipo de aporte: <span style="color:#5f9ea0; font-weight:600;">atenci&oacute;n directa al joven</span>, <span style="color:#e07850; font-weight:600;">apoyo especializado</span> y <span style="color:#7b6b99; font-weight:600;">soporte t&eacute;cnico y administrativo</span>.</p>

                    <style>
                        .equipo-diagrama-wrapper { position: relative; max-width: 980px; margin: 18px auto 28px; padding: 8px 0; }
                        .equipo-anillo { position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; z-index: 0; }
                        .equipo-diagrama { position: relative; display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; grid-template-rows: repeat(4, auto); gap: 14px 18px; align-items: center; z-index: 1; }
                        .equipo-caja { border: 2px solid; border-radius: 14px; padding: 12px 14px; background: #fff; box-shadow: 0 2px 6px rgba(0,0,0,0.04); transition: transform 0.18s ease, box-shadow 0.18s ease; position: relative; z-index: 1; }
                        .equipo-caja:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
                        .equipo-caja .titulo { font-family: 'Anton','Figtree',sans-serif; font-weight: 400; letter-spacing: 0.5px; text-transform: uppercase; font-size: 0.92rem; line-height: 1.1; margin-bottom: 4px; }
                        .equipo-caja .roles { font-size: 0.82rem; color: #4a4a4a; line-height: 1.45; }
                        .equipo-caja.izq { text-align: right; }
                        .equipo-caja.der { text-align: left; }
                        .equipo-color-1 { border-color: #5f9ea0; } .equipo-color-1 .titulo { color: #5f9ea0; }
                        .equipo-color-2 { border-color: #e07850; } .equipo-color-2 .titulo { color: #e07850; }
                        .equipo-color-3 { border-color: #7b6b99; } .equipo-color-3 .titulo { color: #7b6b99; }
                        .equipo-centro { grid-column: 2 / span 2; grid-row: 2 / span 2; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px; z-index: 2; position: relative; }
                        .equipo-centro-circulo { width: 170px; height: 170px; border-radius: 50%; background: #90c8c6; border: 4px solid #5f9ea0; display: flex; align-items: center; justify-content: center; box-shadow: 0 6px 18px rgba(0,0,0,0.14); }
                        .equipo-centro-circulo img { width: 140px; height: auto; }
                        .equipo-centro-label { font-family: 'Anton','Figtree',sans-serif; font-weight: 400; letter-spacing: 0.6px; text-transform: uppercase; font-size: 0.85rem; color: #2F3E3C; text-align: center; max-width: 180px; line-height: 1.15; }
                        @media (max-width: 960px) {
                            .equipo-anillo { display: none; }
                            .equipo-diagrama { grid-template-columns: 1fr 1fr; grid-template-rows: auto; gap: 10px; align-items: start; }
                            .equipo-caja { grid-column: auto !important; grid-row: auto !important; }
                            .equipo-caja.izq, .equipo-caja.der { text-align: center; }
                            .equipo-centro { grid-column: 1 / -1 !important; grid-row: auto !important; order: -1; margin-bottom: 14px; }
                            .equipo-centro-circulo { width: 140px; height: 140px; }
                            .equipo-centro-circulo img { width: 105px; }
                        }
                        @media (max-width: 460px) {
                            .equipo-caja { padding: 9px 10px; }
                            .equipo-caja .titulo { font-size: 0.78rem; letter-spacing: 0.3px; }
                            .equipo-caja .roles { font-size: 0.75rem; line-height: 1.4; }
                        }
                    </style>

                    <div class="equipo-diagrama-wrapper">
                        <svg class="equipo-anillo" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
                            <ellipse cx="50" cy="50" rx="46" ry="44" fill="none" stroke="#5f9ea0" stroke-width="0.45" stroke-dasharray="1.2 1.2" opacity="0.6"/>
                        </svg>
                        <div class="equipo-diagrama">
<!--CAJAS_EQUIPO-->
                            <div class="equipo-centro">
                                <div class="equipo-centro-circulo">
                                    <img src="../imagenes/servicios/forjar_circulo.png" alt="Servicio Forjar Restaurativo">
                                </div>
                                <div class="equipo-centro-label">Equipo de atenci&oacute;n integral</div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>"""

# Generar cajas del equipo desde Excel
equipo_excel = os.path.join(DATOS, "equipo.xlsx")  # hoja "forjar" del libro unico
if os.path.exists(equipo_excel):
    df_equipo = pd.read_excel(equipo_excel, sheet_name="forjar")
    # Posicion -> (grid-column, grid-row, alineacion). Layout octogonal:
    # filas 1 y 4 ocupan columnas centrales (2,3); filas 2 y 3 ocupan columnas
    # exteriores (1,4); el logo central abarca filas 2-3 y columnas 2-3.
    POSICION_GRID = {
        1: (2, 1, "izq"),  # top-izq al centro
        2: (3, 1, "der"),  # top-der al centro
        3: (1, 2, "izq"),  # mid-alta-izq
        4: (4, 2, "der"),  # mid-alta-der
        5: (1, 3, "izq"),  # mid-baja-izq
        6: (4, 3, "der"),  # mid-baja-der
        7: (2, 4, "izq"),  # bot-izq al centro
        8: (3, 4, "der"),  # bot-der al centro
    }
    COLOR_GRUPO = {
        "atencion_directa": "equipo-color-1",
        "apoyo_especializado": "equipo-color-2",
        "soporte": "equipo-color-3",
    }
    cajas_html = ""
    for pos in sorted(df_equipo["Posicion"].unique()):
        sub = df_equipo[df_equipo["Posicion"] == pos]
        componente = str(sub.iloc[0]["Componente"]).strip()
        grupo = str(sub.iloc[0]["Grupo"]).strip()
        # Concatenar roles con guion corto rodeado de espacios
        roles = " - ".join(str(r).strip() for r in sub["Rol"].tolist() if pd.notna(r))
        col, row, align = POSICION_GRID.get(int(pos), (1, 1, "izq"))
        color_class = COLOR_GRUPO.get(grupo, "equipo-color-1")
        cajas_html += f'                            <div class="equipo-caja {align} {color_class}" style="grid-column:{col}; grid-row:{row};">\n'
        cajas_html += f'                                <div class="titulo">{componente}</div>\n'
        cajas_html += f'                                <div class="roles">{roles}</div>\n'
        cajas_html += f'                            </div>\n'
    SECCION_EQUIPO = SECCION_EQUIPO.replace("<!--CAJAS_EQUIPO-->", cajas_html.rstrip())
    print(f"Equipo Forjar generado desde Excel: {len(df_equipo)} roles en {df_equipo['Posicion'].nunique()} componentes")
else:
    SECCION_EQUIPO = SECCION_EQUIPO.replace("<!--CAJAS_EQUIPO-->", "")
    print("Excel de equipo Forjar no encontrado en datos/equipo.xlsx (hoja forjar)")

SECCION_UBICACION = """\
            <div class="content-section" id="ubicacion">
                <div class="card">
                    <h2 class="card-title">Ubicaci&oacute;n</h2>
                    <p style="line-height:1.7;">El servicio Forjar Restaurativo opera en <strong>3 unidades operativas</strong> ubicadas en localidades con alta incidencia, con cobertura para el 100% del Distrito Capital.</p>
                    <p style="line-height:1.7;">Aunque opera desde estos puntos, el servicio recibe j&oacute;venes de <strong>cualquier localidad</strong> de Bogot&aacute;.</p>

                    <div style="margin:20px 0;">
                        <iframe src="mapa_forjar.html" width="100%" height="450" style="border:1px solid #e0e0e0; border-radius:8px;" loading="lazy"></iframe>
                    </div>

                    <div style="margin-top:25px;">
                <h3 class="card-subtitle">Directorio</h3>
                <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
                    <thead>
                        <tr style="background:#2F3E3C; color:#F8F4E1;">
                            <th style="padding:12px 14px; text-align:left; font-weight:700;">Unidad operativa</th>
                            <th style="padding:12px 14px; text-align:left; font-weight:700;">Localidad</th>
                            <th style="padding:12px 14px; text-align:left; font-weight:700;">Direcci&oacute;n</th>
                        </tr>
                    </thead>
                    <tbody>
                        <!--FILAS_DIRECTORIO-->
                    </tbody>
                </table>
            </div>
                </div>
            </div>"""

# Generar filas del directorio leyendo el Excel
directorio_excel = os.path.join(DATOS, "directorio_forjar.xlsx")
if os.path.exists(directorio_excel):
    df_dir = pd.read_excel(directorio_excel)
    filas_html = ""
    for idx, row in df_dir.iterrows():
        bg = '#fafafa' if idx % 2 == 0 else '#fff'
        link_maps = row.get("Link Google Maps", "")
        direccion = str(row["Dirección"])
        if pd.notna(link_maps) and str(link_maps).strip():
            direccion_html = f'<a href="{link_maps}" target="_blank" style="color:var(--accent);">{direccion}</a>'
        else:
            direccion_html = direccion
        filas_html += f'                        <tr style="background:{bg}; border-bottom:1px solid #e0e0e0;">\n'
        filas_html += f'                            <td style="padding:12px 14px; vertical-align:top;"><strong>{row["Nombre unidad operativa"]}</strong></td>\n'
        filas_html += f'                            <td style="padding:12px 14px; vertical-align:top;">{row["Localidad"]}</td>\n'
        filas_html += f'                            <td style="padding:12px 14px; vertical-align:top;">{direccion_html}</td>\n'
        filas_html += f'                        </tr>\n'
    SECCION_UBICACION = SECCION_UBICACION.replace("<!--FILAS_DIRECTORIO-->", filas_html.rstrip())
    print(f"Directorio Forjar generado desde Excel: {len(df_dir)} unidades")

    # Generar mapa Folium con localidades + marcadores de unidades operativas
    import folium
    import json as _json
    import unicodedata

    def _normalizar_loc(texto):
        texto = unicodedata.normalize('NFC', texto).upper().strip()
        texto = texto.replace('Ñ', 'N').replace('ñ', 'n')
        texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
        texto = re.sub(r'^(LA |LOS |EL )', '', texto)
        return texto

    geojson_path = os.path.join(DATOS, "localidades_bogota.geojson")
    # Centro y zoom FIJOS (no fit_bounds): este mapa vive en una pestana
    # oculta y, al cargarse sin tamano, fit_bounds hace que Leaflet se vaya
    # al globo. [4.28, -74.2] zoom 9 encuadra toda Bogota incluida Sumapaz.
    m = folium.Map(location=[4.28, -74.2], zoom_start=9, tiles="CartoDB positron", width="100%", height="100%")

    if os.path.exists(geojson_path):
        with open(geojson_path, encoding="utf-8") as f:
            localidades_gj = _json.load(f)
        locs_forjar = set(_normalizar_loc(l) for l in df_dir["Localidad"].unique())

        def _style_loc(feature):
            nombre = _normalizar_loc(feature["properties"]["nombre"])
            if nombre in locs_forjar:
                return {"fillColor": "#d5e8e8", "color": "#5f9ea0", "weight": 2, "fillOpacity": 0.4}
            else:
                # Localidades sin unidad (incluida Sumapaz): se dibujan igual,
                # con borde gris visible, para ver el croquis completo de Bogota.
                return {"fillColor": "#ececec", "color": "#7a7a7a", "weight": 1.4, "fillOpacity": 0.45}

        folium.GeoJson(localidades_gj, name="Localidades", style_function=_style_loc,
            tooltip=folium.GeoJsonTooltip(fields=["nombre"], aliases=["Localidad:"])).add_to(m)

    for _, row in df_dir.iterrows():
        lat, lon = row.get("Latitud"), row.get("Longitud")
        if pd.notna(lat) and pd.notna(lon):
            popup_html = f'<div style="font-family:Arial; min-width:200px;"><strong style="color:#5f9ea0; font-size:14px;">{row["Nombre unidad operativa"]}</strong><br><span style="color:#666; font-size:12px;">{row["Localidad"]}</span><br><span style="font-size:11px;">{row["Dirección"]}</span></div>'
            folium.CircleMarker(
                location=[lat, lon], radius=8, color="#5f9ea0",
                fill=True, fill_color="#5f9ea0", fill_opacity=0.9, weight=2,
                popup=folium.Popup(popup_html, max_width=250),
                tooltip=row["Nombre unidad operativa"]
            ).add_to(m)

    m.save(os.path.join(BASE, "html", "mapa_forjar.html"))
    print(f"Mapa Forjar generado: mapa_forjar.html")
else:
    print("Excel de directorio Forjar no encontrado")

SECCION_MODALIDADES = """\
            <div class="content-section" id="modalidades">
                <div class="card">
                    <h2 class="card-title">Modalidades de atenci&oacute;n</h2>
                    <p style="margin-bottom:8px;">El servicio Forjar Restaurativo da cumplimiento a sanciones no privativas de la libertad en medio abierto y comunitario. Tambi&eacute;n recibe j&oacute;venes para medidas de restablecimiento de derechos y atiende a quienes contin&uacute;an voluntariamente tras cumplir su sanci&oacute;n.</p>
                    <p style="color:#888; font-size:0.88rem; margin-bottom:28px;">Las modalidades se clasifican seg&uacute;n el tipo de medida, sanci&oacute;n o estrategia impuesta por la autoridad competente.</p>

                    <!-- Categor&iacute;a 1: Sanciones no privativas -->
                    <div style="background:#5f9ea0; color:#fff; padding:12px 20px; border-radius:8px 8px 0 0; font-weight:700; font-size:0.95rem;">1. No privativas de la libertad (sanciones)</div>
                    <div style="background:#F8F4E1; padding:18px 22px 20px; border-radius:0 0 8px 8px; margin-bottom:32px;">
                        <p style="font-size:0.85rem; color:#555; margin:0 0 18px;">Se cumplen en medio abierto, sociofamiliar o comunitario. Buscan la responsabilizaci&oacute;n, reparaci&oacute;n del da&ntilde;o e inclusi&oacute;n social.</p>

                        <div style="margin-bottom:18px;">
                            <h4 style="font-size:1rem; font-weight:700; color:#2F3E3C; margin:0 0 8px; line-height:1.4;"><span style="color:#5f9ea0; margin-right:10px; font-size:1.2rem;">&bull;</span>Libertad asistida y/o vigilada</h4>
                            <div style="padding-left:24px;">
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0 0 10px;">Concesi&oacute;n de libertad otorgada por la autoridad judicial con la condici&oacute;n de que el adolescente o joven se someta a supervisi&oacute;n, asistencia y orientaci&oacute;n de un programa especializado. Busca fortalecer la autonom&iacute;a y la reparaci&oacute;n a trav&eacute;s de espacios pedag&oacute;gicos y pr&aacute;cticas restaurativas. <strong>M&iacute;nimo 10 actividades mensuales.</strong></p>
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0;">Para garantizar la efectividad de esta medida, estas intervenciones mensuales se distribuyen estrat&eacute;gicamente a nivel individual, familiar, grupal y en el contexto del joven. Con este abordaje integral, el programa involucra activamente a la familia o red vincular de apoyo para que lo acompa&ntilde;en en el reconocimiento de su responsabilidad, facilitando procesos de reparaci&oacute;n (directa, indirecta o simb&oacute;lica) hacia las personas afectadas y logrando la resignificaci&oacute;n de su proyecto vital en el marco de la cultura de la legalidad.</p>
                            </div>
                        </div>

                        <div>
                            <h4 style="font-size:1rem; font-weight:700; color:#2F3E3C; margin:0 0 8px; line-height:1.4;"><span style="color:#5f9ea0; margin-right:10px; font-size:1.2rem;">&bull;</span>Prestaci&oacute;n de servicios a la comunidad</h4>
                            <div style="padding-left:24px;">
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0 0 10px;">La prestaci&oacute;n de servicios a la comunidad es una sanci&oacute;n judicial que consiste en la ejecuci&oacute;n de acciones o actividades no remuneradas de utilidad social por parte del adolescente o joven. <strong>M&iacute;nimo 32 horas mensuales:</strong> 24 horas de trabajo comunitario y 8 horas de acompa&ntilde;amiento psicosocial.</p>
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0;">Esta sanci&oacute;n tiene una finalidad restaurativa y formativa, ya que busca que el participante reflexione de forma cr&iacute;tica sobre las consecuencias de su conducta punible. Las acciones ejecutadas posibilitan la reparaci&oacute;n directa, indirecta o simb&oacute;lica a la v&iacute;ctima, a la familia y a la comunidad, facilitando as&iacute; la restauraci&oacute;n de los lazos de confianza afectados, la reconstrucci&oacute;n del tejido social y el fortalecimiento del sentido de ciudadan&iacute;a.</p>
                            </div>
                        </div>
                    </div>

                    <!-- Categor&iacute;a 2: Apoyo y restablecimiento -->
                    <div style="background:#e07850; color:#fff; padding:12px 20px; border-radius:8px 8px 0 0; font-weight:700; font-size:0.95rem;">2. Apoyo y restablecimiento de derechos</div>
                    <div style="background:#F8F4E1; padding:18px 22px 20px; border-radius:0 0 8px 8px; margin-bottom:32px;">
                        <p style="font-size:0.85rem; color:#555; margin:0 0 18px;">Medidas complementarias frente a vulneraciones de derechos.</p>

                        <div style="margin-bottom:18px;">
                            <h4 style="font-size:1rem; font-weight:700; color:#2F3E3C; margin:0 0 8px; line-height:1.4;"><span style="color:#e07850; margin-right:10px; font-size:1.2rem;">&bull;</span>IARAJ</h4>
                            <div style="padding-left:24px;">
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0 0 10px;">La <strong>Intervenci&oacute;n de Apoyo al Restablecimiento en Administraci&oacute;n de Justicia (IARAJ)</strong> es una modalidad que se aplica como medida complementaria, cuyo prop&oacute;sito central es el restablecimiento y la garant&iacute;a de los derechos de los adolescentes y j&oacute;venes bajo el acompa&ntilde;amiento y orientaci&oacute;n de su red familiar. Su implementaci&oacute;n operativa requiere un cumplimiento <strong>m&iacute;nimo de diez (10) actividades mensuales</strong>, desarrolladas mediante intervenciones individuales, grupales, familiares y contextuales, asegurando un abordaje sist&eacute;mico frente a la problem&aacute;tica.</p>
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0;">A trav&eacute;s de este servicio en medio abierto, se ofrece apoyo pedag&oacute;gico, psicol&oacute;gico y psicosocial especializado para que los participantes se vinculen a escenarios de socializaci&oacute;n positiva. La finalidad es brindarles herramientas que les permitan comprender y responsabilizarse de los hechos de riesgo en los que se han involucrado, mitigando as&iacute; los factores de vulnerabilidad social en su entorno.</p>
                            </div>
                        </div>

                        <div>
                            <h4 style="font-size:1rem; font-weight:700; color:#2F3E3C; margin:0 0 8px; line-height:1.4;"><span style="color:#e07850; margin-right:10px; font-size:1.2rem;">&bull;</span>RIAJ</h4>
                            <div style="padding-left:24px;">
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0 0 10px;">La <strong>Ruta Integral de Atenci&oacute;n para J&oacute;venes (RIAJ)</strong> es una estrategia garantizada por la Secretar&iacute;a Distrital de Integraci&oacute;n Social, dirigida espec&iacute;ficamente a j&oacute;venes (incluso mayores de edad) que se encuentran en conflicto con la ley penal, pero cuya situaci&oacute;n jur&iacute;dica a&uacute;n no ha sido definida por la autoridad competente. Al igual que la modalidad IARAJ, su modelo de intervenci&oacute;n exige un <strong>m&iacute;nimo de diez (10) actividades mensuales</strong> de atenci&oacute;n y orientaci&oacute;n junto con el joven y su familia.</p>
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0;">El enfoque principal de la estrategia RIAJ es preventivo y protector, operando a trav&eacute;s de la gesti&oacute;n de servicios, la generaci&oacute;n de alertas y la vinculaci&oacute;n a la oferta institucional del Distrito. Su desarrollo busca favorecer el acceso real a oportunidades de inclusi&oacute;n en &aacute;reas fundamentales como el empleo, el emprendimiento, la educaci&oacute;n, el aprovechamiento del tiempo libre, el arte, la cultura y el deporte, evitando de este modo la escalada en conductas delictivas.</p>
                            </div>
                        </div>
                    </div>

                    <!-- Categor&iacute;a 3: Estrategias transversales -->
                    <div style="background:#1eaf76; color:#fff; padding:12px 20px; border-radius:8px 8px 0 0; font-weight:700; font-size:0.95rem;">3. Estrategias transversales de continuidad</div>
                    <div style="background:#F8F4E1; padding:18px 22px 20px; border-radius:0 0 8px 8px;">
                        <div>
                            <h4 style="font-size:1rem; font-weight:700; color:#2F3E3C; margin:0 0 8px; line-height:1.4;"><span style="color:#1eaf76; margin-right:10px; font-size:1.2rem;">&bull;</span>Estrategia de acompa&ntilde;amiento al egreso</h4>
                            <div style="padding-left:24px;">
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0 0 10px;">Dirigida a quienes culminaron el cumplimiento de diversas modalidades del Sistema de Responsabilidad Penal para Adolescentes (SRPA) y deciden dar continuidad voluntaria a su proceso. <strong>Acompa&ntilde;amiento por 6 meses a 1 a&ntilde;o</strong> para consolidar la inclusi&oacute;n social y productiva mediante la Ruta de Oportunidades Juveniles.</p>
                                <p style="font-size:0.9rem; line-height:1.7; color:#3A3A3A; margin:0;">Para materializar este objetivo, el equipo de la Ruta de Oportunidades Juveniles (ROJ) implementa acciones estrat&eacute;gicas de gesti&oacute;n y articulaci&oacute;n intra e interinstitucional para conectar a los j&oacute;venes y a sus familias con la oferta de servicios locales y distritales. Este esfuerzo se centra en garantizar el acceso efectivo a oportunidades de educaci&oacute;n, empleabilidad, emprendimiento, salud, cultura y deporte, respondiendo a los intereses y necesidades particulares de cada participante. El prop&oacute;sito final de esta fase de transici&oacute;n es fomentar la autonom&iacute;a del joven, consolidar un proyecto de vida con sentido en el marco de la legalidad y asegurar una plena integraci&oacute;n social que act&uacute;e como escudo protector frente a la reincidencia.</p>
                            </div>
                        </div>
                    </div>

                </div>
            </div>"""

# =====================================================================
# Proceso operativo: etapas del flujo, documentos y actores.
# Texto y estructura vienen de datos/forjar_proceso_operativo.xlsx (3 hojas).
# Los enlaces de los documentos vienen de enlaces/enlaces.xlsx (hoja "generales") filtrando
# HTML="Forjar" y SECCION que empiece con "Proceso Operativo".
# =====================================================================
SECCION_PROCESO_OPERATIVO = """\
            <div class="content-section" id="proceso_operativo">
                <div class="card">
                    <h2 class="card-title">Proceso operativo</h2>
                    <p style="line-height:1.7;">Proceso operativo de atenci&oacute;n integral para adolescentes y j&oacute;venes vinculados al SRPA. Conoce las etapas, actividades, registros, formularios y actores clave del servicio.</p>

                    <style>
                        .po-stages-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(230px, 1fr)); gap:16px; margin:20px 0 8px; }
                        .po-stage-card { background:#fff; border-radius:12px; border:1.5px solid #e5e0d3; padding:22px 20px; cursor:pointer; transition:transform 0.18s, box-shadow 0.18s; }
                        .po-stage-card:hover { transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,0.05); }
                        .po-stage-num { font-family:'Anton','Figtree',sans-serif; font-weight:400; font-size:2.6rem; line-height:1; margin-bottom:10px; letter-spacing:1px; }
                        .po-stage-title { font-family:'Antonio','Anton','Figtree',sans-serif; font-weight:700; letter-spacing:0.4px; text-transform:uppercase; font-size:1rem; margin-bottom:6px; color:#2F3E3C; }
                        .po-stage-sub { font-size:0.78rem; color:#666; margin-bottom:10px; font-style:italic; line-height:1.35; }
                        .po-stage-desc { font-size:0.85rem; color:#444; line-height:1.55; }

                        .po-detail-panel { display:none; background:#fff; border-radius:12px; border:1.5px solid #e5e0d3; padding:24px; margin:20px 0 28px; }
                        .po-detail-panel.active { display:block; animation:poFadeIn 0.25s ease; }
                        @keyframes poFadeIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
                        .po-detail-panel h3 { font-family:'Antonio','Anton','Figtree',sans-serif; font-weight:700; letter-spacing:0.4px; text-transform:uppercase; font-size:1.1rem; color:var(--accent); margin-bottom:6px; }
                        .po-detail-intro { font-size:0.9rem; color:#666; margin-bottom:16px; line-height:1.6; }

                        .po-activity-table { width:100%; border-collapse:collapse; font-size:0.85rem; margin-top:8px; }
                        .po-activity-table th { background:#2F3E3C; color:#F8F4E1; font-family:'Antonio','Anton','Figtree',sans-serif; font-size:0.9rem; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; padding:12px 14px; text-align:left; }
                        .po-activity-table tbody tr { background:#fff; }
                        .po-activity-table td { padding:12px 14px; border-bottom:1px solid #e0e0e0; vertical-align:top; line-height:1.55; }
                        .po-activity-table tr:last-child td { border-bottom:none; }
                        .po-activity-table tr:hover td { background:#f5f5f5; }
                        .po-table-wrapper { overflow-x:auto; -webkit-overflow-scrolling:touch; margin:8px -4px 0; padding:0 4px 4px; }
                        @media (max-width: 720px) {
                            .po-detail-panel { padding:16px; }
                            .po-activity-table { font-size:0.78rem; min-width:540px; }
                            .po-activity-table th, .po-activity-table td { padding:8px 10px; }
                        }

                        .po-resp-badge { display:inline-block; font-size:0.72rem; font-weight:500; padding:3px 10px; border-radius:100px; white-space:nowrap; }
                        .po-resp-cespa { background:#d4e6e7; color:#26565a; }
                        .po-resp-administrativo { background:#e8eaed; color:#4a5159; }
                        .po-resp-interdisciplinario { background:#a8c8c9; color:#1d4445; }
                        .po-resp-psicosocial { background:#c5c9ce; color:#3a4148; }
                        .po-resp-roj { background:#7ba3a4; color:#F8F4E1; }
                        .po-resp-otro { background:#dfe3e7; color:#3f464d; }

                        .po-section-block { margin-top:36px; }

                        .po-filter-pills { display:flex; flex-wrap:wrap; gap:8px; margin:14px 0 20px; }
                        .po-pill { padding:6px 16px; border-radius:100px; font-size:0.78rem; font-weight:500; cursor:pointer; border:1.5px solid #e5e0d3; background:#fff; color:#666; transition:all 0.18s; }
                        .po-pill:hover { border-color:var(--accent); color:var(--accent); }
                        .po-pill.active { color:#fff; }
                        .po-pill[data-stage="ingreso"].active { background:#f58b53; border-color:#f58b53; }
                        .po-pill[data-stage="permanencia"].active { background:#5f9ea0; border-color:#5f9ea0; }
                        .po-pill[data-stage="egreso"].active { background:#1eaf76; border-color:#1eaf76; }
                        .po-pill[data-stage="rae"].active { background:#1e7895; border-color:#1e7895; }

                        .po-docs-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(280px, 1fr)); gap:12px; }
                        .po-doc-card { background:#fff; border:1.5px solid #e5e0d3; border-radius:12px; padding:22px 20px 18px; display:flex; align-items:flex-start; gap:14px; transition:box-shadow 0.18s, transform 0.18s; }
                        .po-doc-card:hover { box-shadow:0 4px 12px rgba(0,0,0,0.05); transform:translateY(-2px); }
                        /* El color de la etapa vive en el &iacute;cono, no en una banda superior. */
                        .po-doc-icon { width:36px; min-width:36px; height:36px; display:flex; align-items:center; justify-content:center; flex-shrink:0; color:#2F3E3C; }
                        .po-doc-card.po-doc-ingreso .po-doc-icon { color:#f58b53; }
                        .po-doc-card.po-doc-permanencia .po-doc-icon { color:#5f9ea0; }
                        .po-doc-card.po-doc-egreso .po-doc-icon { color:#1eaf76; }
                        .po-doc-card.po-doc-rae .po-doc-icon { color:#1e7895; }
                        .po-doc-icon svg { width:30px; height:30px; display:block; }
                        .po-doc-info { flex:1; min-width:0; }
                        .po-doc-name { font-size:0.92rem; font-weight:600; color:#2F3E3C; margin-bottom:10px; line-height:1.35; }
                        .po-doc-link { display:inline-flex; align-items:center; gap:6px; font-size:0.78rem; font-weight:500; color:var(--accent); text-decoration:none; }
                        .po-doc-link:hover { text-decoration:underline; }

                        .po-actores-grid { display:grid; grid-template-columns:repeat(3, 1fr); gap:18px 16px; padding-bottom:10px; }
                        .po-actor-card { background:#2B2F3A; color:#F8F4E1; border:none; border-radius:12px; padding:22px 22px 24px; box-shadow:7px 7px 0 var(--accent-bg); }
                        .po-actor-icon { width:30px; height:30px; color:#F8F4E1; margin-bottom:14px; }
                        .po-actor-icon svg { width:100%; height:100%; display:block; }
                        .po-actor-name { font-family:'Antonio','Anton','Figtree',sans-serif; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; font-size:0.98rem; margin-bottom:10px; color:#F8F4E1; }
                        .po-actor-desc { font-size:0.83rem; color:#b8c0c4; line-height:1.55; }
                        @media (max-width: 720px) { .po-actores-grid { grid-template-columns:1fr; } }
                    </style>

                    <h3 class="card-subtitle">Flujo operativo del servicio</h3>
                    <p style="font-size:0.9rem; color:#666; margin-bottom:8px;">Haz clic en una etapa para ver el detalle de actividades, registros y responsables.</p>

                    <div class="po-stages-grid">
<!--CARDS_ETAPAS-->
                    </div>
<!--PANELES_ETAPAS-->

                    <div class="po-section-block">
                        <h3 class="card-subtitle">Documentos y formularios</h3>
                        <p style="font-size:0.9rem; color:#666; margin-bottom:8px;">Accede a los formatos, instrumentos y formularios asociados a cada etapa del servicio.</p>

                        <div class="po-filter-pills">
                            <span class="po-pill" data-stage="ingreso" onclick="filterDocs('ingreso', this)">Ingreso</span>
                            <span class="po-pill" data-stage="permanencia" onclick="filterDocs('permanencia', this)">Permanencia</span>
                            <span class="po-pill" data-stage="egreso" onclick="filterDocs('egreso', this)">Egreso</span>
                            <span class="po-pill" data-stage="rae" onclick="filterDocs('rae', this)">Post-Egreso (RAE)</span>
                        </div>

                        <div class="po-docs-grid">
<!--DOCUMENTOS-->
                        </div>
                    </div>

                    <div class="po-section-block">
                        <h3 class="card-subtitle">Actores clave del proceso</h3>
                        <p style="font-size:0.9rem; color:#666; margin-bottom:14px;">Roles y responsabilidades en la prestaci&oacute;n del servicio.</p>

                        <div class="po-actores-grid">
<!--ACTORES-->
                        </div>
                    </div>
                </div>
            </div>"""

# Color visual de cada etapa (no se guarda en Excel para no confundir al equipo).
# Paleta SDIS reutilizada de los modulos de JCO (generar_gc_jco.py) para mantener
# consistencia visual entre los HTMLs del gestor. La asignacion comunica el flujo:
# tonos calidos al ingreso, frios hacia el cierre y egreso.
COLOR_ETAPA = {
    "ingreso": "#f58b53",      # naranja coral (paleta oficial SDIS Juventud)
    "permanencia": "#5f9ea0",  # aguamarina (accent Forjar)
    "egreso": "#1eaf76",       # verde (paleta oficial SDIS Juventud)
    "rae": "#1e7895",          # azul petroleo (paleta oficial SDIS Juventud)
}

# Iconos SVG outline (estilo Lucide) por actor. Si maniana se agrega un actor
# nuevo a la hoja Excel sin entrada aqui, cae en _ICONO_ACTOR_DEFAULT.
_SVG_PROLOGO = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">'
_ICONO_ACTOR_DEFAULT = (
    _SVG_PROLOGO + '<circle cx="12" cy="12" r="9"/></svg>'
)
ICONOS_ACTOR = {
    "Enlace CESPA": (
        _SVG_PROLOGO
        + '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>'
        + '<path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
        + '</svg>'
    ),
    "Equipo Administrativo": (
        _SVG_PROLOGO
        + '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>'
        + '</svg>'
    ),
    "Equipo Interdisciplinario": (
        _SVG_PROLOGO
        + '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>'
        + '<circle cx="9" cy="7" r="4"/>'
        + '<path d="M23 21v-2a4 4 0 0 0-3-3.87"/>'
        + '<path d="M16 3.13a4 4 0 0 1 0 7.75"/>'
        + '</svg>'
    ),
    "Equipo ROJ": (
        _SVG_PROLOGO
        + '<circle cx="6" cy="19" r="3"/>'
        + '<path d="M9 19h8.5a3.5 3.5 0 0 0 0-7h-11a3.5 3.5 0 0 1 0-7H15"/>'
        + '<circle cx="18" cy="5" r="3"/>'
        + '</svg>'
    ),
    "Familia / Referente": (
        _SVG_PROLOGO
        + '<circle cx="5" cy="7" r="2"/>'
        + '<path d="M2 21v-5c0-1.5 1.5-3 3-3s3 1.5 3 3v5"/>'
        + '<circle cx="12" cy="9" r="1.5"/>'
        + '<path d="M10 21v-4c0-1 1-2 2-2s2 1 2 2v4"/>'
        + '<circle cx="19" cy="7" r="2"/>'
        + '<path d="M16 21v-5c0-1.5 1.5-3 3-3s3 1.5 3 3v5"/>'
        + '</svg>'
    ),
    "Autoridades Competentes": (
        _SVG_PROLOGO
        + '<path d="M12 3v18"/>'
        + '<path d="M7 21h10"/>'
        + '<path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>'
        + '<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/>'
        + '<path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/>'
        + '</svg>'
    ),
}


def _grupo_responsable(nombre):
    """Mapea el nombre del responsable a uno de los 5 grupos de color del badge.

    Es deliberadamente tolerante: si maniana cambian el rotulo de un equipo
    (por ejemplo "Equipo Interdisciplinario Ampliado"), sigue cayendo en el
    mismo grupo de color sin que haya que tocar el Excel ni el script.
    """
    n = str(nombre).lower()
    if "cespa" in n:
        return "cespa"
    if "administrativo" in n:
        return "administrativo"
    if "roj" in n:
        return "roj"
    if "interdisciplinario" in n:
        return "interdisciplinario"
    if "psicosocial" in n or "bina" in n:
        return "psicosocial"
    return "otro"


def _normalizar_etapa_id(texto):
    """Convierte el segmento de etapa de la columna SECCION en el id canonico."""
    t = str(texto).lower().strip()
    if "ingreso" in t:
        return "ingreso"
    if "permanencia" in t:
        return "permanencia"
    if "post" in t or "rae" in t:
        return "rae"
    if "egreso" in t:
        return "egreso"
    return None


# Cargar datos del proceso operativo
proceso_excel = os.path.join(DATOS, "forjar_proceso_operativo.xlsx")
if os.path.exists(proceso_excel):
    df_etapas = pd.read_excel(proceso_excel, sheet_name="etapas")
    df_acts = pd.read_excel(proceso_excel, sheet_name="actividades")
    df_actores = pd.read_excel(proceso_excel, sheet_name="actores")

    # Cards de etapas y paneles desplegables con la tabla de actividades
    cards_html = ""
    paneles_html = ""
    for idx, fila in df_etapas.iterrows():
        eid = str(fila["id"]).strip()
        numero = f"{idx + 1:02d}"
        color = COLOR_ETAPA.get(eid, "#999")

        cards_html += f'                        <div class="po-stage-card" onclick="togglePanel(\'po-panel-{eid}\', this)">\n'
        cards_html += f'                            <div class="po-stage-num" style="color:{color};">{numero}</div>\n'
        cards_html += f'                            <div class="po-stage-title">{fila["titulo"]}</div>\n'
        cards_html += f'                            <div class="po-stage-sub">{fila["subtitulo"]}</div>\n'
        cards_html += f'                            <div class="po-stage-desc">{fila["descripcion_breve"]}</div>\n'
        cards_html += f'                        </div>\n'

        actividades_etapa = df_acts[df_acts["etapa_id"].astype(str).str.strip() == eid]
        filas_act = ""
        for _, act in actividades_etapa.iterrows():
            grupo = _grupo_responsable(act["responsable"])
            filas_act += f'                                <tr>\n'
            filas_act += f'                                    <td>{act["actividad"]}</td>\n'
            filas_act += f'                                    <td>{act["registro"]}</td>\n'
            filas_act += f'                                    <td><span class="po-resp-badge po-resp-{grupo}">{act["responsable"]}</span></td>\n'
            filas_act += f'                                </tr>\n'

        paneles_html += f'                    <div class="po-detail-panel" id="po-panel-{eid}">\n'
        paneles_html += f'                        <h3>Etapa {numero} &mdash; {fila["titulo"]}</h3>\n'
        paneles_html += f'                        <p class="po-detail-intro">{fila["descripcion_panel"]}</p>\n'
        paneles_html += f'                        <div class="po-table-wrapper">\n'
        paneles_html += f'                            <table class="po-activity-table">\n'
        paneles_html += f'                                <thead><tr><th>Actividad</th><th>Registro</th><th>Responsable</th></tr></thead>\n'
        paneles_html += f'                                <tbody>\n{filas_act.rstrip()}\n                                </tbody>\n'
        paneles_html += f'                            </table>\n'
        paneles_html += f'                        </div>\n'
        paneles_html += f'                    </div>\n'

    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--CARDS_ETAPAS-->", cards_html.rstrip())
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--PANELES_ETAPAS-->", paneles_html.rstrip())

    # Cards de actores (3x2). Sombra plana en teal pastel Forjar, uniforme
    # para evitar saturacion cromatica con el resto de la pestania.
    actores_html = ""
    for _, act in df_actores.iterrows():
        nombre = str(act["nombre"]).strip()
        icono = ICONOS_ACTOR.get(nombre, _ICONO_ACTOR_DEFAULT)
        actores_html += f'                            <div class="po-actor-card">\n'
        actores_html += f'                                <div class="po-actor-icon">{icono}</div>\n'
        actores_html += f'                                <div class="po-actor-name">{nombre}</div>\n'
        actores_html += f'                                <div class="po-actor-desc">{act["descripcion"]}</div>\n'
        actores_html += f'                            </div>\n'
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--ACTORES-->", actores_html.rstrip())

    print(f"Proceso operativo Forjar: {len(df_etapas)} etapas, {len(df_acts)} actividades, {len(df_actores)} actores")
else:
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--CARDS_ETAPAS-->", "")
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--PANELES_ETAPAS-->", "")
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--ACTORES-->", "")
    print("Excel forjar_proceso_operativo.xlsx no encontrado en datos/")

# Cargar enlaces de documentos desde enlaces/enlaces.xlsx
enlaces_excel = os.path.join(BASE, "enlaces", "enlaces.xlsx")
if os.path.exists(enlaces_excel):
    df_enlaces = pd.read_excel(enlaces_excel, sheet_name="generales")
    mask = (
        df_enlaces["HTML"].astype(str).str.strip().str.lower() == "forjar"
    ) & (
        df_enlaces["SECCION"].astype(str).str.strip().str.lower().str.startswith("proceso operativo")
    )
    df_docs = df_enlaces[mask].copy()

    docs_html = ""
    docs_validos = 0
    for _, fila in df_docs.iterrows():
        seccion = str(fila["SECCION"]).strip()
        url = str(fila["ENLACE"]).strip()
        # Estructura esperada: "Proceso Operativo - <Etapa> - <Nombre del documento>"
        partes = [p.strip() for p in seccion.split(" - ", 2)]
        if len(partes) < 3:
            print(f"  Aviso: enlace ignorado, formato inesperado en SECCION: {seccion!r}")
            continue
        etapa_id = _normalizar_etapa_id(partes[1])
        if etapa_id is None:
            print(f"  Aviso: enlace ignorado, etapa no reconocida: {partes[1]!r}")
            continue
        nombre = partes[2]

        # Tipo deducido de la URL: forms.office -> formulario (icono documento +
        # verbo "Rellenar"); cualquier otro -> descarga (icono flecha-a-bandeja
        # + verbo "Descargar"). Asi el icono refuerza la accion del enlace.
        es_formulario = "forms.office.com" in url.lower()
        if es_formulario:
            verbo = "Rellenar"
            icono_svg = (
                '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" '
                'stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">'
                '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
                '<polyline points="14 2 14 8 20 8"/>'
                '</svg>'
            )
        else:
            verbo = "Descargar"
            icono_svg = (
                '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" '
                'stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">'
                '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
                '<polyline points="7 10 12 15 17 10"/>'
                '<line x1="12" y1="15" x2="12" y2="3"/>'
                '</svg>'
            )

        # Si el nombre incluye un codigo entre parentesis al final tipo (FOR-PSS-476),
        # lo agregamos al texto del enlace.
        m = re.search(r"\(([A-Z]{2,4}-[A-Z]{2,4}-\d+)\)\s*$", nombre)
        codigo = m.group(1) if m else None
        link_text = f"{verbo} {codigo}" if codigo else verbo

        docs_html += f'                            <div class="po-doc-card po-doc-{etapa_id}" data-stage="{etapa_id}">\n'
        docs_html += f'                                <div class="po-doc-icon">{icono_svg}</div>\n'
        docs_html += f'                                <div class="po-doc-info">\n'
        docs_html += f'                                    <div class="po-doc-name">{nombre}</div>\n'
        docs_html += f'                                    <a href="{url}" target="_blank" class="po-doc-link">{link_text}</a>\n'
        docs_html += f'                                </div>\n'
        docs_html += f'                            </div>\n'
        docs_validos += 1

    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--DOCUMENTOS-->", docs_html.rstrip())
    print(f"Documentos Forjar desde enlaces.xlsx: {docs_validos} enlaces")
else:
    SECCION_PROCESO_OPERATIVO = SECCION_PROCESO_OPERATIVO.replace("<!--DOCUMENTOS-->", "")
    print("enlaces/enlaces.xlsx no encontrado")

SECCION_FLUJO_DATOS = """\
            <div class="content-section" id="flujo_datos">
                <div class="card">
                    <h2 class="card-title">Flujo de gesti&oacute;n de la informaci&oacute;n</h2>
                    <p style="color:#666; margin-bottom:20px;">Proceso de recolecci&oacute;n, registro y reporte de datos en el servicio Forjar Restaurativo.</p>

                    <div class="flujo-pasos">
                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="user-plus"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 01</p>
                                <h3 class="flujo-titulo">Ingreso y acogida</h3>
                                <p class="flujo-responsable">Equipo psicosocial</p>
                                <p class="flujo-texto">Durante el ingreso, el equipo psicosocial (psicolog&iacute;a y trabajo social) realiza la acogida mediante entrevista. Se generan dos registros: el diligenciamiento de la <strong>ficha SIRBE</strong> y la toma de datos en la <strong>Valija Estandarizada</strong>. Se capturan variables b&aacute;sicas (edad, sexo, educaci&oacute;n), transversales (pertenencia &eacute;tnica, discapacidad), de ubicaci&oacute;n geogr&aacute;fica, din&aacute;micas familiares y variables espec&iacute;ficas sobre la remisi&oacute;n judicial y el delito.</p>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="clipboard-list"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 02</p>
                                <h3 class="flujo-titulo">Instrumentos de captura de informaci&oacute;n</h3>
                                <p class="flujo-responsable">Equipo psicosocial</p>
                                <p class="flujo-texto"><strong>Fichas SIRBE:</strong> la informaci&oacute;n se consigna en dos formatos estandarizados de la SDIS: el formato de informaci&oacute;n b&aacute;sica y transversal (FOR-PSS-321) y el formato de variables espec&iacute;ficas del servicio Forjar (FOR-PSS-694).</p>
                                <p class="flujo-texto"><strong>Valija Estandarizada:</strong> base de datos en Excel utilizada internamente por cada unidad operativa. Recopila datos b&aacute;sicos y transversales, asigna un n&uacute;mero de &ldquo;Historia Social&rdquo; para el archivo f&iacute;sico e incluye variables de contraste para verificar la consistencia y calidad de la informaci&oacute;n antes de subirla al sistema oficial.</p>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="monitor"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 03</p>
                                <h3 class="flujo-titulo">Digitalizaci&oacute;n en SIRBE WEB</h3>
                                <p class="flujo-responsable">Enlace CESPA &middot; Equipo psicosocial</p>
                                <p class="flujo-texto">Una vez consolidada la informaci&oacute;n en fichas f&iacute;sicas y Valija Estandarizada, se realiza el cargue al aplicativo <strong>SIRBE WEB</strong> (Sistema de Informaci&oacute;n Misional). All&iacute; queda el estado actualizado de cada usuario: ingresos, seguimientos, incumplimientos de sanci&oacute;n y egresos.</p>
                            </div>
                        </div>

                        <div class="flujo-paso">
                            <div class="flujo-icono"><i data-lucide="shield-check"></i></div>
                            <div>
                                <p class="flujo-orden">Paso 04</p>
                                <h3 class="flujo-titulo">Controles de calidad y reportes</h3>
                                <p class="flujo-responsable">Referente t&eacute;cnico</p>
                                <p class="flujo-texto">El proceso est&aacute; regulado por manuales que definen los roles. Los referentes t&eacute;cnicos revisan mensualmente los datos para ajustes previos a los reportes.</p>
                                <p class="flujo-texto">El tratamiento de datos personales es estrictamente confidencial y se rige por la <strong>Ley Estatutaria 1581 de 2012</strong>.</p>
                                <p class="flujo-texto">La informaci&oacute;n depurada sirve como insumo para preparar mensualmente el <strong>Informe Cualitativo y Cuantitativo del Servicio Forjar Restaurativo</strong>, que visibiliza el volumen de atenciones, cumplimiento de metas y avance del modelo.</p>
                            </div>
                        </div>
                    </div>

                    <h3 class="card-subtitle" style="margin-top:30px;">Diagrama de flujo del proceso</h3>
                    <p style="color:#666; font-size:0.85rem; margin-bottom:12px;">Representaci&oacute;n visual del ciclo de recolecci&oacute;n y digitaci&oacute;n en SIRBE para el servicio Forjar Restaurativo.</p>
                    <div style="overflow-x:auto; margin-bottom:8px;">%%SVG_DIAGRAMA_FORJAR%%</div>
                </div>
            </div>"""

SECCION_DATOS_SIRBE = """\
            <div class="content-section" id="datos_sirbe">
                <div class="card">
                    <h2 class="card-title">Datos SIRBE
                        <span onclick="document.getElementById('sirbe-tooltip-forjar').style.display=document.getElementById('sirbe-tooltip-forjar').style.display==='block'?'none':'block'" style="display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:50%; background:var(--accent); color:#fff; font-size:0.7rem; font-weight:700; cursor:pointer; margin-left:8px; vertical-align:middle;">?</span>
                    </h2>
                    <div id="sirbe-tooltip-forjar" style="display:none; background:#2F3E3C; color:#F8F4E1; border-radius:10px; padding:18px 22px; margin-bottom:18px; font-size:0.88rem; line-height:1.7; position:relative;">
                        <strong>SIRBE</strong> (Sistema de Informaci&oacute;n para el Registro de Beneficiarios) es el aplicativo de datos de la Secretar&iacute;a Distrital de Integraci&oacute;n Social (SDIS) en Bogot&aacute;. Su funci&oacute;n principal es registrar, sistematizar y hacer seguimiento a la informaci&oacute;n de los ciudadanos que acceden a los servicios sociales, proyectos y ayudas de la entidad. La informaci&oacute;n registrada en las &ldquo;Fichas SIRBE&rdquo; es confidencial y su uso se limita a la gesti&oacute;n interna de servicios sociales.
                        <span onclick="document.getElementById('sirbe-tooltip-forjar').style.display='none'" style="position:absolute; top:8px; right:12px; cursor:pointer; opacity:0.6; font-size:1.1rem;">&times;</span>
                    </div>
                    <p style="color:#666; margin-bottom:20px;">C&oacute;mo se estructura el servicio Forjar Restaurativo dentro del sistema de informaci&oacute;n misional SIRBE.</p>

                    <h3 class="card-subtitle">Tipolog&iacute;a en SIRBE</h3>
                    <p style="line-height:1.7;">Forjar funciona en SIRBE bajo la tipolog&iacute;a de <strong>&ldquo;servicio social&rdquo;</strong>.</p>

                    <h3 class="card-subtitle">Entrada al servicio</h3>
                    <p style="line-height:1.7;">Los j&oacute;venes ingresan remitidos por juzgados o por el ICBF, llegan con una historia social, y el primer paso siempre es diligenciar la ficha SIRBE f&iacute;sica espec&iacute;fica para registrar sus datos transversales.</p>

                    <h3 class="card-subtitle">Seguimiento a largo plazo</h3>
                    <p style="line-height:1.7;">Debido a que los j&oacute;venes ingresan para cumplir una sanci&oacute;n, su permanencia en el servicio toma un tiempo prudente. Las modalidades y actuaciones est&aacute;n dise&ntilde;adas para que los j&oacute;venes sigan una ruta, de modo que SIRBE registra todo su historial de movimientos.</p>

                    <h3 class="card-subtitle">Actuaciones adaptadas al sistema penal</h3>
                    <p style="line-height:1.7;">A diferencia de otros servicios, Forjar clasifica los movimientos de los j&oacute;venes en el sistema mediante tres tipos de actuaciones:</p>

                    <p style="line-height:1.7; margin-top:14px;"><strong>Actuaciones de estado:</strong></p>
                    <ul style="margin:6px 0 15px 20px; line-height:2;">
                        <li><span class="badge badge-primary">En atenci&oacute;n</span> El joven se encuentra cumpliendo su sanci&oacute;n o medida activamente.</li>
                        <li><span class="badge badge-primary">Egresado</span> El joven ha finalizado el cumplimiento de su sanci&oacute;n.</li>
                        <li><span class="badge badge-primary">En incumplimiento</span> Funciona en la pr&aacute;ctica como un estado &ldquo;suspendido&rdquo;, pero se cre&oacute; con este nombre espec&iacute;fico para no generar confusiones con la terminolog&iacute;a oficial del Sistema de Responsabilidad Penal para Adolescentes (SRPA).</li>
                    </ul>

                    <p style="line-height:1.7;"><strong>Actuaciones de seguimiento:</strong> Sirven para reportar los traslados que se hacen de los j&oacute;venes entre las diferentes unidades operativas o los cambios entre modalidades del servicio dentro de SIRBE.</p>

                    <p style="line-height:1.7; margin-top:10px;"><strong>Actuaciones de intervenci&oacute;n:</strong> Actualmente se est&aacute;n parametrizando en SIRBE para poder registrar el soporte de los procesos de atenci&oacute;n detallados del joven.</p>
                </div>
            </div>"""

# Iframe de Power BI para estadísticas (el enlace vive en enlaces/enlaces.xlsx, no en el codigo)
POWERBI_SRC = powerbi_src(BASE)

SECCION_ESTADISTICAS = f"""\
            <div class="content-section" id="estadisticas">
                <div class="card">
                    <h2 class="card-title">Estad&iacute;sticas</h2>
                    <iframe title="Seguimiento t&eacute;cnico" width="100%" height="600" src="{POWERBI_SRC}" frameborder="0" allowFullScreen="true" style="border:1px solid #e0e0e0; border-radius:8px;"></iframe>
                </div>
            </div>"""

# =====================================================================
# Reporte a políticas — Forjar
#
# La lógica de lectura de Excels y render de cards vive en
# _comun/reporte_politicas.py (compartida con JCO y otros servicios).
# Aquí quedan solo las constantes específicas de Forjar: las rutas a los
# Excels maestros y la lista de políticas que Felipe mencionó pero aún
# no están en los Excels (se mostrarán como cards "En revisión").
#
# Nota: Forjar no aparece como responsable directo en el Excel de la PPDJ;
# por eso la PPDJ va como "En revisión" pendiente de confirmar con Felipe.
# =====================================================================

_XLSX_PPDJ_FORJAR = os.path.join(BASE, "ejes", "Políticas", "Reporte Política Pública Distrital de Juventud.xlsx")
_XLSX_REPORTES_EXTERNOS_FORJAR = os.path.join(BASE, "ejes", "Políticas", "Reportes Externos Subdirección Juventud 2026.xlsx")

# Pendientes: una entrada por tema. Cada tupla es
# (tipo, tema, pendiente_con, nota_revision).
# Las políticas étnicas transversales (Indígena 1.3.12, Rrom 6.1.8 y
# Palenquera 6.3.15 — parte de Negra/Afro/Palenquera) se leen desde la hoja
# "Mapeo general" del Excel de Reportes Externos. Aplican a los tres
# servicios (Casas, JCO y Forjar) con la observación literal sobre el flujo
# de PUA y DADE (los datos de Casas y Forjar vienen juntos por Diego Forero).
# Pueblo Raizal NO aplica a Forjar (es JCO específico, transferencias
# monetarias). Reincorporación, PDET y PAT no están en la hoja Forjar; si
# Daniela/Felipe confirman que Forjar reporta a alguno, se agrega al Excel.
_POLITICAS_EN_REVISION_FORJAR = []


# Sección "Reporte a políticas" en HTML. El cuerpo del grid se reemplaza
# en el ensamblaje con la salida de html_cards() del módulo común.
SECCION_REPORTE_POLITICAS = """            <div class="content-section" id="reporte_politicas">
                <div class="card">
                    <h2 class="card-title">Reporte a pol&iacute;ticas</h2>

                    <div class="jp-callout">
                        El <strong>Servicio Forjar Restaurativo</strong> reporta de manera transversal a varias <strong>pol&iacute;ticas p&uacute;blicas distritales</strong> y a varios <strong>planes intersectoriales</strong>.
                    </div>

                    <div class="jp-grid">
%%CARDS_POLITICAS%%
                    </div>
                </div>
            </div>"""

# =====================================================================
# JavaScript - navegación del sidebar
# =====================================================================
JS_NAVEGACION = """\
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
}
function togglePanel(id, card) {
    var panel = document.getElementById(id);
    var allPanels = document.querySelectorAll('.po-detail-panel');
    var isOpen = panel.classList.contains('active');
    allPanels.forEach(function(p) { p.classList.remove('active'); });
    if (!isOpen) {
        panel.classList.add('active');
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}
function filterDocs(stage, el) {
    var yaActiva = el.classList.contains('active');
    document.querySelectorAll('.po-filter-pills .po-pill').forEach(function(p) { p.classList.remove('active'); });
    if (yaActiva) {
        // Click en la pestania activa: deseleccionar y volver a mostrar todas.
        document.querySelectorAll('.po-doc-card').forEach(function(card) { card.style.display = ''; });
    } else {
        el.classList.add('active');
        document.querySelectorAll('.po-doc-card').forEach(function(card) {
            card.style.display = (card.dataset.stage === stage) ? '' : 'none';
        });
    }
}"""


# =====================================================================
# Ensamblaje del HTML final
# =====================================================================
def ensamblar_html():
    """Combina todas las partes y devuelve el HTML completo."""
    # Secciones de contenido en orden de aparición
    seccion_flujo = SECCION_FLUJO_DATOS.replace(
        "%%SVG_DIAGRAMA_FORJAR%%", svg_diagrama_forjar()
    )

    # Reporte a políticas: las cards salen de leer los Excels maestros de
    # Felipe (PPDJ + Reportes Externos hoja Forjar) y de
    # _POLITICAS_EN_REVISION_FORJAR. La lectura y render viven en
    # _comun/reporte_politicas.py.
    filas_politicas = leer_politicas(
        xlsx_ppdj=_XLSX_PPDJ_FORJAR,
        xlsx_externos=_XLSX_REPORTES_EXTERNOS_FORJAR,
        hoja_externos="Forjar",
        patron_responsable_ppdj=r"Forjar",
        tema_ppdj="Política Pública Distrital de Juventud",
        pendientes_revision=_POLITICAS_EN_REVISION_FORJAR,
        temas_mapeo_general_incluir=["1.3.12", "6.1.8", "6.3.15"],
    )
    seccion_reporte_politicas = SECCION_REPORTE_POLITICAS.replace(
        "%%CARDS_POLITICAS%%", html_cards(filas_politicas)
    )

    secciones = "\n\n".join([
        SECCION_WELCOME,
        SECCION_LINEA_TIEMPO,
        SECCION_A_TENER_EN_CUENTA,
        SECCION_EQUIPO,
        SECCION_UBICACION,
        SECCION_MODALIDADES,
        SECCION_PROCESO_OPERATIVO,
        seccion_flujo,
        SECCION_DATOS_SIRBE,
        seccion_reporte_politicas,
        seccion_aliados_forjar(),
        SECCION_ESTADISTICAS,
    ])

    html = f"""\
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestor de conocimiento - Servicio Forjar Restaurativo</title>
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
{JS_NAVEGACION}
</script>
    <!-- Lucide Icons (flat icons de l&iacute;nea, MIT license) para la l&iacute;nea de tiempo -->
    <script src="https://unpkg.com/lucide@latest/dist/umd/lucide.min.js"></script>
    <script>lucide.createIcons();</script>
</body>
</html>"""
    return html


# =====================================================================
# Escritura del archivo
# =====================================================================
if __name__ == "__main__":
    html = ensamblar_html()
    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Archivo generado: {ARCHIVO_SALIDA}")
