"""
Genera PDFs imprimibles de los HTML del gestor para subirlos como fuentes a
NotebookLM. Cada PDF preserva el diseño (línea de tiempo, cards, módulos) y
mantiene el texto vectorial seleccionable (NotebookLM lo lee sin OCR).

Genera en notebooklm/:
- 03_casas_de_juventud.pdf  (página principal + 5 ejes)
- 04_jco.pdf
- 05_forjar.pdf             (página principal + restaurativo)
- 06_parche_seguro.pdf

Usa Microsoft Edge en modo headless. Antes de imprimir inyecta un <style>
que fuerza todas las .content-section visibles y oculta el sidebar
(que en navegación normal funciona con JS y por defecto deja en
display:none todas las secciones excepto la activa).
"""

from pathlib import Path
import subprocess
import tempfile

BASE = Path(__file__).resolve().parent.parent
SALIDA = BASE / "notebooklm"

EDGE_RUTAS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

CSS_IMPRESION = """
<style>
/* Forzar visibilidad de todas las secciones al imprimir */
.content-section { display: block !important; }
.sidebar-items { display: block !important; }
.sidebar { display: none !important; }
.main-content {
    margin-left: 0 !important;
    padding-left: 24px !important;
    max-width: 100% !important;
    width: 100% !important;
}
/* Fuente neutra para que el texto del PDF se extraiga sin espacios espurios
   (NotebookLM lee el texto vectorial; las fuentes display tipo Anton/Antonio
   con letter-spacing causan que el extractor corte palabras). */
*, *::before, *::after {
    font-family: Arial, Helvetica, sans-serif !important;
    letter-spacing: normal !important;
    word-spacing: normal !important;
    -webkit-font-feature-settings: normal !important;
    font-feature-settings: normal !important;
}
@media print {
    .content-section { display: block !important; page-break-inside: avoid; }
    .sidebar { display: none !important; }
    body { margin: 0 !important; }
}
</style>
"""

# Por cada PDF, lista de HTMLs a unir (primero el principal, luego los extras)
GRUPOS = {
    "03_casas_de_juventud.pdf": [
        "html/gestion_conocimiento_juventud.html",
        "ejes/Bienestar.html",
        "ejes/Cultura.html",
        "ejes/Inclusion.html",
        "ejes/Liderazgo.html",
        "ejes/SIDICU.html",
    ],
    "04_jco.pdf": [
        "html/gestion_conocimiento_jco.html",
    ],
    "05_forjar.pdf": [
        "html/gestion_conocimiento_forjar.html",
        "forjar/forjar-restaurativo (HTML de Valentina).html",
    ],
    "06_parche_seguro.pdf": [
        "html/gestion_conocimiento_alertas.html",
    ],
}


def _localizar_edge() -> str:
    for ruta in EDGE_RUTAS:
        if Path(ruta).exists():
            return ruta
    raise FileNotFoundError(
        "No se encontró Microsoft Edge en las rutas conocidas. "
        "Edita EDGE_RUTAS en el script."
    )


def _html_imprimible(ruta_html: Path, tmp_dir: Path) -> Path:
    """Crea una copia del HTML con CSS de impresión inyectado."""
    html = ruta_html.read_text(encoding="utf-8")
    if "</head>" in html:
        html_mod = html.replace("</head>", CSS_IMPRESION + "</head>", 1)
    else:
        html_mod = CSS_IMPRESION + html
    tmp = tmp_dir / ruta_html.name
    tmp.write_text(html_mod, encoding="utf-8")
    return tmp


def _imprimir_a_pdf(edge: str, html: Path, pdf: Path) -> None:
    """Imprime un HTML a PDF con Edge headless."""
    file_url = "file:///" + str(html).replace("\\", "/")
    cmd = [
        edge,
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--virtual-time-budget=10000",
        f"--print-to-pdf={pdf}",
        file_url,
    ]
    subprocess.run(cmd, capture_output=True, check=True)


def _unir_pdfs(pdfs: list[Path], destino: Path) -> None:
    """Une varios PDFs en uno solo, en orden."""
    from pypdf import PdfWriter
    w = PdfWriter()
    for p in pdfs:
        w.append(str(p))
    with open(destino, "wb") as f:
        w.write(f)


def _generar_grupo(edge: str, nombre_pdf: str, htmls: list[str], tmp_dir: Path) -> Path:
    pdfs_parciales: list[Path] = []
    for i, ruta_rel in enumerate(htmls):
        ruta_html = BASE / ruta_rel
        if not ruta_html.exists():
            print(f"  AVISO: no existe {ruta_rel}, se omite.")
            continue
        # 1. Inyectar CSS y guardar HTML imprimible en tmp_dir con nombre único
        prefijo = f"{i:02d}_"
        tmp_html_dst = tmp_dir / (prefijo + ruta_html.name)
        html_text = ruta_html.read_text(encoding="utf-8")
        if "</head>" in html_text:
            html_text = html_text.replace("</head>", CSS_IMPRESION + "</head>", 1)
        else:
            html_text = CSS_IMPRESION + html_text
        tmp_html_dst.write_text(html_text, encoding="utf-8")
        # 2. Imprimir a PDF
        pdf_parcial = tmp_dir / (prefijo + ruta_html.stem + ".pdf")
        _imprimir_a_pdf(edge, tmp_html_dst, pdf_parcial)
        pdfs_parciales.append(pdf_parcial)
    destino = SALIDA / nombre_pdf
    if len(pdfs_parciales) == 1:
        # Solo un HTML: mover el PDF
        destino.write_bytes(pdfs_parciales[0].read_bytes())
    else:
        _unir_pdfs(pdfs_parciales, destino)
    return destino


def main() -> None:
    edge = _localizar_edge()
    print(f"Edge: {edge}")
    print(f"Base: {BASE}")
    print(f"Salida: {SALIDA}")
    SALIDA.mkdir(parents=True, exist_ok=True)
    print()

    with tempfile.TemporaryDirectory(prefix="pdf_notebooklm_") as tmp:
        tmp_dir = Path(tmp)
        for nombre, htmls in GRUPOS.items():
            print(f"Generando {nombre} ({len(htmls)} HTML)...")
            destino = _generar_grupo(edge, nombre, htmls, tmp_dir)
            tam = destino.stat().st_size / 1024
            print(f"  -> {destino.name}  ({tam:.1f} KB)")
            print()

    print("Listo. Sube cada PDF a NotebookLM como fuente.")


if __name__ == "__main__":
    main()
