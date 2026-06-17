# Migracion de estructura (run-once, idempotente).
#
# Que hace:
#   1. Quita el sufijo "_2025" de los nombres de los HTML generados.
#   2. Mueve las paginas generadas a la subcarpeta html/ (solo index.html
#      queda en la raiz).
#
# Para lograrlo edita los scripts generadores: cambia la ruta de salida y
# ajusta los enlaces relativos internos de cada pagina (al bajar un nivel a
# html/, imagenes/ pasa a ../imagenes/, index.html a ../index.html, etc.).
# Los mapas folium tambien se mueven a html/, por eso los iframe que los
# embeben no cambian (pagina y mapa quedan juntos).
#
# Es idempotente: volver a correrlo no vuelve a anteponer ../ porque los
# patrones viejos ya no existen tras la primera corrida.

from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
SCRIPTS = BASE / "scripts"
EJES = BASE / "ejes"


def aplicar(ruta: Path, reemplazos: list[tuple[str, str, int]]) -> None:
    """Aplica una lista de (viejo, nuevo, esperado) sobre un archivo.

    Imprime cuantas veces aparecio cada patron y avisa si no coincide con
    lo esperado (esperado=None desactiva la verificacion)."""
    if not ruta.exists():
        print(f"  OMITIDO (no existe): {ruta}")
        return
    texto = ruta.read_text(encoding="utf-8")
    original = texto
    print(f"\n{ruta.relative_to(BASE)}")
    for viejo, nuevo, esperado in reemplazos:
        n = texto.count(viejo)
        marca = "ok" if (esperado is None or n == esperado) else f"AVISO esperaba {esperado}"
        print(f"   {n:>2}x  {viejo[:58]:<58} [{marca}]")
        texto = texto.replace(viejo, nuevo)
    if texto != original:
        ruta.write_text(texto, encoding="utf-8")
        print("   -> guardado")
    else:
        print("   -> sin cambios")


# Patrones comunes para las paginas que bajan a html/
def relativos_comunes() -> list[tuple[str, str, int | None]]:
    return [
        ('"imagenes/', '"../imagenes/', None),
        ('href="index.html"', 'href="../index.html"', 1),
    ]


# --- 1. Home (index.html): se queda en la raiz, solo cambian los enlaces a
#         las paginas que se movieron a html/ ---
aplicar(SCRIPTS / "generar_home_servicios.py", [
    ('"gestion_conocimiento_juventud_2025.html"', '"html/gestion_conocimiento_juventud.html"', 1),
    ('"gestion_conocimiento_jco_2025.html"',      '"html/gestion_conocimiento_jco.html"', 1),
    ('"gestion_conocimiento_forjar_2025.html"',   '"html/gestion_conocimiento_forjar.html"', 1),
    ('"gestion_conocimiento_alertas_2025.html"',  '"html/gestion_conocimiento_alertas.html"', 1),
    ('href="unidades_operativas.html"',           'href="html/unidades_operativas.html"', 1),
])

# --- 2. Casas de Juventud ---
aplicar(SCRIPTS / "generar_juventud.py", relativos_comunes() + [
    ('"ejes/', '"../ejes/', 5),
    ('os.path.join(BASE, "gestion_conocimiento_juventud_2025.html")',
     'os.path.join(BASE, "html", "gestion_conocimiento_juventud.html")', 1),
    ('os.path.join(BASE, "mapa_casas_juventud.html")',
     'os.path.join(BASE, "html", "mapa_casas_juventud.html")', 1),
])

# --- 3. Jovenes con Oportunidades ---
aplicar(SCRIPTS / "generar_gc_jco.py", relativos_comunes() + [
    ('os.path.join(directorio, "gestion_conocimiento_jco_2025.html")',
     'os.path.join(directorio, "html", "gestion_conocimiento_jco.html")', 1),
])

# --- 4. Forjar ---
aplicar(SCRIPTS / "generar_gc_forjar.py", relativos_comunes() + [
    ('os.path.join(BASE, "gestion_conocimiento_forjar_2025.html")',
     'os.path.join(BASE, "html", "gestion_conocimiento_forjar.html")', 1),
    ('os.path.join(BASE, "mapa_forjar.html")',
     'os.path.join(BASE, "html", "mapa_forjar.html")', 1),
])

# --- 5. Alertas / Parche Seguro ---
aplicar(SCRIPTS / "generar_gc_alertas.py", relativos_comunes() + [
    ('NOMBRE_ARCHIVO = "gestion_conocimiento_alertas_2025.html"',
     'NOMBRE_ARCHIVO = "gestion_conocimiento_alertas.html"', 1),
    ('os.path.join(directorio, NOMBRE_ARCHIVO)',
     'os.path.join(directorio, "html", NOMBRE_ARCHIVO)', 1),
])

# --- 6. Unidades operativas ---
aplicar(SCRIPTS / "generar_unidades_operativas.py", relativos_comunes() + [
    ('os.path.join(BASE, "unidades_operativas.html")',
     'os.path.join(BASE, "html", "unidades_operativas.html")', 1),
    ('os.path.join(BASE, "mapa_unidades_operativas.html")',
     'os.path.join(BASE, "html", "mapa_unidades_operativas.html")', 1),
])

# --- 7. Script de PDFs para NotebookLM: solo actualiza los 4 nombres ---
aplicar(SCRIPTS / "generar_pdfs_para_notebooklm.py", [
    ('"gestion_conocimiento_juventud_2025.html"', '"html/gestion_conocimiento_juventud.html"', 1),
    ('"gestion_conocimiento_jco_2025.html"',      '"html/gestion_conocimiento_jco.html"', 1),
    ('"gestion_conocimiento_forjar_2025.html"',   '"html/gestion_conocimiento_forjar.html"', 1),
    ('"gestion_conocimiento_alertas_2025.html"',  '"html/gestion_conocimiento_alertas.html"', 1),
])

# --- 8. Paginas de ejes: el enlace de regreso vive en el HTML (el script de
#         ejes lo preserva, no lo genera). Se edita el HTML directamente. ---
for nombre in ["Bienestar.html", "Cultura.html", "Inclusion.html",
               "Liderazgo.html", "SIDICU.html"]:
    aplicar(EJES / nombre, [
        ('../gestion_conocimiento_juventud_2025.html',
         '../html/gestion_conocimiento_juventud.html', 1),
    ])

print("\nMigracion aplicada. Ahora regenera las paginas corriendo los scripts.")
