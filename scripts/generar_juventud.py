"""
Generador del HTML de Casas de Juventud.

Script extraído automáticamente del notebook generar_gc_juventud.ipynb
el 2026-04-13, con los bugs del notebook corregidos:
  - Strings partidas por newlines literales (sintaxis inválida).
  - Dos </div> extras en el bloque welcome que rompían el layout desktop.
  - Dobles llaves {{}} en reglas CSS del árbol SIRBE.

Para correrlo: python scripts/generar_juventud.py
"""

# ========== celda 1: pqodiu9c96j ==========
# Carga de datos
import pandas as pd
import os, re, sys, unicodedata

# Modulo compartido para la seccion de aliados
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _comun.aliados import seccion_casas as seccion_aliados_casas
# CSS de la línea de tiempo en chevrones (compartido con Forjar y JCO)
from _comun.estilos import CSS_LINEA_TIEMPO_CHEVRON
# Diagrama de flujo del proceso SIRBE para Casas de Juventud
from _comun.diagramas_flujo import svg_diagrama_casas
# Enlace del tablero de Power BI (se lee desde enlaces/enlaces.xlsx)
from _comun.tablero import powerbi_src
# Componente "Reporte a políticas" (compartido con JCO y Forjar)
from _comun.reporte_politicas import (
    CSS_REPORTE_POLITICAS,
    leer_politicas,
    html_cards,
)

# Raíz del proyecto (un nivel arriba de scripts/)
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(BASE, "datos")
PROYECTO = os.path.dirname(BASE)  # directorio padre del proyecto

# Iframe de Power BI (el enlace vive en enlaces/enlaces.xlsx, no en el codigo)
POWERBI_SRC = powerbi_src(BASE)
# Primero intenta leer la base limpia, si no existe lee la original
RUTA_LIMPIA = os.path.join(PROYECTO, "DATOS SIRBE", "intermedias", "sirbe_2025_limpio.xlsx")
RUTA_ORIGINAL = os.path.join(PROYECTO, "DATOS SIRBE", "originales",
                          "REQ-SIMI-42642-1292_BasePUA2025_Juventud.xlsx")

if os.path.exists(RUTA_LIMPIA):
    df_raw = pd.read_excel(RUTA_LIMPIA)
    print("Leyendo base LIMPIA")
else:
    df_raw = pd.read_excel(RUTA_ORIGINAL, sheet_name="Sheet 1")
    print("Leyendo base ORIGINAL (ejecutar limpieza_sirbe_2025.ipynb primero)")
print(f"Filas totales: {len(df_raw):,.0f}".replace(",", "."))

# Filtrar solo Casas de Juventud (Forjar se trabajará aparte)
df = df_raw[df_raw["SERVICIO"] == "CASAS DE LA JUVENTUD"].copy()
print(f"Filas Casas de Juventud: {len(df):,.0f}".replace(",", "."))
print(f"Columnas: {list(df.columns)}")
df.head(3)

# ========== celda 3: wbaehubf8nk ==========
# Homologación por eje (4 ejes oficiales)
# Fuente: PPT "Primera reunión del servicio" + script homologacion_sirbe

HOMOLOGACION_EJES = {
    1485: "BIENESTAR", 1486: "BIENESTAR", 1487: "BIENESTAR",
    1488: "BIENESTAR", 1599: "BIENESTAR", 511: "BIENESTAR",
    1489: "CULTURA", 1490: "CULTURA", 1207: "CULTURA", 1491: "CULTURA",
    1496: "INCLUSIÓN SOCIAL Y PRODUCTIVA", 1497: "INCLUSIÓN SOCIAL Y PRODUCTIVA",
    1498: "INCLUSIÓN SOCIAL Y PRODUCTIVA", 1499: "INCLUSIÓN SOCIAL Y PRODUCTIVA",
    1492: "LIDERAZGO Y PARTICIPACIÓN", 1493: "LIDERAZGO Y PARTICIPACIÓN",
    1494: "LIDERAZGO Y PARTICIPACIÓN", 1495: "LIDERAZGO Y PARTICIPACIÓN",
    1502: "LIDERAZGO Y PARTICIPACIÓN", 1503: "LIDERAZGO Y PARTICIPACIÓN",
    1504: "LIDERAZGO Y PARTICIPACIÓN", 1505: "LIDERAZGO Y PARTICIPACIÓN",
}
df["cod_actividad"] = pd.to_numeric(
    df["NOMACTIVIDAD_CURSO"].astype(str).str[:4], errors="coerce"
)
df["eje"] = df["cod_actividad"].map(HOMOLOGACION_EJES).fillna("SIN CLASIFICAR")

# Distribución por eje
print("Distribución por eje:")
for eje, n in df["eje"].value_counts().items():
    print(f"  {eje:<40s} {n:>6}  ({n/len(df)*100:.1f}%)")

# ========== celda 5: o20vslxismd ==========
# Subtemas por eje (basados en NOMBRE_CURSO)

SUBTEMAS = {
    "BIENESTAR": [
        ("Prevención de consumo de SPA", [r"PREVENCION.*(?:CONSUMO|SPA|SUSTANCIA)", r"PREVENCI.N.*(?:CONSUMO|SPA)", r"PREVENSION.*CONSUMO", r"POREVENCION", r"\bSPA\b", r"SUSTANCIAS PSICO", r"CONSUMO.*SPA", r"CUIDADO FRENTE AL CONSUMO", r"CLOSET PSICOACTIVO", r"CON SUMO CUIDADO", r"PATRONES DE CONSUMO"]),
        ("Prevención de violencias basadas en género", [r"VIOLENCIA", r"VBG", r"PVBG", r"DVBG", r"ESTEREOTIPOS DE GENERO", r"MASCULINIDADES", r"MITOS DEL AMOR ROMANTICO", r"25.?N"]),
        ("Derechos sexuales y reproductivos", [r"DERECHO.? SEXUAL", r"REPRODUCTIV", r"DSDR", r"DSYR", r"DRYDR", r"DSYDR", r"DDSS", r"DS\.?R", r"DDRRYRR", r"ANATOMIA", r"ANTICONCEP", r"SEXUALIDAD", r"MENSTR", r"MI CUERPO MI", r"CONSENTIMIENTO SEXUAL", r"SEMANA ANDINA", r"PROMOCION DSR"]),
        ("Salud mental", [r"SALUD MENTAL", r"AUTOCUIDADO", r"BIENESTAR SALUD", r"SUICIDIO", r"CONDUCTA SUICIDA", r"MINDFUL", r"RESPIRO", r"CUIDADOR", r"JORNADA DE BIENESTAR", r"INTERCAMBIO DE SABERES", r"SALAS DE ESCUCHA", r"SALAS DE MEDITACION", r"ARTE DE SENTIR", r"DIMENSIONES DE LA SALUD"]),
        ("Gestión emocional", [r"GESTION EMOCIONAL", r"REGULACION EMOCIONAL", r"RESOLUCION DE CONFLICTOS", r"COMUNICACION ASERTIVA", r"INTELIGENCIA EMOCIONAL", r"ESTABLECIMIENTO DE LIMITES", r"CUIDAR PARA RESOLVER", r"CIUDAR PARA RESOLVER", r"RECONOCIMIENTO DE EMOCIONES", r"CONECTANDO EMOCIONES"]),
        ("Orientación psicosocial", [r"ORIENTACION.*PSICOSOCIAL", r"PSICOSOCIAL", r"HABILIDADES.*VIDA"]),
    ],
    "CULTURA": [
        ("Artes y formación artística", [r"ARTE", r"PLASTICAS", r"ESCENICAS", r"MUSICALES", r"URBANAS", r"AUDIOVISUAL", r"FOTOGRAFIA", r"GRAFITI", r"GRAFFIT", r"GRAFFITI", r"MURALISMO", r"FORMACION ARTISTICA", r"PRACTICAS ARTISTICAS", r"CLASES ARTISTICAS", r"CIRCO", r"CIRCENSES", r"DANZA", r"TEATRO", r"MUSICA", r"ROCK", r"GUITARRA", r"PIANO", r"BATERIA", r"MALABARISMO", r"SERIGRAFIA", r"CINE", r"PODCAST", r"GRABACION", r"ENSAYAD", r"SALAS DE ENSAYO", r"IDARTES", r"STOP MOTION", r"SAFARI FOTOGRAFICO"]),
        ("Uso de espacios y gestión cultural", [r"USO DE ESPACIOS", r"GESTION CULTURAL", r"GESTION CUTURAL"]),
        ("Deportes y recreación", [r"DEPORTE", r"DEPORTIV", r"CAPOEIRA", r"BOXEO", r"FUTBOL", r"VOLEIBOL", r"ULTIMATE", r"PING PONG", r"SEMILLERO", r"TORNEO", r"ENCUENTRO DEPORTIVO"]),
        ("Eventos y semanas temáticas", [r"SEMANA DE LA JUVENTUD", r"SEMANA DE JUVENTUD", r"SEMANA LOCAL", r"HALLOWEEN", r"CONCIERTO", r"REINADO", r"BIZ NATION"]),
        ("Tiempo libre e idiomas", [r"TIEMPO LIBRE", r"APROVECHAMIENTO", r"IDIOMAS", r"INGLES"]),
    ],
    "INCLUSIÓN SOCIAL Y PRODUCTIVA": [
        ("Empleo y empleabilidad", [r"EMPLEO", r"EMPLEABILIDAD", r"EMBLEABILIDAD", r"INTERMEDIACION LABORAL", r"HOJA DE VIDA", r"FERIA.*EMPLEA", r"INCLUSION LABORAL", r"FORMACION DE EMPLEND", r"CRECIENDO JUNTOS", r"DISRUPTIA"]),
        ("Emprendimiento", [r"EMPRENDIMIENTO", r"EMPRENDIMIENT", r"INNOVERS", r"ISYP", r"GENERACION DE INGRESOS"]),
        ("Educación financiera", [r"EDUCACI.N FINANCIERA", r"FINANZ", r"AHORRO", r"PRESUPUESTO", r"COSTO INVISIBLE"]),
        ("Proyecto de vida y educación flexible", [r"PROYECTO DE VIDA", r"EDUCACION FLEXIBLE", r"EDUCACION RUTA", r"MODELO.*EDUCACION FLEXIBLE", r"RUTAS.*EDUCACION", r"FORMACION RUTAS", r"FORMACION RUTA", r"FORMACION PARA EL PROYECTO", r"ACCESO A LA EDUCACI", r"ACCESO A EDUCACION", r"LECTO.?ESCRITURA", r"PRE.?ICFES", r"PREICFES"]),
        ("Orientación socio-ocupacional", [r"ORIENTACION SOCIO", r"ORIENTACION.*INCLUSION", r"INCLUSION SOCIAL.*PRODUCTIV"]),
        ("Habilidades TIC", [r"HABILIDADES INFORMATICA", r"OFERTA TIC", r"SALA.? TIC", r"FORTALECIMIENTO.*HABILIDADES"]),
    ],
    "LIDERAZGO Y PARTICIPACIÓN": [
        ("Política pública de juventud", [r"POLITICA PUBLICA", r"POLITIC.? PUBLICA", r"PLITICA PUBLICA", r"SICIALIZACION DE POLITICA", r"SOCIALIZACION.*POLITICA", r"SOCIALIZACION POLITIC"]),
        ("Liderazgo juvenil", [r"LIDERAZGO", r"LIDEREAZGO", r"ORGANIZACION JUVENIL", r"PARTICIPACION.*JUVENIL", r"INCIDENCIA"]),
        ("Asesoría jurídica y participación", [r"ASESORIA JURIDICA", r"ASESPROA JURIDICA", r"SESORIA JURIDICA", r"JURIDICA"]),
        ("Derechos humanos y ciudadanía", [r"DERECHOS HUMANOS", r"PROMOCION.*DDHH", r"CONTROL SOCIAL"]),
        ("Voluntariado y diálogos", [r"VOLUNTARIADO", r"DIALOGOS INTERGENERACIONAL", r"ESCENARIOS DE DIALOGO"]),
    ],
}

def subtematizar(nombre, eje):
    """Asigna un subtema dentro del eje correspondiente."""
    if pd.isna(nombre) or eje not in SUBTEMAS:
        return "Otros"
    texto = nombre.upper()
    for subtema, patrones in SUBTEMAS[eje]:
        for p in patrones:
            if re.search(p, texto):
                return subtema
    return "Otros"

df["subtema"] = df.apply(lambda r: subtematizar(r["NOMBRE_CURSO"], r["eje"]), axis=1)

# Revisar subtemas por eje
for eje in ["BIENESTAR", "CULTURA", "INCLUSIÓN SOCIAL Y PRODUCTIVA", "LIDERAZGO Y PARTICIPACIÓN"]:
    sub = df[df["eje"] == eje]
    print(f"\n{eje}:")
    for st, n in sub["subtema"].value_counts().items():
        print(f"    {st:<45s} {n:>5}  ({n/len(sub)*100:.1f}%)")

# ========== celda 6: f8nrz21tpre ==========
# Revisar registros sin clasificar (para ajustar reglas si es necesario)
sin_clasificar = df[df["eje"] == "SIN CLASIFICAR"]
print(f"Registros SIN CLASIFICAR: {len(sin_clasificar)} ({len(sin_clasificar)/len(df)*100:.1f}%)")
if len(sin_clasificar) > 0:
    print("\nCursos sin clasificar (top 20):")
    for curso, n in sin_clasificar["NOMBRE_CURSO"].value_counts().head(20).items():
        cod = sin_clasificar[sin_clasificar["NOMBRE_CURSO"] == curso]["cod_actividad"].iloc[0]
        print(f"  cod={int(cod) if pd.notna(cod) else '?':>5}  {n:>4}x  {curso}")

# Registros con subtema "Otros" dentro de ejes clasificados
otros = df[(df["eje"] != "SIN CLASIFICAR") & (df["subtema"] == "Otros")]
print(f"\nRegistros con subtema 'Otros' (dentro de ejes): {len(otros)} ({len(otros)/len(df)*100:.1f}%)")
if len(otros) > 0:
    print("\nCursos en 'Otros' (top 15):")
    for curso, n in otros["NOMBRE_CURSO"].value_counts().head(15).items():
        eje = otros[otros["NOMBRE_CURSO"] == curso]["eje"].iloc[0]
        print(f"  [{eje[:12]:<12}] {n:>4}x  {curso}")

# ========== celda 8: h6migofumn8 ==========
# Normalización de nombres de curso
def normalizar_curso(nombre):
    """Agrupa variaciones del mismo curso (prefijos BI/EM, tildes, errores)."""
    if pd.isna(nombre):
        return "SIN NOMBRE"
    t = nombre.upper().strip()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    # Quitar prefijos
    for p in [r"^BI\s+TALLER\s+(?:DE\s+|EN\s+)?", r"^EM\s+TALLER\s+(?:DE\s+|EN\s+)?",
              r"^TALLER\s+(?:DE\s+|EN\s+|SOBRE\s+)?", r"^TALLERES\s+(?:DE\s+|EN\s+|INFORMATIVOS\s+DE\s+)?",
              r"^BI\s+", r"^EM\s+", r"^BIENESTAR\s+", r"^CULTURA\s+",
              r"^FORMACION\s+EN\s+", r"^SOCIALIZACION\s+DE\s+",
              r"^SOCIALIZACION\s+", r"^JORNADA\s+DE\s+",
              r"^PROMOCION\s+Y?\s*", r"^INCLUSION\s+SOCIAL\s+"]:
        t = re.sub(p, "", t)
    # Normalizar variaciones conocidas
    for patron, reemplazo in [
        # --- DSDR ---
        (r"DERECHOS SEXUALES Y DERECHOS REPRODUCTIVOS", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUIALES DERECHOS REPRODUCTIVOS", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES\s+DERECHOS REPRODUCTIVOS", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DSYDR.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+SEMANA.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+DECIDE.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+AUTOCUIDADO.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+MITOS.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+GRATIFERIA", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+FERIA.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS\s+TALLER.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"INFORMATIVO EN DERECHOS SEXUALES Y REPRODUCTIVOS", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"LABORATORIO DERECHOS SEXUALES Y REPRODUCTIVOS", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"INTERCAMBIO DE SABERES TEMATICA DSYR", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DECIME COMO CUANDO.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"CONOCE TUS DERECHOS.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"^DSDR.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"MURAL SOBRE CONSENTIMIENTO SEXUAL", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"CINE COMUNITARIO 8M BIENESTAR DS Y DR", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"FERIA PROMOCION DSR.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y REPRODUCTIVOS SEMANA DE JUVENTUD", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        (r"DERECHOS SEXUALES Y DERECHOS REPRODUCTIVOS.*", "DERECHOS SEXUALES Y REPRODUCTIVOS"),
        # --- Prevención consumo SPA ---
        (r"PREVENCI.?N DE CONSUMO DE SUSTANCIAS?\s?PSICO?ACTIVAS?\s?(?:SPA)?", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION CONSUMO DE?\s?(?:SUSTANCIAS?\s?)?(?:PSICO?ACTIVAS?\s?)?(?:SPA)?", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO?DE?\s?SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DEL? CONSUMO DE SUSTANCIAS PSICOACTIVAS SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE SUSTANCIAS PSICOACTIVAS SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMODE SUSTANCIAS PSICOACTIVAS", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION Y CONSUMO DE SUSTANCIAS PSICOACTIVAS SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO DE SUSTACIAS PSICOACTIVAS", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO DE SUTANCIAS PSICOACTIVAS", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENSION DE? CONSUMO SUSTANCIAS PSICOACTIVAS SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"POREVENCION DE CONSUMO DE SUSTANCIAS PSICOACTIVAS", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE SPA.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION SPA.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"SPA GESTION EMOCIONAL.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DECIDE COMO CUANDO.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"SEMANA DE PREVENCION DEL CONSUMO DE SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"SPA MITOS DEL CONSUMO", "PREVENCION DE CONSUMO DE SPA"),
        (r"CUIDADO FRENTE AL CONSUMO.*SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION CONSUMOI DE SUSTACIOAS.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DEL CONSUMO DE SPA.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION CONSUMO SPA.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCCION DEL CONSUMO DE SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION EN CONSUMO.*SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO DE SUSTANCIAS PSICOACTIVAS SPA", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO DE SUSTANCIAS PSICOACTIVAS$", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION DE CONSUMO DE SUSTANCIA PSICOACTIVAS.*", "PREVENCION DE CONSUMO DE SPA"),
        (r"PREVENCION CONSUMO SUSTANCIAS.*", "PREVENCION DE CONSUMO DE SPA"),
        # --- Prevención VBG ---
        (r"PREVENCI.?N DE VIOLENCIAS?\s?BASADA?S?\s?EN\s?G.?NERO?S?.*", "PREVENCION DE VBG"),
        (r"PREVENCION (?:DE |EN )?VIVENCIAS BASADAS EN GENERO", "PREVENCION DE VBG"),
        (r"PREVENCION (?:EN |DE )?VIOLENCIAS? BASADA?S? EN GENERO?S?.*", "PREVENCION DE VBG"),
        (r"PREVENCION DEVIOLENCIAS BASADAS EN GENERO", "PREVENCION DE VBG"),
        (r"PREVENCIN DE VIOLENCIAS BASADAS EN GENERO", "PREVENCION DE VBG"),
        (r"PREVENCION DE VIOLENCIA BASADOS EN GENERO", "PREVENCION DE VBG"),
        (r"PREVENCION DE VIOLENCIAS BASADO EN GENERO", "PREVENCION DE VBG"),
        (r"PREVENCION DE VIOLENCIAS\b", "PREVENCION DE VBG"),
        (r"PREVENCION DE VBG\b", "PREVENCION DE VBG"),
        (r"PVBG.*", "PREVENCION DE VBG"),
        (r"DVBG.*", "PREVENCION DE VBG"),
        (r"POR ELLAS POR TODAS POR NOSOTRAS", "PREVENCION DE VBG"),
        (r"PREVENCION DE VIOLENCIAS DIGITALES", "PREVENCION DE VBG"),
        (r"OFERTA Y PROYECTO ALCALDIA PVBG", "PREVENCION DE VBG"),
        (r"PREVENCION DE VIOLENCIAS ESTEREOTIPOS DE GENERO", "PREVENCION DE VBG"),
        (r"CINE FORO PVBG", "PREVENCION DE VBG"),
        (r"TALLERES DE PREVENCIN DE VIOLENCIAS", "PREVENCION DE VBG"),
        (r"^PREVENCIN DE VIOLENCIAS$", "PREVENCION DE VBG"),
        (r"^PREVENCION DE VIOLENCIAS$", "PREVENCION DE VBG"),
        # --- Salud mental ---
        (r"SALUD MENTAL\s+CONVERSATORIO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+GESTION EMOCIONAL", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+AUTOCUIDADO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+BIENESTAR.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+TRANSFORMACI?O?N CULTURAL.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+JORNADA.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+PREVENCION.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+TOMA DE DECISIONES", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+RECONOCIMIENTO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+DIMENSIONES.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+ESTABLECIMIENTO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+IDENTIDAD.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+CONECTANDO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+CUERPOS.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+CUIDADO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+CICLO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+SERVICIO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+PODCAST.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+TEJIENDO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+MANDALAS", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+EL ARTE.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+SEMANA.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+CONCIERTO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+FERIA.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+GRIPO.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+JORNADAS.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+TALLER\b.*", "SALUD MENTAL"),
        (r"SALUD MENTAL\s+GRUPO FOCAL", "SALUD MENTAL"),
        (r"^GESTION DE EMOCIONAL$", "GESTION EMOCIONAL Y RESOLUCION DE CONFLICTOS"),
        (r"^CONVERSATORIO SALUD MENTAL.*", "SALUD MENTAL"),
        (r"^BIENESTAR E IDENTIDAD$", "SALUD MENTAL"),
        # --- Cultura: uso de espacios ---
        (r"USO DE ESPACI?OS?\s?PARA\s?(?:EL\s)?DESARROLL?O?S?\s?(?:DE\s)?(?:LAS?\s)?PR[AE]CTICAS?\s?(?:ARTISTICAS?|JUVENILES)(?:\s?CDJ)?.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO (?:DE ESPACOS|PARA EL DESARROLLO DE PRACTICAS ARTISTICAS).*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USOS? DE ESPACIOS PARA DESARROLLO PREACTICAS ARTISTICAS", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS LIBRES DE PRACTICA ARTISTICA", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO ADECUADO DE ESPACIOS PARA EL DESARROLLO DE PRACTICAS ART.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS CDJ", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS CONCURSO.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS DE CDJ.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS SEMANA.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS PARA DESARROLLO DE PRACTICAS ARTISTICAS.*", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS PARA DESARROLLOS DE PRACTICAS ARTISTICAS", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS PARA DESARROLOS DE PRACTICAS ARTISTICAS", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        (r"USO DE ESPACIOS POLITICA PUBLICA.*", "POLITICA PUBLICA DE JUVENTUD"),
        (r"USO DE ESPACIOS PARA DESARROLLO DE PRACTICAS ARTISTICAS", "USO DE ESPACIOS PARA PRACTICAS ARTISTICAS"),
        # --- Inclusión: empleo ---
        (r"^EMPLEO\s+SEMANA.*", "EMPLEO"),
        (r"^EMPLEO\s+FERIA.*", "EMPLEO"),
        (r"^EMPLEO\s+FORMACION.*", "EMPLEO"),
        (r"^EMPLEO\s+TALLER.*", "EMPLEO"),
        (r"^EMPLEO\s+ANDI.*", "EMPLEO"),
        (r"^EMPLEO\s+HERRAMIENTAS.*", "EMPLEO"),
        (r"SEMANA DE LA JUVENTUD EMPLEO", "EMPLEO"),
        (r"SEMANA DE LA JUVENTUD FERIA.*", "EMPLEO"),
        (r"^EMPLEOSEMANA.*", "EMPLEO"),
        (r"^EMPLEO_SEMANA.*", "EMPLEO"),
        (r"SERVICIOS\s+DE INTERMEDIACION LABORAL", "EMPLEO"),
        (r"SERVICIO DE INTERMEDIACION LABORAL", "EMPLEO"),
        (r"FERIA EMPLEABILIDAD EMPLEO", "EMPLEO"),
        (r"ACCIONES FORMACION EN EMPLEABILIDAD", "EMPLEO"),
        (r"FORMACION DE EMPLEABILIDAD.*", "EMPLEO"),
        (r"EMPLEABILIDAD\b", "EMPLEO"),
        (r"^EMPLO$", "EMPLEO"),
        (r"INCLUSION SOCIAL Y PRODUCTIVA EMPLEO", "EMPLEO"),
        (r"^ACCIONES PARA LA INCLUSION LABORAL.*", "EMPLEO"),
        # --- Inclusión: emprendimiento ---
        (r"EMPRENDIMIENTO.INNOVERS", "EMPRENDIMIENTO INNOVERS"),
        (r"EMPRENDIMIENTP", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+SEMANA.*", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+FERIA.*", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+FERIAS.*", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+TALLERES.*", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+TALLER.*", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+ISYP", "EMPRENDIMIENTO"),
        (r"^EMPRENDIMIENTO\s+FORMACION.*", "EMPRENDIMIENTO"),
        (r"ISYP EMPRENDIMIENTO", "EMPRENDIMIENTO"),
        (r"Y PRODUCTIVA EMPRENDIMIENTO", "EMPRENDIMIENTO"),
        (r"INCLUSION PRODUCTIVA EMPRENDIMIENTO", "EMPRENDIMIENTO"),
        # --- Inclusión: proyecto de vida ---
        (r"^DEL PROYECTO DE VIDA", "PROYECTO DE VIDA"),
        (r"PARA EL PROYECTO DE VIDA.*", "PROYECTO DE VIDA"),
        (r"PROYECTO DE VIDA\b", "PROYECTO DE VIDA"),
        # --- Educación financiera ---
        (r"EDUCACI.N FINANCIERA.*", "EDUCACION FINANCIERA"),
        (r"ORIENTACION SOCIO.?OCUPACIONAL.*", "ORIENTACION SOCIO-OCUPACIONAL"),
        # --- Habilidades TIC ---
        (r"HABILIDADES INFORMATIC.?S?\s?OFERTA\s?TICS?", "HABILIDADES TIC"),
        (r"HABILIDADES INFORMATIOCAS OFERTA TIC", "HABILIDADES TIC"),
        (r"OFERTA TIC HABILIDADES INFORMATICAS", "HABILIDADES TIC"),
        (r"FORTALECIMIENTO DE HABILIDADES.*TIC", "HABILIDADES TIC"),
        (r"FORTALECIMIENTO HABILIDADES SALA TIC", "HABILIDADES TIC"),
        (r"SALA TIC.*", "HABILIDADES TIC"),
        (r"HABILIDADES INFORMATICA.*", "HABILIDADES TIC"),
        # --- Educación flexible ---
        (r"(?:EDUCACION\s+)?(?:FORMACION\s+)?RUTAS?\s?DE?\s?MODELO\s?DE?\s?EDUCACI.?N FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"FORMACION\s+RUTA\s+(?:DE\s+)?MODELO\s+DE\s+EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"MODELO\s?FORMACION\s?RUTA\s?DE\s?EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"EDUCACION RUTA DE MODELO DE EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"FORMACIN RUTAS MODELO DE EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"RUTAS DE MODELO DE EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        (r"RUTA DE MODELO DE EDUCACION FLEXIBLE", "EDUCACION FLEXIBLE"),
        # --- Liderazgo: política pública ---
        (r"POLITIC.? PUBLICA DE JUVENTUD.*", "POLITICA PUBLICA DE JUVENTUD"),
        (r"^POLITICA PUBLICA$", "POLITICA PUBLICA DE JUVENTUD"),
        (r"POLITICA PUBLICA JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        (r"PLITICA PUBLICA DE JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        (r"SICIALIZACION DE POLITICA PUBLICA DE JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        (r"POLTICA PUBLICA.*", "POLITICA PUBLICA DE JUVENTUD"),
        (r"LIDERAZGO DE POLITICA PUBLICA JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        (r"USO DE ESPACIOS\s+POLITICA PUBLICA DE JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        (r"^LIDERAZGO DE POLITICA PUBLICA DE JUVENTUD", "POLITICA PUBLICA DE JUVENTUD"),
        # --- Liderazgo juvenil ---
        (r"LIDERE?A?ZGO JUVENIL", "LIDERAZGO JUVENIL"),
        (r"^LIDERAZGO$", "LIDERAZGO JUVENIL"),
        (r"^PARTICIPACION$", "LIDERAZGO JUVENIL"),
        (r"INCIDENCIA PARTICIPACION Y LIDERAZGO", "LIDERAZGO JUVENIL"),
        (r"^Y PRODUCTIVA$", "ORIENTACION SOCIO-OCUPACIONAL"),
        # --- Asesoría jurídica ---
        (r"GESTION CUTURAL", "GESTION CULTURAL"),
        (r"ASES.?RO?I?A JURIDICA Y PARTICIPACI.?N?.*", "ASESORIA JURIDICA Y PARTICIPACION"),
        (r"ASESORIA JURIDICA\b.*", "ASESORIA JURIDICA Y PARTICIPACION"),
        (r"^FUTBOL Y TRANSFORMACION SOCIAL ASESORIAS JURIDICAS", "ASESORIA JURIDICA Y PARTICIPACION"),
    ]:
        t = re.sub(patron, reemplazo, t)
    # Quitar sufijos de detalle
    t = re.split(r"\s+(?:STAND\b|COLEGIO\b|IED\b|ICBF\b|SENA\b|UNAL\b|IDIPRON\b)", t)[0]
    t = t.replace("_", "").replace(".", "").strip()
    t = re.sub(r"\s+", " ", t).strip()
    return t

df["curso_limpio"] = df["NOMBRE_CURSO"].apply(normalizar_curso)
print(f"Cursos normalizados: {df['NOMBRE_CURSO'].nunique()} -> {df['curso_limpio'].nunique()} agrupados")

# ========== celda 10: a3815kgwwz ==========
# Funciones de formato (español colombiano: miles con punto, decimales con coma)
def cap(texto):
    """Mayúscula solo en la primera letra de la frase."""
    if not texto:
        return texto
    t = str(texto).strip().lower()
    siglas = ["spa", "tic", "vbg", "dsdr", "ruv", "lgbtiq+", "srpa", "sirbe", "sdis", "ppdj"]
    for s in siglas:
        t = re.sub(r'\b' + s + r'\b', s.upper(), t, flags=re.IGNORECASE)
    return t[0].upper() + t[1:]

def fmt(n):
    if isinstance(n, float):
        e = int(n); d = round((n - e) * 100)
        return f"{e:,}".replace(",", ".") + f",{d:02d}"
    return f"{n:,}".replace(",", ".")

def pctv(valor):
    return f"{valor:.1f}".replace(".", ",")

# Estadísticas generales (solo Casas de Juventud)
total_atenciones = len(df)
personas_unicas = df["IDPERSONA"].nunique()
n_localidades = df["LOCALIDAD_ATENCION"].nunique()
cursos_por_persona = round(total_atenciones / personas_unicas, 2)

print(f"Total atenciones: {fmt(total_atenciones)}")
print(f"Personas únicas:  {fmt(personas_unicas)}")
print(f"Localidades:      {n_localidades}")
print(f"Cursos/persona:   {str(cursos_por_persona).replace('.', ',')}")

# ========== celda 11: equipo_gen_001 ==========
# Generar sección Equipo desde Excel + imagen de estructura
# El equipo vive en el libro unico datos/equipo.xlsx, hoja "casas_juventud".
equipo_excel = os.path.join(DATOS, "equipo.xlsx")

# Colores fijos por subgrupo
COLORES_SUBGRUPO = {
    "Unidades operativas": "#253C5C",
    "Estrategia móvil": "#c0392b",
    "SIDICU": "#e67e22",
    "Administrativo": "#253C5C",
    "Equipo de ejes del servicio (calidad)": "#2c3e50",
}
DESC_SUBGRUPO = {
    "Equipo de ejes del servicio (calidad)": "Garantizan la calidad técnica del servicio, procesos de formación, diseño de herramientas pedagógicas y ejecución de eventos.",
}
# Subgrupos que usan el estilo "botón apilado" rojo con highlight amarillo
# (igual al esquema visual de Casas de Juventud con ícono + offset amarillo)
SUBGRUPOS_ESTILO_ROJO = {"Unidades operativas", "Estrategia móvil"}
# Grupos que usan el estilo oscuro con highlight amarillo (ejes del servicio)
GRUPOS_ESTILO_OSCURO = {"Equipo transversal"}

if os.path.exists(equipo_excel):
    df_equipo = pd.read_excel(equipo_excel, sheet_name="casas_juventud")

    # Reorganización del equipo:
    # - Renombrar el grupo de modalidades para simplificar el título
    # - Mover SIDICU al "Equipo transversal" (se mostrará como tarjeta horizontal inferior)
    # - Reordenar para que SIDICU quede al final dentro del Equipo transversal
    df_equipo["Grupo"] = df_equipo["Grupo"].replace({
        "Modalidades del servicio (implementación en territorio)": "Modalidades del servicio",
    })
    df_equipo.loc[df_equipo["Subgrupo"] == "SIDICU", "Grupo"] = "Equipo transversal"
    mask_sidicu = df_equipo["Subgrupo"] == "SIDICU"
    df_equipo = pd.concat([df_equipo[~mask_sidicu], df_equipo[mask_sidicu]], ignore_index=True)

    equipo_html = '<div class="card">\n'
    equipo_html += '                <h2 class="card-title">Equipo</h2>\n'
    equipo_html += '                <p style="color:#666; margin-bottom:20px;">Estructura organizacional de Casas de Juventud.</p>\n'

    # Imagen de estructura del servicio
    equipo_html += '                <div style="text-align:center; margin:20px 0 30px; background:#fff; border-radius:12px; padding:20px;">\n'
    equipo_html += '                    <img src="imagenes/servicios%20sdis%20juventud.png" alt="Estructura del servicio" style="max-width:100%;">\n'
    equipo_html += '                </div>\n'

    # 1. Coordinadora (desde Excel)
    df_coord = df_equipo[df_equipo["Grupo"] == "Coordinación"]
    if not df_coord.empty:
        coord_nombre = df_coord.iloc[0]["Nombre"]
        coord_cargo = df_coord.iloc[0]["Cargo"]
        equipo_html += f'                <div style="background:#2B2F3A; color:#fff; border-radius:10px; padding:18px 25px; text-align:center; margin-bottom:25px;">\n'
        equipo_html += f'                    <div style="font-size:0.75rem; text-transform:uppercase; letter-spacing:1.5px; opacity:0.8; margin-bottom:4px;">1. {coord_cargo}</div>\n'
        equipo_html += f'                    <div style="font-size:1.15rem; font-weight:700; color:#fff;">{coord_nombre}</div>\n'
        equipo_html += '                </div>\n'
        df_equipo = df_equipo[df_equipo["Grupo"] != "Coordinación"]


    # Generar secciones por grupo
    grupo_num = 2
    for grupo in df_equipo["Grupo"].unique():
        sub_grupo = df_equipo[df_equipo["Grupo"] == grupo]
        equipo_html += f'\n                <h3 class="card-subtitle">{grupo_num}. {grupo}</h3>\n'
        # Grid de tarjetas — más gap cuando haya tarjetas con offset amarillo
        hay_estilo_rojo = any(s in SUBGRUPOS_ESTILO_ROJO for s in sub_grupo["Subgrupo"].unique())
        estilo_oscuro_grupo = grupo in GRUPOS_ESTILO_OSCURO
        hay_offset = hay_estilo_rojo or estilo_oscuro_grupo
        gap_grid = "22px" if hay_offset else "16px"
        mb_grid = "32px" if hay_offset else "25px"
        # En "Equipo transversal" usamos 2 columnas fijas para poder poner una
        # tarjeta horizontal (SIDICU) como tercera, que ocupa toda la fila.
        if estilo_oscuro_grupo:
            grid_cols = "repeat(2, 1fr)"
        else:
            grid_cols = "repeat(auto-fit, minmax(250px, 1fr))"
        equipo_html += f'                <div style="display:grid; grid-template-columns:{grid_cols}; gap:{gap_grid}; margin-bottom:{mb_grid};">\n'

        sub_num = 1
        subgrupos_lista = list(sub_grupo["Subgrupo"].unique())
        for idx_sub, subgrupo in enumerate(subgrupos_lista):
            personas = sub_grupo[sub_grupo["Subgrupo"] == subgrupo]
            color = COLORES_SUBGRUPO.get(subgrupo, "#253C5C")
            desc = DESC_SUBGRUPO.get(subgrupo, "")
            estilo_rojo = subgrupo in SUBGRUPOS_ESTILO_ROJO
            estilo_oscuro = estilo_oscuro_grupo

            if estilo_rojo:
                # Tarjeta rojo-vino con offset amarillo (estilo Casas de Juventud)
                card_style = "background:#9B1C1C; border-radius:12px; padding:20px; box-shadow:7px 7px 0 #F5B800; color:#fff;"
                title_color = "#F5B800"
                title_weight = "; font-weight:700"
                name_size = "0.95rem"
                cargo_color = "rgba(255,255,255,0.75)"
                desc_color = "rgba(255,255,255,0.8)"
                badge_style = "padding:3px 9px; background:transparent; border:1px solid rgba(255,255,255,0.45); color:#fff; border-radius:5px; font-size:0.72rem;"
                obs_color = "rgba(255,255,255,0.6)"
            elif estilo_oscuro:
                # Tarjeta oscura con offset amarillo (ejes del servicio)
                card_style = "background:#2B2F3A; border-radius:12px; padding:20px; box-shadow:7px 7px 0 #F5B800; color:#fff;"
                # La tercera tarjeta en adelante ocupa toda la fila (horizontal)
                if idx_sub >= 2:
                    card_style += " grid-column:1 / -1;"
                title_color = "#F5B800"
                title_weight = "; font-weight:700"
                name_size = "0.95rem"
                cargo_color = "rgba(255,255,255,0.75)"
                desc_color = "rgba(255,255,255,0.8)"
                badge_style = "padding:3px 9px; background:transparent; border:1px solid rgba(255,255,255,0.45); color:#fff; border-radius:5px; font-size:0.72rem;"
                obs_color = "rgba(255,255,255,0.6)"
            else:
                card_style = f"border-top:3px solid {color}; border-radius:10px; padding:18px; background:#f8f9fa; box-shadow:0 1px 4px rgba(0,0,0,0.06);"
                title_color = color
                title_weight = ""
                name_size = "0.9rem"
                cargo_color = "#888"
                desc_color = "#888"
                badge_style = "padding:6px 12px; background:#e8ecf1; border-radius:6px; font-size:0.82rem;"
                obs_color = "#aaa"

            equipo_html += f'                    <div style="{card_style}">\n'
            equipo_html += f'                        <h4 style="font-size:{name_size}; color:{title_color}; margin:0 0 {"4px" if desc else "12px"}{title_weight};">{grupo_num}.{sub_num} {subgrupo}</h4>\n'
            if desc:
                equipo_html += f'                        <p style="font-size:0.8rem; color:{desc_color}; margin:0 0 12px;">{desc}</p>\n'

            # Personas con nombre
            for _, p in personas.iterrows():
                nombre = str(p["Nombre"]).strip() if pd.notna(p["Nombre"]) and str(p["Nombre"]).strip() else ""
                cargo = str(p["Cargo"]).strip() if pd.notna(p["Cargo"]) else ""
                numero = str(int(p["Numero"])) if pd.notna(p.get("Numero")) and str(p["Numero"]).strip() not in ["", "nan"] else ""
                obs = str(p["Observaciones"]).strip() if pd.notna(p.get("Observaciones")) and str(p["Observaciones"]).strip() not in ["", "nan"] else ""

                if nombre:
                    equipo_html += f'                        <div style="padding:5px 0;"><strong>{nombre}</strong> <span style="color:{cargo_color}; font-size:0.85rem;">— {cargo}</span>'
                    if obs:
                        equipo_html += f' <span style="color:{obs_color}; font-size:0.8rem;">({obs})</span>'
                    equipo_html += '</div>\n'

            # Conteos (badges)
            badges = []
            for _, p in personas.iterrows():
                nombre = str(p["Nombre"]).strip() if pd.notna(p["Nombre"]) and str(p["Nombre"]).strip() else ""
                if not nombre:
                    cargo = str(p["Cargo"]).strip() if pd.notna(p["Cargo"]) else ""
                    numero = str(int(p["Numero"])) if pd.notna(p.get("Numero")) and str(p["Numero"]).strip() not in ["", "nan"] else ""
                    obs = str(p["Observaciones"]).strip() if pd.notna(p.get("Observaciones")) and str(p["Observaciones"]).strip() not in ["", "nan"] else ""
                    if numero:
                        badge_text = f'<strong>{numero}</strong> {cargo}'
                        if obs:
                            badge_text += f' ({obs})'
                        badges.append(badge_text)

            if badges:
                mt_badges = "12px" if (estilo_rojo or estilo_oscuro) else "10px"
                equipo_html += f'                        <div style="margin-top:{mt_badges}; display:flex; gap:8px; flex-wrap:wrap;">\n'
                for b in badges:
                    equipo_html += f'                            <span style="{badge_style}">{b}</span>\n'
                equipo_html += '                        </div>\n'

            equipo_html += '                    </div>\n'
            sub_num += 1

        equipo_html += '                </div>\n'
        grupo_num += 1

    equipo_html += '            </div>'
    print(f"Equipo generado desde Excel: {len(df_equipo)} filas")
else:
    equipo_html = '<div class="card"><h2 class="card-title">Equipo</h2><p>No se encontró equipo_casas_juventud.xlsx</p></div>'
    print("Excel de equipo no encontrado")

# ========== celda 12: directorio_gen_001 ==========
# Generar directorio de Casas de Juventud desde Excel
directorio_excel = os.path.join(DATOS, "directorio_casas_juventud.xlsx")

if os.path.exists(directorio_excel):
    df_dir = pd.read_excel(directorio_excel)
    
    directorio_html = '<div style="margin-top:25px;">\n'
    directorio_html += '                <h3 class="card-subtitle">Directorio</h3>\n'
    directorio_html += '                <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">\n'
    directorio_html += '                    <thead>\n'
    directorio_html += '                        <tr>\n'
    directorio_html += '                            <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; text-align:left; font-weight:700; border-bottom:none;">Casa de Juventud</th>\n'
    directorio_html += '                            <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; text-align:left; font-weight:700; border-bottom:none;">Localidad</th>\n'
    directorio_html += '                            <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; text-align:left; font-weight:700; border-bottom:none;">Dirección</th>\n'
    directorio_html += '                            <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; text-align:left; font-weight:700; border-bottom:none;">Barrio</th>\n'
    directorio_html += '                            <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; text-align:left; font-weight:700; border-bottom:none;">Correo de contacto</th>\n'
    directorio_html += '                        </tr>\n'
    directorio_html += '                    </thead>\n'
    directorio_html += '                    <tbody>\n'

    for idx, row in df_dir.iterrows():
        bg = '#fafafa' if idx % 2 == 0 else '#fff'
        # Si la fila tiene link de Google Maps, envuelve la dirección en <a>
        link_maps = row.get("Link Google Maps", "")
        if pd.notna(link_maps) and str(link_maps).strip():
            direccion_html = f'<a href="{link_maps}" target="_blank" style="color:#253C5C;">{row["Dirección"]}</a>'
        else:
            direccion_html = str(row["Dirección"])
        # Correo (con mailto si esta presente)
        correo = row.get("Correo", "")
        if pd.notna(correo) and str(correo).strip():
            correo_html = f'<a href="mailto:{correo}" style="color:#253C5C;">{correo}</a>'
        else:
            correo_html = '<span style="color:#999;">—</span>'
        directorio_html += f'                        <tr style="background:{bg}; border-bottom:1px solid #e0e0e0;">\n'
        directorio_html += f'                            <td style="padding:12px 14px; vertical-align:top;"><strong>{row["Casa de Juventud"]}</strong></td>\n'
        directorio_html += f'                            <td style="padding:12px 14px; vertical-align:top;">{row["Localidad"]}</td>\n'
        directorio_html += f'                            <td style="padding:12px 14px; vertical-align:top;">{direccion_html}</td>\n'
        directorio_html += f'                            <td style="padding:12px 14px; vertical-align:top;">{row["Barrio"]}</td>\n'
        directorio_html += f'                            <td style="padding:12px 14px; vertical-align:top; word-break:break-all;">{correo_html}</td>\n'
        directorio_html += '                        </tr>\n'
    
    directorio_html += '                    </tbody>\n'
    directorio_html += '                </table>\n'
    directorio_html += '            </div>'
    print(f"Directorio generado desde Excel: {len(df_dir)} casas")
else:
    directorio_html = '<p style="color:#888;">No se encontró directorio_casas_juventud.xlsx</p>'
    print("Excel de directorio no encontrado")

# ========== celda 13: mapa_gen_001 ==========
import json
# Generar mapa interactivo de Casas de Juventud
import folium
import unicodedata
import re

geojson_path = os.path.join(DATOS, "localidades_bogota.geojson")

def normalizar(texto):
    """Normaliza nombres: quita tildes, Ñ->N, prefijos LA/LOS/EL"""
    texto = unicodedata.normalize('NFC', texto).upper().strip()
    texto = texto.replace('Ñ', 'N').replace('ñ', 'n')
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'^(LA |LOS |EL )', '', texto)
    return texto

if os.path.exists(directorio_excel) and os.path.exists(geojson_path):
    df_mapa = pd.read_excel(directorio_excel)
    with open(geojson_path, encoding="utf-8") as f:
        localidades_gj = json.load(f)

    locs_con_casa = set(normalizar(l) for l in df_mapa["Localidad"].unique())

    m = folium.Map(location=[4.624, -74.105], zoom_start=11, tiles="CartoDB positron", width="100%", height="100%")

    def style_localidad(feature):
        nombre = normalizar(feature["properties"]["nombre"])
        if nombre in locs_con_casa:
            return {"fillColor": "#d5d5d5", "color": "#253C5C", "weight": 1.5, "fillOpacity": 0.35}
        else:
            return {"fillColor": "#ffffff", "color": "#ccc", "weight": 1, "fillOpacity": 0.1}

    folium.GeoJson(localidades_gj, name="Localidades", style_function=style_localidad,
        tooltip=folium.GeoJsonTooltip(fields=["nombre"], aliases=["Localidad:"])).add_to(m)

    for _, row in df_mapa.iterrows():
        if pd.notna(row.get("Latitud")) and pd.notna(row.get("Longitud")):
            popup_html = f'<div style="font-family:Arial; min-width:200px;"><strong style="color:#253C5C; font-size:14px;">{row["Casa de Juventud"]}</strong><br><span style="color:#666; font-size:12px;">📍 {row["Localidad"]}</span><br><span style="font-size:11px;">{row["Dirección"]}</span><br><span style="font-size:11px; color:#888;">Barrio: {row["Barrio"]}</span></div>'
            folium.CircleMarker(location=[row["Latitud"], row["Longitud"]], radius=7, color="#253C5C",
                fill=True, fill_color="#c0392b", fill_opacity=0.9, weight=2,
                popup=folium.Popup(popup_html, max_width=250), tooltip=row["Casa de Juventud"]).add_to(m)

    ruta_mapa = os.path.join(BASE, "mapa_casas_juventud.html")
    m.save(ruta_mapa)

    # Folium asigna IDs hexadecimales aleatorios (32 chars) al mapa y a cada
    # capa en cada corrida. Eso ensucia el git diff con cambios cosmeticos.
    # Aqui los reemplazamos por IDs estables segun orden de aparicion, asi
    # ejecutar el script con los mismos datos da el mismo HTML byte a byte.
    import re
    with open(ruta_mapa, "r", encoding="utf-8") as f:
        html_mapa = f.read()
    # Captura "prefijo_<32 hex>" donde el prefijo no esta pegado a un identificador
    # anterior. No usamos \b al final porque algunos identificadores tienen
    # sufijos como "_styler" o "_add" pegados al hex.
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

    print(f"Mapa generado con {len(df_mapa)} casas, {len(locs_con_casa)} localidades destacadas")
else:
    print("Faltan archivos para generar mapa")

# ========== celda 14: 7ba9i7apdl ==========
# Modalidades
def limpiar_mod(nom):
    if pd.isna(nom): return "Sin modalidad"
    return cap(re.sub(r"^\d+\s+", "", str(nom)))

df["ml"] = df["NOMMODALIDAD"].apply(limpiar_mod)
modalidades_casas = df.groupby("ml").size().sort_values(ascending=False)

# Localidades
loc_stats = df.groupby("LOCALIDAD_ATENCION").size().sort_values(ascending=False)

# Demografía
sexo = df["SEXO"].value_counts()
edad = df["GRUPO_ETAREO_1DIA_ANUAL_ACTUA"].value_counts()
etnia = df["ETNIA_ACTUAL"].value_counts()
ruv = df["RUV_ACTUAL"].value_counts()
disc = df["SITUACION_DISCAPACIDAD_ACTUAL"].value_counts(dropna=False)

print("Modalidades Casas:")
print(modalidades_casas)
print(f"\nSexo: {dict(sexo)}")
print(f"Edad: {dict(edad)}")

# ========== celda 16: eqlou1jrbs ==========
# Información de cada eje (imágenes locales, descripciones, actuaciones SIRBE)

EJES_INFO = {
    "BIENESTAR": {
        "id": "bienestar",
        "imagen": "imagenes/bienestar.png",
        "desc": "Acciones de prevencion integral que incluyen prevencion de consumo de SPA, prevencion de violencias basadas en genero VBG, derechos sexuales y reproductivos DDSSYR, salud mental y orientacion psicosocial.",
        "actuacion": "Prevencion Integral (464) — Ruta: 12",
    },
    "CULTURA": {
        "id": "cultura",
        "imagen": "imagenes/cultura.png",
        "desc": "Formacion artistica, uso de espacios para practicas artisticas y culturales, deportes, eventos y aprovechamiento del tiempo libre con enfasis en intereses juveniles.",
        "actuacion": "Manejo Adecuado de Tiempo Libre (1003) — Ruta: 10",
    },
    "INCLUSIÓN SOCIAL Y PRODUCTIVA": {
        "id": "inclusion",
        "imagen": "imagenes/inclusion.png",
        "desc": "Formacion para el proyecto de vida, emprendimiento y empleabilidad, educacion financiera, orientacion socio-ocupacional y habilidades TIC para la inclusion laboral.",
        "actuacion": "Formacion para el Proyecto de Vida (1005) — Ruta: 8",
    },
    "LIDERAZGO Y PARTICIPACIÓN": {
        "id": "liderazgo",
        "imagen": "imagenes/liderazgo.png",
        "desc": "Socializacion de politica publica de juventud, asesoria juridica, liderazgo juvenil, derechos humanos, voluntariado intergeneracional y participacion ciudadana.",
        "actuacion": "Asesoria Juridica y Participacion (1004), Politica Publica de Juventud (739), Voluntariado Intergeneracional (1007) — Ruta: 5",
    },
}

ORDEN_EJES = ["BIENESTAR", "CULTURA", "INCLUSIÓN SOCIAL Y PRODUCTIVA", "LIDERAZGO Y PARTICIPACIÓN"]
print("Ejes configurados:", len(EJES_INFO))

# ========== celda 17: ra5omu7qguf ==========
# Generar quick links que abren los HTMLs externos de cada eje
sidebar_ejes = ""
quick_links = ""

ejes_html = {
    "BIENESTAR": "ejes/Bienestar.html",
    "CULTURA": "ejes/Cultura.html",
    "INCLUSIÓN SOCIAL Y PRODUCTIVA": "ejes/Inclusion.html",
    "LIDERAZGO Y PARTICIPACIÓN": "ejes/Liderazgo.html",
}

for eje in ORDEN_EJES:
    info = EJES_INFO[eje]
    href = ejes_html[eje]
    quick_links += f"""                        <a class="quick-link" href="{href}" style="text-decoration:none; color:inherit;">
                            <div class="quick-link-img"><img src="{info['imagen']}" alt="{cap(eje)}"></div>
                            <div class="quick-link-title">{cap(eje)}</div>
                        </a>\n"""

print("Quick links generados (apuntan a HTMLs externos)")

# ========== celda 18: euovll5uqkb ==========
# Generar secciones auxiliares + datos JSON para filtros interactivos
import json as _json

# --- Modalidades ---
mcH = ""
for m, n in modalidades_casas.items():
    mcH += f'                        <div class="list-item"><div><strong>{m}</strong></div><span class="badge badge-primary">{fmt(n)} atenciones ({pctv(n/total_atenciones*100)}%)</span></div>\n'

ejes_badges = ""
for eje in ORDEN_EJES:
    n_eje = (df["eje"] == eje).sum()
    p = n_eje / total_atenciones * 100
    bc = "badge-success" if p >= 10 else "badge-info"
    ejes_badges += f'                        <span class="badge {bc}">{cap(eje)} ({pctv(p)}%)</span>\n'

# --- Corrección tildes localidades ---
tildes_loc = {
    "Ciudad bolivar": "Ciudad Bol\u00edvar",
    "Engativa": "Engativ\u00e1",
    "Fontibon": "Fontib\u00f3n",
    "San cristobal": "San Crist\u00f3bal",
    "Antonio nari\u00f1o": "Antonio Nari\u00f1o",
    "Rafael uribe": "Rafael Uribe Uribe",
    "Los martires": "Los M\u00e1rtires",
    "Barrios unidos": "Barrios Unidos",
    "Puente aranda": "Puente Aranda",
}

def nombre_loc(loc):
    c = cap(loc)
    return tildes_loc.get(c, c)

# === DATOS JSON PARA FILTROS INTERACTIVOS ===
# Localidades por eje
datos_loc = {}
# Total
loc_total = df.groupby("LOCALIDAD_ATENCION").size().reset_index(name="atenciones").sort_values("atenciones", ascending=False).set_index("LOCALIDAD_ATENCION")
datos_loc["Total"] = [
    {"loc": nombre_loc(l), "at": int(r["atenciones"])}
    for l, r in loc_total.iterrows()
]
# Por eje
for eje in ORDEN_EJES:
    sub = df[df["eje"] == eje]
    loc_eje = sub.groupby("LOCALIDAD_ATENCION").size().reset_index(name="atenciones").sort_values("atenciones", ascending=False).set_index("LOCALIDAD_ATENCION")
    ult_eje = sub.drop_duplicates(subset="IDPERSONA", keep="last")
    pe_eje = ult_eje.groupby("LOCALIDAD_ATENCION").size()
    datos_loc[cap(eje)] = [
        {"loc": nombre_loc(l), "at": int(r["atenciones"])}
        for l, r in loc_eje.iterrows()
    ]

# Demografía por eje (edad, sexo, etnia, ruv, discapacidad)
datos_demo = {}
for filtro in ["Total"] + [cap(e) for e in ORDEN_EJES]:
    sub = df if filtro == "Total" else df[df["eje"] == [e for e in ORDEN_EJES if cap(e) == filtro][0]]
    n = len(sub)
    if n == 0:
        continue
    datos_demo[filtro] = {
        "edad": [{"cat": str(g), "n": int(v)} for g, v in sub["GRUPO_ETAREO_1DIA_ANUAL_ACTUA"].value_counts().items()],
        "sexo": [{"cat": cap(str(g)), "n": int(v)} for g, v in sub["SEXO"].value_counts().items()],
        "etnia": [{"cat": cap(str(g)), "n": int(v)} for g, v in sub["ETNIA_ACTUAL"].value_counts().items()],
        "ruv": [{"cat": "Registrados en RUV" if g == "SI" else "No registrados", "n": int(v)} for g, v in sub["RUV_ACTUAL"].value_counts().items()],
        "disc": [{"cat": cap(str(g)) if pd.notna(g) else "Sin informacion", "n": int(v)} for g, v in sub["SITUACION_DISCAPACIDAD_ACTUAL"].value_counts(dropna=False).items()],
        "total": n,
    }

datos_loc_json = _json.dumps(datos_loc, ensure_ascii=False)
datos_demo_json = _json.dumps(datos_demo, ensure_ascii=False)

# --- Localidades HTML estático (para la tabla inicial) ---
loc_at = df.groupby("LOCALIDAD_ATENCION").size().reset_index(name="atenciones")
loc_df = loc_at.copy()
loc_df.columns = ["localidad", "atenciones"]
loc_df = loc_df.sort_values("atenciones", ascending=False).reset_index(drop=True)
loc_df["pct"] = loc_df["atenciones"] / total_atenciones * 100
loc_df["localidad_fmt"] = loc_df["localidad"].apply(nombre_loc)

t3 = loc_df.head(3)["localidad_fmt"].tolist()
t3p = loc_df.head(3)["pct"].sum()
b3 = loc_df.tail(3)["localidad_fmt"].tolist()

# --- Resumen ---
ejes_resumen = ""
for eje in ORDEN_EJES:
    n = (df["eje"] == eje).sum()
    p = n / total_atenciones * 100
    pers = df[df["eje"] == eje]["IDPERSONA"].nunique()
    ejes_resumen += f'                            <tr><td>{cap(eje)}</td><td>{fmt(n)}</td><td>{pctv(p)}%</td><td>{fmt(pers)}</td></tr>\n'

# --- Brechas ---
sin_clasificar = df[df["eje"] == "SIN CLASIFICAR"]
n_sin = len(sin_clasificar)
pct_sin = n_sin / total_atenciones * 100

sinH = ""
if n_sin > 0:
    for cod, n in sin_clasificar["NOMACTIVIDAD_CURSO"].value_counts().items():
        sinH += f'                            <tr><td>{cod}</td><td>{fmt(n)}</td><td>{pctv(n/n_sin*100)}%</td></tr>\n'

# --- Enfoque diferencial ---
difH = ""
for lab, n, p in [
    ("Personas con discapacidad", (df["SITUACION_DISCAPACIDAD_ACTUAL"]=="SI").sum(), (df["SITUACION_DISCAPACIDAD_ACTUAL"]=="SI").sum()/total_atenciones*100),
    ("Poblacion afrodescendiente", df["ETNIA_ACTUAL"].str.contains("AFRO|NEGRO",case=False,na=False).sum(), df["ETNIA_ACTUAL"].str.contains("AFRO|NEGRO",case=False,na=False).sum()/total_atenciones*100),
    ("Poblacion indigena", df["ETNIA_ACTUAL"].str.contains("INDIGENA",case=False,na=False).sum(), df["ETNIA_ACTUAL"].str.contains("INDIGENA",case=False,na=False).sum()/total_atenciones*100),
    ("Poblacion LGBTIQ+", df["NOMORIENTACION_ACTUAL"].str.contains("HOMO|BISEX|LESB|GAY|PANSES",case=False,na=False).sum(), df["NOMORIENTACION_ACTUAL"].str.contains("HOMO|BISEX|LESB|GAY|PANSES",case=False,na=False).sum()/total_atenciones*100),
]:
    difH += f'                            <tr><td>{lab}</td><td>{fmt(n)}</td><td>{pctv(p)}%</td></tr>\n'

print("Datos JSON generados:")
print(f"  Localidades: {len(datos_loc)} filtros, {len(datos_loc['Total'])} localidades")
print(f"  Demografía: {len(datos_demo)} filtros")
print("Secciones auxiliares generadas")

# ========== celda 19: mbd92b0e1u ==========
# CSS + ensamblar HTML completo + guardar archivo

CSS = """        @import url('https://fonts.googleapis.com/css2?family=Anton&family=Antonio:wght@400;700&family=Figtree:wght@400;500;600;700;800&display=swap');
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Figtree', 'Segoe UI', sans-serif; background-color: #ffffff; color: #2F3E3C; }
        .header { background: #2F3E3C; color: #F8F4E1; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.15); }
        .header h1 { font-size: 1.5rem; font-weight: 700; } .header .subtitle { font-size: 0.9rem; opacity: 0.85; }
        .home-btn { font-size: 1.5rem; cursor: pointer; padding: 5px 12px; border-radius: 8px; transition: background 0.2s; } .home-btn:hover { background: rgba(255,255,255,0.15); }
        .container { display: flex; min-height: calc(100vh - 80px); }
        .sidebar { width: 280px; background: #fff; border-right: 1px solid #e0e0e0; padding: 20px 0; overflow-y: auto; }
        .sidebar-section { margin-bottom: 10px; }
        .sidebar-title { padding: 12px 20px; font-weight: 600; color: #253C5C; cursor: pointer; display: flex; justify-content: space-between; align-items: center; transition: background 0.2s; }
        .sidebar-title:hover { background: transparent; opacity: 0.75; }
        .sidebar-title .arrow { display: inline-block; width: 8px; height: 8px; border-top: 2px solid #253C5C; border-right: 2px solid #253C5C; transform: rotate(45deg); font-size: 0; color: transparent; transition: transform 0.25s ease; margin-right: 4px; }
        .sidebar-title.active .arrow { transform: rotate(135deg); margin-top: -4px; }
        .sidebar-items { display: none; padding-left: 20px; } .sidebar-items.show { display: block; }
        .sidebar-item { padding: 10px 20px; cursor: pointer; font-size: 0.9rem; color: #3A3A3A; transition: all 0.2s; }
        .sidebar-item:hover { background: transparent; color: #253C5C; }
        .sidebar-item.active { background: transparent; color: #253C5C; font-weight: 600; }
        .main-content { flex: 1; padding: 30px; overflow-y: auto; }
        .content-section { display: none; } .content-section.active { display: block; }
        .card { background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); padding: 25px; margin-bottom: 20px; }
        .card-title { font-size: 1.4rem; color: #253C5C; margin-bottom: 15px; }
        /* Subt&iacute;tulos en Antonio Bold may&uacute;sculas (consistente con Forjar/JCO/Bienestar/Parche seguro). */
        .card-subtitle { font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-size: 1.05rem; color: #3A5275; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; margin: 25px 0 10px 0; }
        /* Tabla comparativa Modelo anterior vs Nuevo modelo 2026 ("A tener en cuenta").
           Tres columnas en desktop con headers en p&iacute;ldoras de color, &iacute;conos Lucide
           en cada celda, fondo crema. Mano halftone decorativa arriba a la izquierda.
           En m&oacute;vil colapsa a una columna donde cada caracter&iacute;stica se vuelve un
           bloque con dos sub-cards (Anterior / Nuevo 2026) etiquetadas con badge de color
           para que el contraste no se pierda. */
        .comp-tabla { position: relative; background: #f5efd2; border-radius: 14px; padding: 24px 24px 18px; margin: 32px 0 28px; box-shadow: 0 2px 10px rgba(0,0,0,0.04); overflow: visible; }
        .comp-mano-decor { position: absolute; top: -38px; left: -32px; width: 120px; height: auto; transform: rotate(-14deg); pointer-events: none; z-index: 2; }
        .comp-tabla-grid { display: grid; grid-template-columns: 1.1fr 1.4fr 1.4fr; gap: 14px 22px; align-items: center; }
        .comp-header { display: flex; align-items: center; justify-content: center; padding: 10px 16px; border-radius: 22px; font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.8rem; line-height: 1.2; text-align: center; color: #F8F4E1; }
        .comp-header-1 { background: #2F3E3C; }
        .comp-header-2 { background: #5B7E9D; }
        .comp-header-3 { background: #253C5C; }
        .comp-row-sep { grid-column: 1 / -1; border-top: 1px solid rgba(47, 62, 60, 0.12); margin: 6px 0; }
        .comp-caracteristica { font-family: 'Figtree', 'Segoe UI', sans-serif; font-weight: 700; color: #2F3E3C; font-size: 0.95rem; line-height: 1.3; padding: 10px 4px; }
        /* Desktop: celdas sin fondo, contraste por color de icono y headers de columna.
           M&oacute;vil: el contraste se da por solidez (Anterior blanca, Nuevo s&oacute;lido azul
           Casas con texto crema). Evita el borde lateral coloreado (tic AI). */
        .comp-celda { display: flex; align-items: center; gap: 14px; padding: 10px 4px; font-family: 'Figtree', 'Segoe UI', sans-serif; font-size: 0.9rem; color: #2F3E3C; line-height: 1.45; }
        .comp-icono { flex-shrink: 0; width: 56px; height: 56px; border-radius: 50%; display: flex; align-items: center; justify-content: center; }
        .comp-icono svg { width: 28px; height: 28px; stroke-width: 1.7; }
        .comp-anterior .comp-icono { background: rgba(91, 126, 157, 0.18); color: #5B7E9D; }
        .comp-nuevo .comp-icono { background: rgba(37, 60, 92, 0.15); color: #253C5C; }
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
            .comp-celda { padding: 12px 14px; margin-bottom: 8px; align-items: flex-start; border-radius: 10px; }
            .comp-anterior { background: #ffffff; color: #2F3E3C; }
            .comp-nuevo { background: #253C5C; color: #F8F4E1; }
            .comp-nuevo .comp-icono { background: rgba(248, 244, 225, 0.15); color: #F8F4E1; }
            .comp-celda-badge { display: inline-block; font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.62rem; padding: 2px 8px; border-radius: 10px; color: #F8F4E1; margin-bottom: 4px; }
            .comp-anterior .comp-celda-badge { background: #5B7E9D; }
            .comp-nuevo .comp-celda-badge { background: rgba(248, 244, 225, 0.2); color: #F8F4E1; }
            .comp-icono { width: 40px; height: 40px; }
            .comp-icono svg { width: 20px; height: 20px; }
        }
        .card p { margin-bottom: 15px; }
        .badge { display: inline-block; padding: 5px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 500; margin-right: 8px; margin-bottom: 8px; }
        .badge-primary { background: #DFE6F0; color: #253C5C; } .badge-success { background: #e6f9f0; color: #1EAE76; }
        .badge-warning { background: #FFF3E0; color: #F58B53; } .badge-info { background: #e0f4f5; color: #1E9DA3; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 20px 0; }
        .stat-card { background: linear-gradient(135deg, #253C5C 0%, #5A708E 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; }
        .stat-card.alt { background: linear-gradient(135deg, #1E9DA3 0%, #1EAE76 100%); }
        .stat-card.alt2 { background: linear-gradient(135deg, #F4676E 0%, #F58B53 100%); }
        .stat-number { font-size: 2rem; font-weight: 700; } .stat-label { font-size: 0.9rem; opacity: 0.9; }
        .list-group { border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; }
        .list-item { padding: 15px 20px; border-bottom: 1px solid #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .list-item:last-child { border-bottom: none; } .list-item:hover { background: #faf8f2; }
        .progress-bar { height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; width: 100px; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #253C5C, #F4676E); border-radius: 4px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #e0e0e0; }
        th { background: #faf8f2; font-weight: 600; color: #253C5C; } tr:hover { background: #faf8f2; }
        #homologacion tr:hover { background: transparent; }
        .methodology-box { background: #EAEFF5; border-left: 4px solid #253C5C; padding: 15px 20px; margin: 15px 0; border-radius: 0 8px 8px 0; }
        .materials-list { display: flex; flex-wrap: wrap; gap: 10px; margin: 10px 0; }
        .welcome-section { text-align: center; padding: 60px 20px; }
        .welcome-section h2 { font-size: 2rem; color: #253C5C; margin-bottom: 20px; }
        .welcome-section p { font-size: 1.1rem; color: #666; max-width: 600px; margin: 0 auto 30px; }
        .quick-links { display: grid; grid-template-columns: repeat(2, 190px); justify-content: center; gap: 30px; }
        .quick-link { background: transparent; padding: 25px 20px; border-radius: 16px; text-align: center; cursor: pointer; transition: transform 0.3s; width: 190px; }
        .quick-link:hover { transform: translateY(-6px); }
        .quick-link-img { width: 120px; height: 120px; margin: 0 auto 15px; border-radius: 50%; overflow: hidden; border: 3px solid #DFE6F0; box-shadow: 0 4px 15px rgba(102,58,147,0.15); transition: border-color 0.3s, transform 0.3s; }
        .quick-link:hover .quick-link-img { border-color: #253C5C; transform: scale(1.05); }
        .quick-link:hover .sidicu-img { border-color: #253C5C; }
        .quick-link-img img { width: 100%; height: 100%; object-fit: cover; }
        .quick-link-title { font-family: 'Anton', 'Figtree', sans-serif; font-weight: 400; font-size: 1.05rem; line-height: 1.15; letter-spacing: 0.5px; background: #5d5e95; color: #c7eb91; padding: 9px 14px 7px; display: inline-block; text-align: center; max-width: 100%; }
        .hbar-row { display: flex; align-items: center; padding: 8px 0; border-bottom: 1px solid #f0f0f0; }
        .hbar-row:last-child { border-bottom: none; }
        .hbar-label { width: 200px; font-size: 0.9rem; color: #3A3A3A; flex-shrink: 0; }
        .hbar-track { flex: 1; height: 22px; background: #E4EAF1; border-radius: 11px; overflow: hidden; margin: 0 15px; }
        .hbar-fill { height: 100%; border-radius: 11px; transition: width 0.5s ease; min-width: 2px; }
        .hbar-value { width: 150px; font-size: 0.85rem; color: #555; text-align: right; flex-shrink: 0; }
        .hbar-pct { color: #999; }
        .hbar-group { background: #fff; border: 1px solid #e0e0e0; border-radius: 10px; padding: 15px 20px; margin: 10px 0; }
        .filter-btns { display: flex; gap: 8px; flex-wrap: wrap; margin: 15px 0 20px; }
        .filter-btn { padding: 8px 18px; border: 2px solid #253C5C; border-radius: 25px; background: white; color: #253C5C; font-family: inherit; font-size: 0.85rem; font-weight: 600; cursor: pointer; transition: all 0.2s; }
        .filter-btn:hover { background: #EAEFF5; }
        .filter-btn.active { background: #253C5C; color: white; }
        /* --- Estructura SIRBE tipo arbol --- */
        .tree-level { margin-left: 25px; padding: 8px 0; }
        .tree-label { font-weight: 600; color: #4A5A78; }
        .tree-sublabel { color: #666; font-size: 0.9rem; margin-left: 10px; }
        .tree-item {
            padding: 6px 15px;
            border-left: 2px solid #DFE6F0;
            margin-left: 20px;
            margin-top: 4px;
        }
        @media (max-width: 768px) { .container { flex-direction: column; } .sidebar { width: 100%; border-right: none; border-bottom: 1px solid #e0e0e0; } .quick-link-img { width: 90px; height: 90px; } }
        /* Flujo de gestión: pasos en columna con ícono redondo a la izquierda y conector punteado entre ellos. */
        .flujo-pasos { display: flex; flex-direction: column; gap: 28px; margin: 22px 0 8px; }
        .flujo-paso { display: grid; grid-template-columns: 64px 1fr; gap: 22px; align-items: flex-start; position: relative; }
        .flujo-paso:not(:last-child)::before { content: ''; position: absolute; left: 31px; top: 64px; bottom: -28px; width: 0; border-left: 2px dashed #DFE8F1; }
        .flujo-icono { width: 56px; height: 56px; border-radius: 50%; background: #EAEFF5; color: #253C5C; display: flex; align-items: center; justify-content: center; flex-shrink: 0; z-index: 1; }
        .flujo-icono svg { width: 26px; height: 26px; stroke-width: 1.8; }
        .flujo-orden { font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; font-size: 0.72rem; color: #888; margin: 0 0 2px; }
        .flujo-titulo { font-family: 'Antonio', 'Anton', 'Figtree', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; font-size: 1.15rem; color: #253C5C; margin: 0 0 4px; }
        .flujo-responsable { font-size: 0.75rem; color: #888; text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 12px; font-weight: 600; }
        .flujo-texto { font-size: 0.9rem; color: #3a3a3a; line-height: 1.7; margin: 0 0 8px; }
        .flujo-texto:last-child { margin-bottom: 0; }
        .flujo-texto strong { color: #2f3e3c; font-weight: 700; }
        .flujo-texto ul { margin: 6px 0 0 20px; padding: 0; }
        .flujo-texto ul li { margin-bottom: 4px; }
        @media (max-width: 720px) {
            .flujo-paso { grid-template-columns: 48px 1fr; gap: 14px; }
            .flujo-paso:not(:last-child)::before { left: 23px; top: 48px; }
            .flujo-icono { width: 40px; height: 40px; }
            .flujo-icono svg { width: 20px; height: 20px; }
        }"""

# Línea de tiempo en chevrones: source of truth en _comun/estilos.py para no duplicar.
CSS = CSS + "\n        " + CSS_LINEA_TIEMPO_CHEVRON.replace("\n", "\n        ")
# Componente "Reporte a políticas" (compartido con JCO y Forjar).
CSS = CSS + "\n        " + CSS_REPORTE_POLITICAS.replace("\n", "\n        ")

# ============================================================
# Reporte a políticas — Casas de Juventud a nivel servicio
#
# La lógica de lectura y render vive en _comun/reporte_politicas.py.
# Aquí solo se declara la lista de pendientes específicos de Casas.
#
# A nivel SERVICIO (esta página raíz) se listan los productos del Excel
# de la Política Pública Distrital de Juventud con responsable "Casas
# de Juventud" o "Todos". Los productos de otras políticas que sí son
# específicos de un eje (Bienestar, Cultura, Inclusión, Liderazgo)
# viven dentro de la sección "Reporte a políticas" de cada eje y al
# final de aquellas hay una nota apuntando a esta sección de servicio.
# ============================================================

# Las tres políticas étnicas transversales (Indígena 1.3.12, Rrom 6.1.8 y
# Palenquera 6.3.15 — parte de Negra/Afro/Palenquera) se leen desde la hoja
# "Mapeo general" del Excel de Reportes Externos. En esa hoja tienen
# responsable "Todos" y aplican a los tres servicios (Casas, JCO y Forjar).
# La observación literal sobre el flujo de PUA y DADE vive en el Excel.
#
# Raizal 4.1.12 NO aplica a Casas (es JCO específico, transferencias
# monetarias). Negra/Afro 1.3.9 aplica solo a Casas (habla de "Casa de la
# juventud que implemente el protocolo...") y viene de la hoja Casas.
#
# PAT de Víctimas se eliminó de los pendientes: es el mismo instrumento que
# el PAD Víctimas, que ya aparece confirmado en esta sección.
_POLITICAS_EN_REVISION_CASAS = []

# Productos del Excel de Reportes Externos hoja "Casas de Juventud"
# que NO son específicos de un eje sino del servicio Casas como tal,
# y por lo tanto se muestran en esta sección (no en Bienestar/Cultura/
# Inclusión/Liderazgo). Estos quedan excluidos del eje Bienestar vía
# Configuracion_politicas_bienestar.xlsx hoja Exclusiones.
_TEMAS_NIVEL_SERVICIO_CASAS = [
    "Negra, Afrocolombiana",  # producto 1.3.9 Casa de la juventud que implementa enfoque diferencial étnico
    "LGBTI",                  # 3 productos, todos en revisión con Daniela
    "PAD Víctimas",           # producto transversal a los tres servicios
    "Plan operativo de reincorporación",  # transversal del servicio Casas, no de un eje específico
    "PDET",                   # transversal del servicio Casas (territorios PDET + Sumapaz)
]

filas_politicas_casas = leer_politicas(
    xlsx_ppdj=os.path.join(BASE, "ejes", "Políticas", "Reporte Política Pública Distrital de Juventud.xlsx"),
    xlsx_externos=os.path.join(BASE, "ejes", "Políticas", "Reportes Externos Subdirección Juventud 2026.xlsx"),
    hoja_externos="Casas de Juventud",
    patron_responsable_ppdj=r"Casas de Juventud|Todos",
    pendientes_revision=_POLITICAS_EN_REVISION_CASAS,
    temas_externos_incluir=_TEMAS_NIVEL_SERVICIO_CASAS,
    temas_mapeo_general_incluir=["1.3.12", "6.1.8", "6.3.15"],
)
cards_politicas_casas = html_cards(filas_politicas_casas)

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestor de Conocimiento - Casas de Juventud 2025</title>
    <style>
{CSS}
    </style>
</head>
<body>
    <header class="header">
        <div>
            <h1>Gestor de Conocimiento - Casas de Juventud</h1>
            <div class="subtitle">Subdirecci&oacute;n para la Juventud | SDIS</div>
        </div>
        <div style="display:flex; align-items:center; gap:15px;">
            <a class="home-btn" href="index.html" title="Todos los servicios" style="text-decoration:none; color:#F8F4E1;">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#F8F4E1" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </a>
            <div class="home-btn" onclick="showContent('welcome')" title="Inicio Casas de Juventud">
                <img src="imagenes/servicios/casas-de-juventud.png" alt="Casas de Juventud" style="height:32px; border-radius:16px; object-fit:contain; vertical-align:middle;">
            </div>
        </div>
    </header>
    <div class="container">
        <nav class="sidebar">
            <div class="sidebar-title" onclick="showContent('welcome')" style="cursor:pointer;"><span>Inicio</span></div>
            <div class="sidebar-section"><div class="sidebar-title active" onclick="toggleSection(this)"><span>Contexto</span><span class="arrow">&#9654;</span></div>
                <div class="sidebar-items show"><div class="sidebar-item" onclick="showContent('linea_tiempo')">Línea de tiempo</div><div class="sidebar-item" onclick="showContent('cambios2026')">A tener en cuenta</div><div class="sidebar-item" onclick="showContent('equipo')">Equipo</div><div class="sidebar-item" onclick="showContent('ubicacion')">Ubicación</div></div></div>

            <div class="sidebar-section"><div class="sidebar-title" onclick="toggleSection(this)"><span>Gestión de datos</span><span class="arrow">&#9654;</span></div>
                <div class="sidebar-items">
                    <div class="sidebar-item" onclick="showContent('flujo_datos')">Flujo de gestión de la información</div>
                    <div class="sidebar-item" onclick="showContent('homologacion')">Homologación SIRBE</div>
                </div></div>



            <div class="sidebar-section"><div class="sidebar-title" onclick="showContent('reporte_politicas')" style="cursor:pointer;"><span>Reporte a pol&iacute;ticas</span></div></div>

            <div class="sidebar-section"><div class="sidebar-title" onclick="showContent('aliados')" style="cursor:pointer;"><span>Aliados</span></div></div>

            <div class="sidebar-section"><div class="sidebar-title" onclick="showContent('resumen')" style="cursor:pointer;"><span>Estadísticas</span></div></div>
        </nav>
        <main class="main-content">
            <div class="content-section active" id="welcome"><div class="welcome-section">
                <div style="font-family:'Anton','Figtree',sans-serif; font-weight:400; font-size:1.9rem; line-height:1.05; letter-spacing:1px; text-transform:uppercase; background:#2d2a28; color:#f4f5de; padding:14px 24px 11px; margin:0 auto 28px; display:block; width:fit-content; max-width:100%; text-align:center;">Casas de Juventud</div>
                <p style="text-align:left; max-width:800px; margin:0 auto 30px; font-size:0.95rem; line-height:1.7; color:#444;">Las Casas de Juventud son espacios distritales dirigidos a jóvenes entre 14 y 28 años, concebidos como el eje territorial de la Política Pública Distrital de Juventud (PPDJ). Articulan la oferta de servicios de la administración distrital para garantizar acceso equitativo a oportunidades de desarrollo personal, social y laboral, y funcionan como la principal puerta de entrada al ecosistema de servicios del Distrito.</p>
                
                <p style="text-align:left; max-width:800px; margin:20px auto 25px; font-size:0.95rem; line-height:1.7; color:#444;">Para materializar este propósito y fortalecer los proyectos de vida de los jóvenes, el servicio de Casas de Juventud estructura su oferta integral a través de los siguientes ejes:</p>
                <div class="quick-links">
{quick_links}                </div>

                <div style="margin-top:20px; display:flex; justify-content:center;">
                    <a class="quick-link" href="ejes/SIDICU.html" style="max-width:400px; width:100%; padding:15px 20px; text-decoration:none; color:inherit;">
                        <div class="sidicu-img" style="width:100%; height:180px; border-radius:12px; overflow:hidden; border:3px solid #DFE6F0; box-shadow:0 4px 15px rgba(102,58,147,0.15); transition:border-color 0.3s, transform 0.3s; margin-bottom:12px;">
                            <img src="imagenes/sidicu.png" alt="SIDICU" style="width:100%; height:100%; object-fit:cover;">
                        </div>
                        <div class="quick-link-title" style="font-size:1.1rem;">SIDICU — Sistema Distrital del Cuidado</div>
                    </a>
                </div>
            </div></div>

            <div class="content-section" id="linea_tiempo"><div class="card">
                <h2 class="card-title">Línea de tiempo</h2>
                <p style="color:#555; margin-bottom:24px; line-height:1.7;">Las Casas de Juventud surgieron en el marco de la Política Pública de Juventud 2006&ndash;2016 y se han consolidado como la principal puerta de entrada al ecosistema de servicios distritales para jóvenes entre 14 y 28 años.</p>
                <div class="linea-tiempo" style="--lt-cols: 7;">
                    <article class="lt-hito">
                        <div class="lt-chevron lt-c1"><i data-lucide="lightbulb"></i>1991</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i1"><i data-lucide="lightbulb"></i></div>
                            <div class="lt-titulo lt-t1">Primeros intentos</div>
                            <div class="lt-texto">Surgen los primeros esfuerzos por crear Casas de Juventud en Bogotá como espacios de encuentro y participación juvenil.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c2"><i data-lucide="file-text"></i>2006</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i2"><i data-lucide="file-text"></i></div>
                            <div class="lt-titulo lt-t2">PPDJ 2006&ndash;2016</div>
                            <div class="lt-texto">Se expide el <strong>Decreto 482 de 2006</strong>, que da origen formal al marco en el que posteriormente se implementarán las Casas de Juventud como espacios de participación y desarrollo juvenil.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c3"><i data-lucide="door-open"></i>2011</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i3"><i data-lucide="door-open"></i></div>
                            <div class="lt-titulo lt-t3">Apertura de primeras sedes</div>
                            <div class="lt-texto">Proyecto 764 &ldquo;Jóvenes activando su ciudadanía&rdquo;: las Casas se conciben como espacios para la participación activa y el fortalecimiento organizativo. Se abren las primeras en <strong>Fontibón, Mártires, Antonio Nariño y Barrios Unidos</strong>.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c4"><i data-lucide="building"></i>2016</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i4"><i data-lucide="building"></i></div>
                            <div class="lt-titulo lt-t4">13 casas en operación</div>
                            <div class="lt-texto">Para este año el Distrito ya cuenta con <strong>13 Casas de Juventud</strong> (SDIS, 2016), consolidando el servicio en varias localidades de la ciudad.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c5"><i data-lucide="users"></i>2016&ndash;19</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i5"><i data-lucide="users"></i></div>
                            <div class="lt-titulo lt-t5">Distrito Joven</div>
                            <div class="lt-texto">Proyecto 1116 en la administración <em>Bogotá Mejor para Todos</em>: fortalecimiento de capacidades juveniles, ampliación de oportunidades y apropiación de la política pública. Se consolidan los espacios y se abren nuevas sedes.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c6"><i data-lucide="landmark"></i>2020</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i6"><i data-lucide="landmark"></i></div>
                            <div class="lt-titulo lt-t6">PPDJ 2019&ndash;2030</div>
                            <div class="lt-texto">Se implementa la nueva Política Pública Distrital de Juventud. El servicio cuenta con <strong>18 Casas de Juventud</strong> distribuidas en el Distrito, además de una <strong>Unidad Móvil</strong> que amplía la cobertura territorial.</div>
                        </div>
                    </article>

                    <article class="lt-hito">
                        <div class="lt-chevron lt-c7"><i data-lucide="target"></i>2027</div>
                        <div class="lt-cuerpo">
                            <div class="lt-icono lt-i7"><i data-lucide="target"></i></div>
                            <div class="lt-titulo lt-t7">Una casa por localidad</div>
                            <div class="lt-texto">Producto <strong>1.1.7</strong> del CONPES D.C. 08: <em>&ldquo;Número de localidades de la ciudad con al menos una casa de la juventud en operación&rdquo;</em>. Meta: <strong>mínimo una Casa de Juventud por localidad para 2027</strong>.</div>
                        </div>
                    </article>
                </div>
            </div></div>

            <div class="content-section" id="cambios2026"><div class="card">
                <h2 class="card-title">A tener en cuenta</h2>
                <p style="line-height:1.7; margin-bottom:8px;">El servicio Casas de Juventud cambi&oacute; sustancialmente entre la administraci&oacute;n anterior y el modelo vigente desde 2026. La tabla resume las cinco transformaciones principales; debajo se desarrolla cada una en detalle.</p>

                <div class="comp-tabla">
                    <img class="comp-mano-decor" src="imagenes/manos/3.png" alt="">
                    <div class="comp-tabla-grid">
                        <div class="comp-header comp-header-1">Caracter&iacute;stica</div>
                        <div class="comp-header comp-header-2">Modelo anterior</div>
                        <div class="comp-header comp-header-3">Nuevo modelo 2026</div>

                        <div class="comp-row-sep"></div>
                        <div class="comp-caracteristica">Enfoque principal</div>
                        <div class="comp-celda comp-anterior"><div class="comp-icono"><i data-lucide="clock"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Anterior</span><span>Aprovechamiento del tiempo libre</span></div></div>
                        <div class="comp-celda comp-nuevo"><div class="comp-icono"><i data-lucide="settings"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Nuevo 2026</span><span>Fortalecimiento de habilidades y oportunidades</span></div></div>

                        <div class="comp-row-sep"></div>
                        <div class="comp-caracteristica">Base de planeaci&oacute;n</div>
                        <div class="comp-celda comp-anterior"><div class="comp-icono"><i data-lucide="message-circle"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Anterior</span><span>Gesti&oacute;n aut&oacute;noma del gestor</span></div></div>
                        <div class="comp-celda comp-nuevo"><div class="comp-icono"><i data-lucide="map"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Nuevo 2026</span><span>Diagn&oacute;stico territorial y plan de acci&oacute;n</span></div></div>

                        <div class="comp-row-sep"></div>
                        <div class="comp-caracteristica">Estructura</div>
                        <div class="comp-celda comp-anterior"><div class="comp-icono"><i data-lucide="git-branch"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Anterior</span><span>Oferta dispersa</span></div></div>
                        <div class="comp-celda comp-nuevo"><div class="comp-icono"><i data-lucide="layers"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Nuevo 2026</span><span>4 ejes tem&aacute;ticos</span></div></div>

                        <div class="comp-row-sep"></div>
                        <div class="comp-caracteristica">Inscripci&oacute;n</div>
                        <div class="comp-celda comp-anterior"><div class="comp-icono"><i data-lucide="clipboard-check"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Anterior</span><span>Convocatorias puntuales</span></div></div>
                        <div class="comp-celda comp-nuevo"><div class="comp-icono"><i data-lucide="route"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Nuevo 2026</span><span>Sistema estructurado de enrutamiento</span></div></div>

                        <div class="comp-row-sep"></div>
                        <div class="comp-caracteristica">Modelo de financiamiento</div>
                        <div class="comp-celda comp-anterior"><div class="comp-icono"><i data-lucide="coins"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Anterior</span><span>Proyectos de inversi&oacute;n diferenciados</span></div></div>
                        <div class="comp-celda comp-nuevo"><div class="comp-icono"><i data-lucide="dollar-sign"></i></div><div class="comp-celda-texto"><span class="comp-celda-badge">Nuevo 2026</span><span>Proyecto de inversi&oacute;n unificado</span></div></div>
                    </div>
                </div>

                <h3 class="card-subtitle">Transformaci&oacute;n del servicio (enfoque y estructura)</h3>
                <p style="line-height:1.7;">Hasta el 2024, el funcionamiento de las Casas de Juventud se pensaba casi exclusivamente para el &ldquo;aprovechamiento del tiempo libre&rdquo; y depend&iacute;a de la gesti&oacute;n aut&oacute;noma que lograra cada gestor. A partir de 2025, el servicio se transform&oacute; y se estructur&oacute; en cuatro ejes tem&aacute;ticos principales: Bienestar, Cultura, Inclusi&oacute;n Productiva y Participaci&oacute;n.</p>

                <h3 class="card-subtitle">Cambio metodol&oacute;gico 2026 (base de planeaci&oacute;n)</h3>
                <p style="line-height:1.7;">Para 2026, el funcionamiento tuvo un cambio metodol&oacute;gico. La planeaci&oacute;n de la oferta pas&oacute; a ser bimensual y ahora exige obligatoriamente un diagn&oacute;stico territorial. Esto significa que los gestores deben realizar an&aacute;lisis de fuentes secundarias, mapeo de actores, recorridos territoriales y cartograf&iacute;a social para crear un plan de acci&oacute;n que cruce las necesidades reales de los j&oacute;venes con las metas de la pol&iacute;tica p&uacute;blica. Adicionalmente, se estandariz&oacute; el tipo de oferta en todas las casas bajo formatos espec&iacute;ficos: experiencias, talleres, ciclos formativos, semilleros y laboratorios.</p>

                <h3 class="card-subtitle">Enrutamiento (inscripci&oacute;n)</h3>
                <p style="line-height:1.7;">Se implement&oacute; un nuevo paso llamado &ldquo;enrutamiento&rdquo;. Cuando un joven llega a la Casa de Juventud, se le aplica un cuestionario inicial para identificar sus intereses y necesidades puntuales. A partir de all&iacute;, se le orienta hacia la oferta institucional adecuada.</p>

                <h3 class="card-subtitle">Consolidaci&oacute;n de proyectos de inversi&oacute;n (financiamiento)</h3>
                <p style="line-height:1.7;">Durante la administraci&oacute;n 2020&ndash;2024, la atenci&oacute;n a j&oacute;venes operaba a trav&eacute;s de dos proyectos de inversi&oacute;n independientes: el <a href="https://www.integracionsocial.gov.co/images/_docs/2024/Gestion/EBI/19072024-7740.pdf" target="_blank" style="color:#3A5275;">7740</a> (Generaci&oacute;n J&oacute;venes con Derechos), que abarcaba la oferta general de las Casas de Juventud, y el <a href="https://www.integracionsocial.gov.co/images/_docs/2024/Gestion/EBI/22042024-7753-Prevencion-maternidad-paternidad-temprana-Bogota.pdf" target="_blank" style="color:#3A5275;">7753</a>, un proyecto espec&iacute;fico de maternidades y paternidades tempranas que atend&iacute;a poblaci&oacute;n desde los 10 a&ntilde;os. Con el cambio de administraci&oacute;n, ambos proyectos se unificaron en un solo proyecto de inversi&oacute;n (<a href="https://www.integracionsocial.gov.co/images/_docs/2024/Gestion/EBI/14022025-7940.pdf" target="_blank" style="color:#3A5275;">el 7940</a>). El componente de maternidades y paternidades dej&oacute; de ser un proyecto independiente y pas&oacute; a funcionar como una l&iacute;nea de trabajo dentro del eje de bienestar, con atenci&oacute;n exclusiva a j&oacute;venes entre 14 y 28 a&ntilde;os.</p>

                <h3 class="card-subtitle">M&iacute;nimos operativos y regulaci&oacute;n de alianzas</h3>
                <p style="line-height:1.7;">El nivel central define &ldquo;m&iacute;nimos operativos&rdquo; mensuales por cada eje que todas las localidades deben cumplir. Estos m&iacute;nimos se ajustan a momentos espec&iacute;ficos del a&ntilde;o (por ejemplo, actividades conmemorativas del 8M o inducci&oacute;n a nuevos consejeros). Tambi&eacute;n se estructur&oacute; la figura del voluntariado: quienes deseen impartir talleres deben llenar un formulario, presentar documentos de seguridad y someter su metodolog&iacute;a a revisi&oacute;n t&eacute;cnica.</p>
            </div></div>

            <div class="content-section" id="equipo">{equipo_html}</div>

            <div class="content-section" id="homologacion"><div class="card">
                <h2 class="card-title">Tabla de homologación por eje
                    <span class="sirbe-help" onclick="document.getElementById('sirbe-tooltip').style.display=document.getElementById('sirbe-tooltip').style.display==='block'?'none':'block'" style="display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:50%; background:#253C5C; color:#fff; font-size:0.7rem; font-weight:700; cursor:pointer; margin-left:8px; vertical-align:middle; position:relative;">?</span>
                </h2>
                <div id="sirbe-tooltip" style="display:none; background:#2F3E3C; color:#F8F4E1; border-radius:10px; padding:18px 22px; margin-bottom:18px; font-size:0.88rem; line-height:1.7; position:relative;">
                    <strong>SIRBE</strong> (Sistema de Información para el Registro de Beneficiarios) es el aplicativo de datos de la Secretaría Distrital de Integración Social (SDIS) en Bogotá. Su función principal es registrar, sistematizar y hacer seguimiento a la información de los ciudadanos que acceden a los servicios sociales, proyectos y ayudas de la entidad. La información registrada en las "Fichas SIRBE" es confidencial y su uso se limita a la gestión interna de servicios sociales.
                    <span onclick="document.getElementById('sirbe-tooltip').style.display='none'" style="position:absolute; top:8px; right:12px; cursor:pointer; opacity:0.6; font-size:1.1rem;">×</span>
                </div>
                <p style="color:#666; font-size:0.9rem; margin-bottom:15px;">Así se organiza la información de cada eje en SIRBE y en la ficha física. Fuente: tabla de homologación institucional 2026.</p>

<h3 class="card-subtitle">Estructura oficial en SIRBE</h3>
                <div style="background:#f8f9fa; border-radius:10px; padding:20px 25px; margin-bottom:25px; border:1px solid #e0e0e0;">
                    <p style="font-weight:600; color:#3A3A3A; margin-bottom:8px;">Jerarquía de campos en SIRBE</p>
                    <p style="font-size:0.85rem; color:#666; margin-bottom:12px;">Cada registro en SIRBE tiene estos 5 niveles de desagregación:</p>
                    <div style="font-size:0.88rem; line-height:1.8;">
                        <div><span style="font-family:monospace; background:#DFE6F0; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.8rem;">SERVICIO</span> — Casas de Juventud</div>
                        <div style="margin-left:25px; border-left:2px solid #ddd; padding-left:15px; margin-top:4px;">
                            <div><span style="font-family:monospace; background:#DFE6F0; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.8rem;">NOMMODALIDAD</span> — Atención Inclusiva BMT 2939 / Estrategia Móvil BMT 2938</div>
                            <div style="margin-left:25px; border-left:2px solid #ddd; padding-left:15px; margin-top:4px;">
                                <div><span style="font-family:monospace; background:#DFE6F0; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.8rem;">ACTUACION_INTERV_CURSO</span> — Varía por eje (Prevención Integral, Manejo Adecuado de Tiempo Libre, etc.)</div>
                                <div style="margin-left:25px; border-left:2px solid #ddd; padding-left:15px; margin-top:4px;">
                                    <div><span style="font-family:monospace; background:#DFE6F0; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.8rem;">NOMACTIVIDAD_CURSO</span> — Actividades oficiales del eje (ver tablas abajo)</div>
                                    <div style="margin-left:25px; border-left:2px solid #ddd; padding-left:15px; margin-top:4px;">
                                        <div><span style="font-family:monospace; background:#DFE6F0; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.8rem;">NOMBRE_CURSO</span> — Campo de texto abierto donde se digita el nombre del curso. Debe coincidir con la tabla de homologación, pero al ser manual genera variaciones</div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:20px;">
                        <thead>
                            <tr>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:18%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Oferta Casas de Juventud 2026</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:22%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Actuación en SIRBE misional</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Nombre de la actividad</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Tipo de actividad (ficha física y SIRBE)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td rowspan="5" style="padding:10px; border:1px solid #ddd; font-weight:600; vertical-align:middle; text-align:center; background:#fafafa;">BIENESTAR</td>
                                <td rowspan="5" style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">PREVENCIÓN INTEGRAL</td>
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Derechos sexuales y derechos reproductivos</td>
                                <td rowspan="2" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle;">7. Talleres informativos en prevención (1486 en SIRBE)</td>
                            </tr>
                            <tr>
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Prevención de violencias basadas en género</td>
                            </tr>
                            <tr style="background:#fafafa;">
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Salud mental</td>
                                <td style="padding:8px 10px; border:1px solid #ddd;">8. Acompañamiento y orientación psicosocial (511 en SIRBE)</td>
                            </tr>
                            <tr>
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Prevención de consumo de sustancias psicoactivas SPA</td>
                                <td style="padding:8px 10px; border:1px solid #ddd;">9. Cuidado frente al consumo responsable de SPA (1487 en SIRBE)</td>
                            </tr>
                            <tr style="background:#fafafa;">
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Salas de escucha</td>
                                <td style="padding:8px 10px; border:1px solid #ddd;">10. Centros de escucha (1485 en SIRBE)</td>
                            </tr>
                        </tbody>
                    </table>

                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:20px;">
                        <thead>
                            <tr>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:18%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Oferta Casas de Juventud 2026</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:22%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Actuación en SIRBE misional</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Nombre de la actividad</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Tipo de actividad (ficha física y SIRBE)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td rowspan="11" style="padding:10px; border:1px solid #ddd; font-weight:600; vertical-align:middle; text-align:center; background:#fafafa;">CULTURA</td>
                                <td rowspan="11" style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">MANEJO ADECUADO DE TIEMPO LIBRE</td>
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Medio ambiente</td>
                                <td rowspan="9" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle;">13. Formación artística focalizada (1489 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Deportes</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Artes musicales</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Artes audiovisuales</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Artes plásticas</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Artes escénicas</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Artes circenses</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Artes urbanas</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Gestión cultural</td></tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Uso de espacios para desarrollo de prácticas artísticas</td>
                                <td style="padding:8px 10px; border:1px solid #ddd;">11. Aprovechamiento del tiempo libre con énfasis en intereses juveniles (1207 en SIRBE)</td>
                            </tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Semana de la juventud</td>
                                <td style="padding:8px 10px; border:1px solid #ddd; background:#fafafa;">15. Semana de la juventud (1491 en SIRBE)</td>
                            </tr>
                        </tbody>
                    </table>

                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:20px;">
                        <thead>
                            <tr>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:18%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Oferta Casas de Juventud 2026</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:22%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Actuación en SIRBE misional</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Nombre de la actividad</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Tipo de actividad (ficha física y SIRBE)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td rowspan="8" style="padding:10px; border:1px solid #ddd; font-weight:600; vertical-align:middle; text-align:center; background:#fafafa;">INCLUSIÓN SOCIAL Y PRODUCTIVA</td>
                                <td rowspan="8" style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">FORMACIÓN PARA EL PROYECTO DE VIDA</td>
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Habilidades informáticas / Oferta TIC</td>
                                <td rowspan="3" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle;">16. Salas TIC para el fortalecimiento de habilidades y capacidades juveniles (1496 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Idiomas</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Lecto escritura / Matemáticas</td></tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Empleo</td>
                                <td rowspan="3" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle; background:#fafafa;">17. Formación en emprendimiento y empleabilidad (1497 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Emprendimiento</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Fest oportunidades</td></tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td style="padding:6px 10px; border:1px solid #ddd; text-align:center;">Educación financiera</td>
                                <td rowspan="2" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle;">18. Orientación socio-ocupacional dirigida a jóvenes (1498 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; text-align:center; background:#fafafa;">Formación (rutas de modelo de educación flexible)</td></tr>
                        </tbody>
                    </table>

                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:20px;">
                        <thead>
                            <tr>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:18%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Oferta Casas de Juventud 2026</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:22%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Actuación en SIRBE misional</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Nombre de la actividad</th>
                                <th style="background:#2F3E3C; color:#F8F4E1; padding:12px 14px; width:30%; text-align:left; font-weight:700; border:1px solid #2F3E3C;">Tipo de actividad (ficha física y SIRBE)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td rowspan="7" style="padding:10px; border:1px solid #ddd; font-weight:600; vertical-align:middle; text-align:center; background:#fafafa;">LIDERAZGO Y PARTICIPACIÓN</td>
                                <td rowspan="3" style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">Asesoría jurídica y participación</td>
                                <td rowspan="3" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle; text-align:center;">Asesoría jurídica y participación</td>
                                <td style="padding:6px 10px; border:1px solid #ddd;">22. Promoción y fortalecimiento de actividades de organización juvenil (1494 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; background:#fafafa;">23. Escenarios de diálogos intergeneracionales y de saberes que fortalezcan la ciudadanía juvenil (1495 en SIRBE)</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd;">20. Atención y orientación jurídica a jóvenes (1492 en SIRBE)</td></tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td rowspan="3" style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">Política pública de juventud</td>
                                <td rowspan="3" style="padding:8px 10px; border:1px solid #ddd; vertical-align:middle; text-align:center;">Política pública de juventud</td>
                                <td style="padding:6px 10px; border:1px solid #ddd; background:#fafafa;">3. Socialización de política pública de juventud (1502 en SIRBE)</td>
                            </tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd;">4. Derechos humanos y promoción a la participación (1503 en SIRBE)</td></tr>
                            <tr><td style="padding:6px 10px; border:1px solid #ddd; background:#fafafa;">5. Comités operativos locales de juventud (1504 en SIRBE)</td></tr>
                            <tr style="border-top:2px solid #ccc;">
                                <td style="padding:10px; border:1px solid #ddd; vertical-align:middle; text-align:center; background:#fafafa;">Voluntariado intergeneracional</td>
                                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center;">Voluntariado intergeneracional</td>
                                <td style="padding:6px 10px; border:1px solid #ddd;">6. Voluntariado intergeneracional (1505 en SIRBE)</td>
                            </tr>
                        </tbody>
                    </table>

            </div></div>

                        <div class="content-section" id="ubicacion"><div class="card">
                <h2 class="card-title">Ubicación</h2>
                <p style="color:#666; margin-bottom:20px;">Casas de Juventud en Bogotá.</p>
                <div style="margin-bottom:20px;">
                    <a href="https://servicios.sdis.gov.co/index.php/casas-juventud" target="_blank" style="font-family:'Anton','Figtree',sans-serif; font-weight:400; font-size:1.05rem; line-height:1.15; letter-spacing:0.5px; background:#c7eb91; color:#5d5e95; padding:9px 14px 7px; display:inline-block; text-align:center; max-width:100%; text-decoration:none;">Link de agendamiento de actividades en Casas de Juventud</a>
                </div>
                <div style="text-align:center; margin-bottom:25px;">
                    <img src="imagenes/Casas%20Mapa.jpg" alt="Mapa de Casas de Juventud" style="max-width:100%; border-radius:12px;">
                </div>
                <div style="text-align:center; margin-bottom:25px;">
                    <iframe src="mapa_casas_juventud.html" style="width:100%; height:500px; border:none; border-radius:12px;"></iframe>
                </div>
                {directorio_html}
            </div></div>

            <div class="content-section" id="flujo_datos"><div class="card">
                <h2 class="card-title">Flujo de gestión de la información</h2>
                <p style="color:#666; margin-bottom:20px;">Proceso completo desde la ejecución de una actividad hasta el reporte oficial de metas.</p>

                <div class="flujo-pasos">
                    <div class="flujo-paso">
                        <div class="flujo-icono"><i data-lucide="users"></i></div>
                        <div>
                            <p class="flujo-orden">Paso 01</p>
                            <h3 class="flujo-titulo">Ejecución de la actividad</h3>
                            <p class="flujo-responsable">Gestor &middot; aliado externo</p>
                            <p class="flujo-texto">La actividad (taller, charla, evento) es ejecutada por el gestor territorial, un aliado externo, voluntario o equipo transversal. Los gestores culturales, por ejemplo, rotan por las localidades dictando talleres específicos de su línea artística.</p>
                        </div>
                    </div>

                    <div class="flujo-paso">
                        <div class="flujo-icono"><i data-lucide="clipboard-list"></i></div>
                        <div>
                            <p class="flujo-orden">Paso 02</p>
                            <h3 class="flujo-titulo">Recolección en campo</h3>
                            <p class="flujo-responsable">Gestor &middot; aliado externo</p>
                            <p class="flujo-texto">Se diligencia la <strong>Ficha SIRBE 328</strong> (FOR-PSS-328), un formato físico único con dos secciones:</p>
                            <p class="flujo-texto"><ul><li><strong>Sección A (&ldquo;cabezote&rdquo;):</strong> la diligencia el gestor. Registra nombre de la actividad, fecha, horas, lugar, modalidad y tipo de actividad.</li><li><strong>Sección B (listado de asistencia):</strong> la llenan los jóvenes participantes. Registran datos sociodemográficos y responden preguntas transversales de política pública (por ejemplo, derechos vulnerados o ruta de oportunidades juveniles). Cada joven firma como constancia de participación.</li></ul></p>
                        </div>
                    </div>

                    <div class="flujo-paso">
                        <div class="flujo-icono"><i data-lucide="inbox"></i></div>
                        <div>
                            <p class="flujo-orden">Paso 03</p>
                            <h3 class="flujo-titulo">Entrega física</h3>
                            <p class="flujo-responsable">Administrativo</p>
                            <p class="flujo-texto">Los gestores entregan las fichas físicas y listados al administrativo de la Casa de Juventud. <strong>Plazo: días 1 al 20 de cada mes.</strong></p>
                        </div>
                    </div>

                    <div class="flujo-paso">
                        <div class="flujo-icono"><i data-lucide="monitor"></i></div>
                        <div>
                            <p class="flujo-orden">Paso 04</p>
                            <h3 class="flujo-titulo">Digitación en SIRBE</h3>
                            <p class="flujo-responsable">Administrativo &middot; digitador</p>
                            <p class="flujo-texto">El administrativo traduce la ficha física al sistema SIRBE. <strong>Plazo: días 20 al 28 del mismo mes.</strong></p>
                            <p class="flujo-texto"><ul><li>Selecciona modalidad y actuación de listas desplegables.</li><li>Digita el <strong>nombre del curso</strong> en un campo de texto abierto, usando la <strong>tabla de homologación</strong> como referencia.</li><li>Carga los jóvenes asistentes. El sistema asigna automáticamente el estado &ldquo;Atendido Curso&rdquo;.</li></ul></p>
                        </div>
                    </div>

                    <div class="flujo-paso">
                        <div class="flujo-icono"><i data-lucide="bar-chart-3"></i></div>
                        <div>
                            <p class="flujo-orden">Paso 05</p>
                            <h3 class="flujo-titulo">Reporte oficial DADE y desagregación manual</h3>
                            <p class="flujo-responsable">DADE &middot; Líder administrativa</p>
                            <p class="flujo-texto">La primera semana del mes siguiente, la <strong>DADE</strong> (Dirección de Análisis y Diseño Estratégico) emite el reporte oficial consolidado a partir de los datos cargados en el sistema. Debido a que DADE entrega la información de manera general, la líder administrativa debe descargar este reporte y realizar un trabajo manual para desagregar los datos por unidad operativa y por equipo móvil, y así poder evaluar el cumplimiento de las metas internas.</p>
                        </div>
                    </div>
                </div>

                
            
                <div style="text-align:center; margin:25px 0 10px;">
                    <img src="imagenes/SIRBE%20informacion.png" alt="Flujo de información SIRBE" style="max-width:100%; border-radius:12px; box-shadow:0 2px 12px rgba(0,0,0,0.08);">
                </div>

                <h3 class="card-subtitle" style="margin-top:30px;">A tener en cuenta</h3>
                <p style="line-height:1.7;">En el reporte de datos es importante entender cómo se diligencia la información. El gestor que registra la ficha y reporta la actividad no siempre es quien la ejecuta. Esto ocurre por tres dinámicas operativas:</p>
                <ul style="margin:10px 0 15px 25px; line-height:1.7;">
                    <li><strong>Aliados y voluntarios:</strong> parte de la oferta es ejecutada por agentes externos. En estos casos, el gestor asume el acompañamiento logístico y la recolección de datos.</li>
                    <li><strong>Gestores transversales:</strong> existe un equipo especializado de gestores culturales que rota por las localidades dictando talleres específicos de su línea artística.</li>
                    <li><strong>Eventos masivos:</strong> en actividades de gran escala pueden participar varios gestores apoyando la logística, y todos terminan registrando fichas a su nombre para el mismo espacio.</li>
                </ul>

                <h3 class="card-subtitle" style="margin-top:30px;">Diagrama de flujo del proceso</h3>
                <p style="color:#666; font-size:0.85rem; margin-bottom:12px;">Representación visual del ciclo de recolección y digitación en SIRBE para las Casas de Juventud.</p>
                <div style="overflow-x:auto; margin-bottom:8px;">%%SVG_DIAGRAMA_CASAS%%</div>
            </div></div>

            <div class="content-section" id="reporte_politicas"><div class="card">
                <h2 class="card-title">Reporte a pol&iacute;ticas</h2>

                <div class="jp-callout">
                    Esta secci&oacute;n lista los productos a los que reporta el servicio <strong>Casas de Juventud</strong> de manera transversal &mdash; principalmente los de la <strong>Pol&iacute;tica P&uacute;blica Distrital de Juventud</strong>. Para productos espec&iacute;ficos por eje, ver el <em>Reporte a pol&iacute;ticas</em> dentro de cada eje (Bienestar, Cultura, Inclusi&oacute;n, Liderazgo).
                </div>

                <div class="jp-grid">
{cards_politicas_casas}
                </div>
            </div></div>

{seccion_aliados_casas()}

            <div class="content-section" id="resumen"><div class="card">
                <h2 class="card-title">Resumen general</h2>
                <p style="margin-bottom:15px;"><strong>Tablero oficial:</strong> Seguimiento técnico (Power BI)</p>
                <iframe title="Seguimiento técnico - Casas de Juventud" width="100%" height="600" src="{POWERBI_SRC}" frameborder="0" allowFullScreen="true" style="border:1px solid #e0e0e0; border-radius:8px;"></iframe>
            </div></div>

                </main>
    </div>
    <script>
        function toggleSection(el){{el.classList.toggle('active');el.nextElementSibling.classList.toggle('show');}}
        function showContent(id){{document.querySelectorAll('.content-section').forEach(s=>s.classList.remove('active'));document.querySelectorAll('.sidebar-item').forEach(i=>i.classList.remove('active'));var s=document.getElementById(id);if(s)s.classList.add('active');if(event&&event.target)event.target.closest('.sidebar-item')?.classList.add('active');}}

        var DL={datos_loc_json};
        var DD={datos_demo_json};

        function fmtN(n){{return n.toLocaleString('es-CO');}}

        function setActive(cid,btn){{document.querySelectorAll('#'+cid+' .filter-btn').forEach(function(b){{b.classList.remove('active')}});if(btn)btn.classList.add('active');}}

        function filtrarLoc(eje){{
            var d=DL[eje];if(!d)return;
            var total=0;d.forEach(function(r){{total+=r.at}});
            var mx=d[0].at;
            var h='';
            d.forEach(function(r){{
                var w=(r.at/mx*100).toFixed(0);
                h+='<div class="hbar-row"><div class="hbar-label">'+r.loc+'</div>'
                  +'<div class="hbar-track"><div class="hbar-fill" style="width:'+w+'%;background:#253C5C;"></div></div>'
                  +'<div class="hbar-value">'+fmtN(r.at)+' atenciones</div></div>';
            }});
            h+='<div class="hbar-row" style="font-weight:700;background:#EAEFF5;border-radius:8px;margin-top:5px;">'
              +'<div class="hbar-label">Total Bogot\u00e1</div><div class="hbar-track"></div>'
              +'<div class="hbar-value">'+fmtN(total)+' atenciones</div></div>';
            document.getElementById('loc-bars').innerHTML=h;
            if(event&&event.target)setActive('loc-filters',event.target);
        }}

        function filtrarDemo(eje){{
            var d=DD[eje];if(!d)return;
            var t=d.total;
            ['edad','sexo','etnia','ruv','disc'].forEach(function(cat){{
                var el=document.getElementById('demo-'+cat);
                if(!el||!d[cat])return;
                var h='';
                d[cat].forEach(function(r){{
                    var p=(r.n/t*100);
                    h+='<div class="hbar-row"><div class="hbar-label">'+r.cat+'</div>'
                      +'<div class="hbar-track"><div class="hbar-fill" style="width:'+p.toFixed(1)+'%;background:#253C5C;"></div></div>'
                      +'<div class="hbar-value">'+fmtN(r.n)+' <span class="hbar-pct">('+p.toFixed(1).replace('.',',')+'%)</span></div></div>';
                }});
                el.innerHTML=h;
            }});
            if(event&&event.target)setActive('demo-filters',event.target);
        }}

        document.addEventListener('DOMContentLoaded',function(){{filtrarLoc('Total');filtrarDemo('Total')}});
    </script>
    <!-- Lucide Icons (flat icons de l&iacute;nea, MIT license) para el flujo de gesti&oacute;n -->
    <script src="https://unpkg.com/lucide@latest/dist/umd/lucide.min.js"></script>
    <script>lucide.createIcons();</script>
</body>
</html>"""

# Inyectar el SVG del diagrama de flujo de Casas de Juventud
html = html.replace("%%SVG_DIAGRAMA_CASAS%%", svg_diagrama_casas())

# Guardar
salida = os.path.join(BASE, "gestion_conocimiento_juventud_2025.html")
with open(salida, "w", encoding="utf-8") as f:
    f.write(html)

print(f"HTML generado: {salida}")
print(f"Tamano: {len(html):,} caracteres".replace(",", "."))
print(f"Casas de Juventud 2025: {fmt(total_atenciones)} atenciones, {fmt(personas_unicas)} personas unicas")

