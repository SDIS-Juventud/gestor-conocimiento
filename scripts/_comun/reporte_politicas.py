# -*- coding: utf-8 -*-
"""
Componente compartido para la sección "Reporte a políticas" que aparece
en los gestores de JCO, Forjar y (próximamente) Parche Seguro.

Lee los Excels maestros de Felipe y arma un grid de tarjetas con una card
por política o plan al que el servicio le reporta. Cada servicio aporta:
  - Su propio filtro de responsable en PPDJ (regex).
  - Su hoja específica en "Reportes Externos Subdirección Juventud 2026.xlsx".
  - Su lista de pendientes "En revisión" (políticas mencionadas por Felipe
    que aún no aparecen en los Excels maestros).

Exporta:
  - CSS_REPORTE_POLITICAS: bloque CSS para incluir al final del EXTRAS_CSS
    del servicio. Define las clases .jp-callout y .jp-card-*.
  - leer_politicas(...): devuelve la lista ordenada de dicts con las filas.
  - html_cards(filas): renderiza la lista de filas como HTML del grid.

Regla de orden: Política Pública Distrital de Juventud (PPDJ) primero,
luego el resto de Políticas Públicas, y al final los Planes. Dentro de
cada grupo, las confirmadas van antes que las "En revisión".
"""
import os
import re
import pandas as pd


CSS_REPORTE_POLITICAS = """
/* Reporte a pol&iacute;ticas: tarjetas peque&ntilde;as por pol&iacute;tica o
   plan. Una sola superficie por card (header verde institucional + body
   crema), jerarqu&iacute;a por tipograf&iacute;a y color, sin chips ni
   recuadros anidados. El estado "En revisi&oacute;n" usa grises neutros
   para indicar pendiente sin alarmar. */
.jp-callout { color: #3a3a3a; margin: 0 0 22px; line-height: 1.7; font-size: 0.95rem; max-width: 820px; }
.jp-callout strong { color: #2F3E3C; font-weight: 700; }
.jp-grid { display: grid; grid-template-columns: 1fr; gap: 18px; margin: 18px 0 0; }
.jp-card { background: #fbf8ec; border-radius: 12px; padding: 0; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.04); }
.jp-card-header { background: #2F3E3C; color: #F8F4E1; padding: 16px 22px 14px; }
.jp-card-tema { font-family: 'Antonio','Anton','Figtree',sans-serif; font-weight: 700; font-size: 1.05rem; margin: 0; line-height: 1.3; letter-spacing: 0.01em; }
.jp-card-body { padding: 18px 22px 20px; }
.jp-card-productos { list-style: none; padding: 0; margin: 0; }
.jp-card-producto { font-family: 'Figtree', sans-serif; font-size: 0.92rem; color: #3a3a3a; line-height: 1.55; padding-left: 14px; position: relative; margin-bottom: 10px; }
.jp-card-producto:last-child { margin-bottom: 0; }
.jp-card-producto::before { content: '·'; position: absolute; left: 0; top: -2px; color: #2F3E3C; font-weight: 700; font-size: 1.2rem; line-height: 1; }
.jp-card-codigo { font-family: 'Antonio','Anton','Figtree',sans-serif; font-weight: 700; font-size: 0.95rem; color: #2F3E3C; margin-right: 6px; letter-spacing: 0.02em; }
.jp-card-prod-obs { font-family: 'Figtree', sans-serif; font-size: 0.85rem; color: #555; line-height: 1.6; margin: 6px 0 0; }
.jp-card-pendiente { background: #f0eee9; box-shadow: none; }
.jp-card-pendiente .jp-card-header { background: #6b6963; color: #f0eee9; }
.jp-card-revision-nota { font-size: 0.92rem; line-height: 1.6; color: #555; margin: 0; }
.jp-card-revision-nota strong { color: #2F3E3C; font-weight: 700; }
"""


def _limpiar_celda(valor):
    """Convierte NaN, '-', None y strings vacíos en cadena vacía. Los Excels
    tienen celdas inconsistentes (algunas con '-' explícito, otras vacías,
    otras con 'nan' por la conversión de pandas)."""
    if valor is None:
        return ""
    s = str(valor).strip()
    if not s or s.lower() == "nan" or s == "-":
        return ""
    return s


def _separar_codigo_producto(texto):
    """Separa '3.1.4. Jóvenes vinculados a...' en (codigo, producto). Si no
    encuentra patrón numérico al inicio, devuelve ('', texto)."""
    s = (texto or "").strip()
    if not s:
        return "", ""
    m = re.match(r"^(\d+(?:\.\d+)+)\.?\s+(.+)$", s)
    if m:
        return m.group(1), m.group(2).strip()
    return "", s


def leer_politicas(
    xlsx_ppdj,
    xlsx_externos=None,
    hoja_externos=None,
    patron_responsable_ppdj=None,
    tema_ppdj="Política Pública Distrital de Juventud",
    pendientes_revision=None,
    temas_externos_incluir=None,
):
    """Lee los dos Excels maestros y devuelve una lista ordenada de dicts.

    Parámetros:
        xlsx_ppdj: ruta al archivo de la Política Pública Distrital de Juventud.
        xlsx_externos: ruta al archivo de Reportes Externos. Si es None,
            se omite la lectura de externos (útil para Casas a nivel servicio,
            cuya hoja específica se reporta por eje, no a nivel servicio).
        hoja_externos: nombre de la hoja específica del servicio en el archivo
            de externos (ej. "Jóvenes Con Oportunidades", "Forjar"). Si es
            None, se omite la lectura de externos.
        patron_responsable_ppdj: regex case-insensitive que identifica al
            servicio en la columna "Equipo(s) responsables de la implementación"
            del Excel de PPDJ. Si es None, no se filtra (devuelve todas las
            filas activas de PPDJ).
        tema_ppdj: nombre canónico de la PPDJ tal como aparece en la columna
            "Tema". Sirve para colocar la PPDJ siempre de primera al ordenar.
        pendientes_revision: lista de tuplas (tipo, tema, pendiente_con, nota)
            con los temas que Felipe mencionó pero aún no están en los Excels.
            Se muestran como tarjetas "En revisión". Si es None, no se agrega
            ningún pendiente.
        temas_externos_incluir: whitelist opcional para la lectura de
            Externos. Si se pasa, solo se incluyen las filas cuyo tema o
            producto contenga (case-insensitive) alguno de los strings de
            la lista. Útil para traer productos puntuales a una sección
            (ej. Afro 1.3.9 y PAD Víctimas a la sección servicio Casas).
            Si es None, no hay filtro.

    Devuelve lista de dicts con las claves: tipo, tema, codigo, producto,
    indicador, meta, periodicidad, responsable, observaciones, estado
    ('confirmado' | 'revision'), pendiente_con, nota_revision.
    """
    if pendientes_revision is None:
        pendientes_revision = []
    filas = []

    # 1) Lectura de PPDJ — filtra por servicio y por Activo=Sí. Si el
    # patrón es None, no se filtra por responsable.
    if os.path.exists(xlsx_ppdj):
        try:
            df = pd.read_excel(xlsx_ppdj, sheet_name="Subdirección para la Juventud")
            if patron_responsable_ppdj:
                mascara = (
                    df["Equipo(s) responsables de la implementación"].astype(str)
                    .str.contains(patron_responsable_ppdj, case=False, na=False, regex=True)
                )
            else:
                mascara = pd.Series([True] * len(df))
            activos = df["Activo"].astype(str).str.strip().str.lower() == "sí"
            for _, fila in df[mascara & activos].iterrows():
                codigo, producto = _separar_codigo_producto(
                    _limpiar_celda(fila.get("Producto esperado"))
                )
                filas.append({
                    "tipo": _limpiar_celda(fila.get("Tipo de reporte")) or "Política Pública",
                    "tema": _limpiar_celda(fila.get("Tema")),
                    "codigo": codigo,
                    "producto": producto,
                    "indicador": _limpiar_celda(fila.get("Indicador de producto")),
                    "meta": _limpiar_celda(fila.get("Meta 2026")),
                    "periodicidad": _limpiar_celda(fila.get("Cada cuánto se reporta")),
                    "responsable": _limpiar_celda(fila.get("Responsable del reporte")),
                    "observaciones": _limpiar_celda(fila.get("Observaciones")),
                    "estado": "confirmado",
                })
        except Exception as e:
            print(f"  ! No se pudo leer hoja PPDJ: {e}")

    # 2) Lectura de la hoja específica del servicio en Reportes Externos
    if xlsx_externos and hoja_externos and os.path.exists(xlsx_externos):
        try:
            df = pd.read_excel(xlsx_externos, sheet_name=hoja_externos)
            for _, fila in df.iterrows():
                tema_fila = _limpiar_celda(fila.get("Tema"))
                codigo, producto = _separar_codigo_producto(
                    _limpiar_celda(fila.get("Productos"))
                )
                if not producto:
                    continue
                # Whitelist opcional: si se pasó temas_externos_incluir,
                # solo se incluyen filas cuyo tema o producto contenga
                # alguno de los strings de la lista (case-insensitive).
                if temas_externos_incluir:
                    blanco = (tema_fila + " " + producto).lower()
                    if not any(t.lower() in blanco for t in temas_externos_incluir):
                        continue
                # En algunas hojas existe "Persona responsable del reporte"
                # y en otras solo "Equipo responsable del reporte". Se usa
                # la persona si está disponible, si no el equipo.
                persona = _limpiar_celda(fila.get("Persona responsable del reporte"))
                equipo = _limpiar_celda(fila.get("Equipo responsable del reporte"))
                responsable = persona if persona else equipo
                filas.append({
                    "tipo": _limpiar_celda(fila.get("Tipo de reporte")) or "Política Pública",
                    "tema": tema_fila,
                    "codigo": codigo,
                    "producto": producto,
                    "indicador": "",
                    "meta": _limpiar_celda(fila.get("Meta 2026")),
                    "periodicidad": _limpiar_celda(fila.get("Cada cuánto se reporta")),
                    "responsable": responsable,
                    "observaciones": _limpiar_celda(fila.get("Observaciones")),
                    "estado": "confirmado",
                })
        except Exception as e:
            print(f"  ! No se pudo leer hoja '{hoja_externos}': {e}")

    # 3) Agregar pendientes "En revisión"
    for tipo, tema, pendiente_con, nota in pendientes_revision:
        filas.append({
            "tipo": tipo,
            "tema": tema,
            "codigo": "",
            "producto": "",
            "indicador": "",
            "meta": "",
            "periodicidad": "",
            "responsable": "",
            "observaciones": "",
            "estado": "revision",
            "pendiente_con": pendiente_con,
            "nota_revision": nota,
        })

    # 4) Orden: PPDJ primero, luego otras Políticas, luego Planes. Dentro de
    # cada grupo, las confirmadas antes que las en revisión.
    def _clave(f):
        es_ppdj = 0 if f["tema"] == tema_ppdj else 1
        es_plan = 1 if f["tipo"].lower().startswith("plan") else 0
        es_revision = 1 if f["estado"] == "revision" else 0
        return (es_ppdj, es_plan, es_revision, f["tema"])

    filas.sort(key=_clave)
    return filas


def html_cards(filas):
    """Renderiza la lista de filas devuelta por leer_politicas() como HTML
    del grid de tarjetas. Una card por política o plan (no por producto):
    el header muestra el tema y el body lista todos los productos del Excel
    que pertenecen a ese tema. No se muestran metas, observaciones,
    responsables ni indicadores en este nivel (eso vive en el Excel y se
    consulta directamente cuando hace falta).

    Las tarjetas "En revisión" se muestran como una card simple con la
    nota corta de quién tiene pendiente la redacción.
    """
    # Agrupar filas por (estado, tipo, tema), preservando el orden de
    # llegada (ya viene ordenado por leer_politicas).
    grupos = []  # lista de tuplas (clave, items) para conservar orden
    indice = {}
    for f in filas:
        clave = (f["estado"], f["tipo"], f["tema"])
        if clave not in indice:
            indice[clave] = len(grupos)
            grupos.append((clave, []))
        grupos[indice[clave]][1].append(f)

    cards = []
    for (estado, tipo, tema), items in grupos:
        if estado == "revision":
            # Una pendiente trae una sola fila por tema (los pendientes se
            # declaran como entradas individuales). Tomamos la primera.
            f = items[0]
            # Si la entrada trae una nota propia (nota_revision), se usa
            # esa; si no, se muestra el "En revisión con {persona}".
            nota = f.get("nota_revision", "")
            if nota:
                cuerpo = f'<p class="jp-card-revision-nota">{nota}</p>'
            else:
                cuerpo = f'<p class="jp-card-revision-nota"><strong>En revisi&oacute;n con {f["pendiente_con"]}.</strong></p>'
            cards.append(
                '                    <article class="jp-card jp-card-pendiente">\n'
                '                        <div class="jp-card-header">\n'
                f'                            <h3 class="jp-card-tema">{tema}</h3>\n'
                '                        </div>\n'
                '                        <div class="jp-card-body">\n'
                f'                            {cuerpo}\n'
                '                        </div>\n'
                '                    </article>'
            )
            continue

        # Card confirmada: lista de productos del tema. Si un producto
        # tiene observación (descripción operativa de cómo se reporta),
        # se muestra como párrafo bajo el producto. Otros campos del
        # Excel (meta, responsable, indicador) no se muestran a nivel
        # de card — viven en el Excel y se consultan ahí cuando hace falta.
        productos_html = []
        for f in items:
            codigo = f["codigo"]
            prod = f["producto"]
            obs = f.get("observaciones", "")
            if codigo:
                cabeza = f'<span class="jp-card-codigo">{codigo}</span> {prod}'
            else:
                cabeza = prod
            extra = ""
            if obs:
                extra = f'\n                                    <p class="jp-card-prod-obs">{obs}</p>'
            productos_html.append(
                f'                                <li class="jp-card-producto">{cabeza}{extra}</li>'
            )
        lista_prods = "\n".join(productos_html)

        cards.append(
            '                    <article class="jp-card">\n'
            '                        <div class="jp-card-header">\n'
            f'                            <h3 class="jp-card-tema">{tema}</h3>\n'
            '                        </div>\n'
            '                        <div class="jp-card-body">\n'
            '                            <ul class="jp-card-productos">\n'
            f'{lista_prods}\n'
            '                            </ul>\n'
            '                        </div>\n'
            '                    </article>'
        )
    return "\n\n".join(cards)
