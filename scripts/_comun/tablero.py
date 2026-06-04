# -*- coding: utf-8 -*-
"""Lector del enlace del tablero de Power BI.

El enlace vive en enlaces/enlaces.xlsx (Hoja1), en la fila cuya SECCION es
"Tablero Power BI". De esta forma se gestiona desde el Excel y nunca queda
escrito dentro del codigo de los generadores.

Si existe una fila con HTML == nombre del servicio, esa tiene prioridad sobre
la fila general (HTML == "Todos"), por si algun servicio llega a tener un
tablero propio distinto al compartido.
"""
import os
import pandas as pd

_SECCION_POWERBI = "tablero power bi"


def powerbi_src(base_dir, servicio="Todos"):
    """Devuelve la URL del tablero de Power BI desde enlaces.xlsx.

    base_dir: raiz del proyecto (la carpeta que contiene enlaces/).
    servicio: valor de la columna HTML a priorizar (ej. "Forjar", "Alertas").
    Devuelve "" y avisa por consola si no encuentra el enlace.
    """
    ruta = os.path.join(base_dir, "enlaces", "enlaces.xlsx")
    if not os.path.exists(ruta):
        print("  ! enlaces.xlsx no encontrado: el tablero quedara sin enlace")
        return ""
    df = pd.read_excel(ruta, sheet_name="Hoja1")
    mask = df["SECCION"].astype(str).str.strip().str.lower() == _SECCION_POWERBI
    filas = df[mask]
    if filas.empty:
        print("  ! No hay fila 'Tablero Power BI' en enlaces.xlsx")
        return ""
    espec = filas[filas["HTML"].astype(str).str.strip().str.lower()
                  == str(servicio).strip().lower()]
    fila = espec.iloc[0] if not espec.empty else filas.iloc[0]
    return str(fila["ENLACE"]).strip()
