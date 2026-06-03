# -*- coding: utf-8 -*-
"""
Crea datos/equipo.xlsx, el libro unico que centraliza el equipo de los cuatro
servicios, una hoja por servicio. Cada hoja conserva la estructura propia del
servicio (no son columnas uniformes, porque cada pagina muestra el equipo
distinto).

Las hojas casas_juventud y forjar se copian identicas desde los Excel viejos,
que se conservan como respaldo. Las hojas jco y parche_seguro se crean aqui a
partir del contenido que hoy esta escrito directamente en los scripts.

Se ejecuta una sola vez. No vuelve a correrse en el flujo normal de generacion.
"""
import os
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(BASE, "datos")

# Hojas que ya existen en Excel: se copian tal cual para no alterar nada.
df_casas = pd.read_excel(os.path.join(DATOS, "equipo_casas_juventud.xlsx"))
df_forjar = pd.read_excel(os.path.join(DATOS, "equipo_forjar.xlsx"))

# Hoja JCO: organigrama de cuatro bloques que hoy vive en generar_gc_jco.py.
# Tipo de fila:
#   destacado -> bloque de coordinacion (nombre grande)
#   item      -> linea "Cargo: Nombre"
#   etiqueta  -> sub-equipo sin nombre propio (Nombre lleva el texto, Cargo vacio)
df_jco = pd.DataFrame(
    [
        (1, "Coordinación general", "destacado", "Líder del servicio", "Ana Catalina Suárez"),
        (2, "Equipo territorial", "item", "Referente psicosocial y de alertas", "Diana Lozano"),
        (2, "Equipo territorial", "item", "Referente logística y de FDL", "Pamela Barón"),
        (2, "Equipo territorial", "item", "Referente proceso formativo", "Alejandra Sosa Aponte"),
        (2, "Equipo territorial", "item", "Referente inclusión productiva", "Edgardo Montes"),
        (2, "Equipo territorial", "etiqueta", "", "Equipo psicosocial"),
        (2, "Equipo territorial", "etiqueta", "", "Equipo de alertas"),
        (3, "Equipo metodológico y gestión documental", "item", "Referente metodológico", "Ana María Altamar"),
        (3, "Equipo metodológico y gestión documental", "item", "Referente control político", "Alejandro Osorio"),
        (4, "Equipo administrativo", "item", "Líder administrativo", "John Garzón"),
        (4, "Equipo administrativo", "item", "Líder financiero", "David Quiceno"),
        (4, "Equipo administrativo", "item", "Gestión documental", "Andrés Rodríguez"),
        (4, "Equipo administrativo", "etiqueta", "", "3 apoyos"),
    ],
    columns=["Bloque", "Titulo_bloque", "Tipo", "Cargo", "Nombre"],
)

# Hoja Parche seguro: por ahora solo la persona de contacto. Se dejan las
# columnas Cargo y Correo para cuando el servicio entregue el equipo completo.
df_parche = pd.DataFrame(
    [("Paulla Nikoll Murillo Velandia", "", "")],
    columns=["Nombre", "Cargo", "Correo"],
)

salida = os.path.join(DATOS, "equipo.xlsx")
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    df_casas.to_excel(writer, sheet_name="casas_juventud", index=False)
    df_forjar.to_excel(writer, sheet_name="forjar", index=False)
    df_jco.to_excel(writer, sheet_name="jco", index=False)
    df_parche.to_excel(writer, sheet_name="parche_seguro", index=False)

print("Creado:", salida)
print("Hojas: casas_juventud(%d) forjar(%d) jco(%d) parche_seguro(%d)" % (
    len(df_casas), len(df_forjar), len(df_jco), len(df_parche)))
