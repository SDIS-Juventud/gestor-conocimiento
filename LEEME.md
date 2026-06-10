# Gestor de Conocimiento - Subdirección para la Juventud (SDIS)

Este proyecto genera los HTML del gestor de conocimiento de los cuatro servicios de la Subdirección para la Juventud: Casas de Juventud, Forjar, Jóvenes con Oportunidades (JCO) y Alertas.

Todos los HTML de la raíz son **generados automáticamente** por scripts de Python. No se deben editar a mano: si se editan, los cambios se pierden la próxima vez que se regenere el archivo.

## Estructura del proyecto

```
gestor-conocimiento-2025/
├── datos/              ← insumos (Excels, GeoJSON). Sí se editan a mano.
├── scripts/            ← generadores Python. Se editan cuando cambia la lógica.
│   └── _comun/         ← CSS y helpers compartidos entre generadores.
├── imagenes/           ← imágenes usadas por los HTML. Sí se editan a mano.
├── ejes/               ← HTML de los ejes de Casas de Juventud (por ahora a mano).
├── archivo/            ← versiones históricas y artefactos jubilados. Solo lectura.
├── Notes/              ← vault Obsidian del proyecto. No se sube a git.
│
├── index.html                              ← generado, NO editar
├── gestion_conocimiento_juventud_2025.html ← generado, NO editar
├── gestion_conocimiento_forjar_2025.html   ← generado, NO editar
├── gestion_conocimiento_jco_2025.html      ← generado, NO editar
├── gestion_conocimiento_alertas_2025.html  ← generado, NO editar
├── mapa_casas_juventud.html                ← generado, NO editar
└── mapa_forjar.html                        ← generado, NO editar
```

## Qué genera qué

| Salida | Generador | Insumos |
|---|---|---|
| `index.html` | `scripts/generar_home_servicios.py` | (contenido hardcoded en el script) |
| `gestion_conocimiento_juventud_2025.html` | `scripts/generar_juventud.py` | `datos/equipo_casas_juventud.xlsx`, `datos/directorio_casas_juventud.xlsx`, base SIRBE limpia, `datos/localidades_bogota.geojson` |
| `gestion_conocimiento_forjar_2025.html` | `scripts/generar_gc_forjar.py` | `datos/equipo_forjar.xlsx`, `datos/directorio_forjar.xlsx`, `datos/forjar_proceso_operativo.xlsx`, `enlaces/enlaces.xlsx` (filas con HTML="Forjar") |
| `gestion_conocimiento_jco_2025.html` | `scripts/generar_gc_jco.py` | (contenido hardcoded) |
| `gestion_conocimiento_alertas_2025.html` | `scripts/generar_gc_alertas.py` | `enlaces/enlaces.xlsx` (filas con HTML="Alertas" y SECCION="Protocolos - ...") |
| `mapa_casas_juventud.html` | `scripts/generar_juventud.py` (efecto colateral) | `datos/directorio_casas_juventud.xlsx`, `datos/localidades_bogota.geojson` |
| `mapa_forjar.html` | `scripts/generar_gc_forjar.py` (efecto colateral) | `datos/directorio_forjar.xlsx`, `datos/localidades_bogota.geojson` |

## Cómo hacer cambios comunes

### Cambiar el equipo de Casas de Juventud

1. Abrir `datos/equipo_casas_juventud.xlsx` en Excel.
2. Editar la fila del cambio (Nombre, Cargo, etc.).
3. Guardar.
4. Correr: `python scripts/generar_juventud.py`
5. Abrir `gestion_conocimiento_juventud_2025.html` en el navegador para verificar.

### Cambiar el directorio de casas o forjar

1. Editar `datos/directorio_casas_juventud.xlsx` o `datos/directorio_forjar.xlsx`.
2. Regenerar:
   - Para casas: `python scripts/generar_juventud.py`
   - Para forjar: `python scripts/generar_gc_forjar.py` (regenera HTML + `mapa_forjar.html`)

### Cambiar contenido de texto de juventud / forjar / jco / alertas

El contenido está hardcoded en el Python correspondiente (excepto equipo y directorio, que están en Excel). Para cambiar un texto:

1. Abrir el `.py` del servicio en `scripts/`.
2. Buscar el texto, editarlo dentro del string de Python.
3. Correr el script.

**Importante:** nunca editar el `.html` directamente. El cambio se pierde la siguiente regeneración.

### Cambiar la pestaña "Proceso operativo" de Forjar

La pestaña Proceso operativo de Forjar lee de dos Excels:

- `datos/forjar_proceso_operativo.xlsx` — tres hojas:
    - `etapas` (id, titulo, subtitulo, descripcion_breve, descripcion_panel)
    - `actividades` (etapa_id, actividad, registro, responsable)
    - `actores` (nombre, descripcion)
- `enlaces/enlaces.xlsx` hoja "generales" — para los enlaces a documentos. Filas con HTML="Forjar" y SECCION="Proceso Operativo - <Etapa> - <Nombre> (<código opcional>)". Las etapas válidas son Ingreso, Permanencia, Egreso, Post-Egreso (RAE).

Para agregar un nuevo documento o formulario:
1. Abrir `enlaces/enlaces.xlsx` hoja "generales".
2. Agregar una fila nueva: HTML="Forjar", SECCION siguiendo el patrón, ENLACE con la URL.
3. Si la URL es de `forms.office.com`, el script la marca como formulario; cualquier otra URL se marca como descarga.
4. Correr `python scripts/generar_gc_forjar.py`.

Para cambiar el texto de una etapa, una actividad o un actor: editar la celda correspondiente en `datos/forjar_proceso_operativo.xlsx` y regenerar.

**Lo que vive en el Python (no en Excel):** colores de cada etapa y de los badges de responsables, mapeo del nombre del responsable a su grupo de color, y los íconos de las cards de documentos. No editar en el Excel.

### Cambiar la pestaña "Protocolos" de Alertas

La pestaña Protocolos del HTML de Alertas lee de `enlaces/enlaces.xlsx` hoja "generales", filtrando filas con HTML="Alertas" y SECCION que empiece con "Protocolos - ".

Para agregar un nuevo protocolo:
1. Abrir `enlaces/enlaces.xlsx` hoja "generales".
2. Agregar una fila nueva: HTML="Alertas", SECCION="Protocolos - Nombre del documento", ENLACE con la URL del PDF.
3. Correr `python scripts/generar_gc_alertas.py`.

El script muestra el nombre del documento (lo que va después de "Protocolos - ") como título de cada card, y un botón "Descargar PDF" enlazado a la URL.

### Cambiar estilos (colores, tipografía)

Los estilos comunes viven en `scripts/_comun/estilos.py` (después de la refactorización de abril 2026). Cambiar ahí y regenerar los cuatro servicios + el home.

### Cambiar los HTML de los ejes (Bienestar, Cultura, Inclusión, Liderazgo, SIDICU)

Por ahora estos HTML se editan a mano en `ejes/`. La conversión a generador está pendiente — ver `Notes/arquitectura-proyecto.md`.

## Regenerar todo

Cuando hay cambios que afectan al home o al CSS compartido, conviene regenerar todo en este orden:

```bash
cd gestor-conocimiento-2025
python scripts/generar_home_servicios.py     # index.html
python scripts/generar_juventud.py           # juventud + mapa_casas
python scripts/generar_gc_forjar.py          # forjar + mapa_forjar
python scripts/generar_gc_jco.py             # jco
python scripts/generar_gc_alertas.py         # alertas
```

## Qué está en `archivo/`

Carpeta de solo lectura con artefactos jubilados:

- `archivo/notebooks_viejos/` — notebooks que ya no se usan (`generar_gc_juventud.ipynb`, reemplazado por `generar_juventud.py`).
- `archivo/scripts_viejos/` — versiones anteriores de scripts (`generar_home_servicios_v1.py`).
- `archivo/html_historicos/` — backups de los HTML de casas de juventud (V1 a V13).
- `archivo/ejes_manuales_backup/` — backup de los HTML de ejes antes de convertirlos a generador.

Nada de esto debe correrse ni editarse. Está aquí solo por si necesitas consultar el estado histórico.

## Documentación adicional

Ver `Notes/` para notas detalladas sobre decisiones técnicas, lecciones aprendidas, y contexto del proyecto. En particular:

- `Notes/arquitectura-proyecto.md` — inventario completo del proyecto y diagnóstico.
- `Notes/decision-migrar-notebooks-a-py.md` — por qué migramos los notebooks a `.py`.
