# mcp-pptx-server (pptx-editor)

Un servidor MCP (Model Context Protocol) diseñado para crear, leer, editar y dar formato a presentaciones de PowerPoint (`.pptx`) utilizando `python-pptx`. Este servidor permite a agentes de IA interactuar de manera nativa con archivos de PowerPoint a través de una API limpia y estructurada.

## Características

- **Ciclo de vida completo**: Crear, abrir, guardar, listar y cerrar presentaciones en memoria.
- **Lectura profunda**: Inspeccionar la estructura, leer textos, tablas y notas del presentador de forma individual o global, y realizar búsquedas de texto.
- **Edición completa**: Modificar títulos, textos de formas individuales, notas del presentador, realizar buscar-y-reemplazar global y reordenar o eliminar slides.
- **Creación de contenido**: Insertar nuevos slides con layouts específicos, cajas de texto personalizadas, imágenes de disco y tablas estructuradas con datos.
- **Formato y estilos**: Personalizar fuentes (familia, tamaño, negrita, cursiva, color) en formas y configurar colores de fondo sólidos para los slides.

---

## Estructura del Proyecto

```
mcp_pptx_server/
├── __init__.py
├── __main__.py
├── server.py                 # Punto de entrada FastMCP con registro de herramientas
├── presentation_manager.py   # Gestión del ciclo de vida en memoria (create/open/save/close)
├── presentation_reader.py    # Lectura e inspección profunda de diapositivas
├── presentation_editor.py    # Edición, reemplazo, eliminación y duplicación
├── slide_builder.py          # Adición de diapositivas, textos, imágenes y tablas
└── formatting.py             # Formateo de fuentes y colores de fondo
pyproject.toml                # Metadatos del paquete y dependencias
requirements.txt              # Requerimientos de pip
README.md                     # Documentación de referencia
.gitignore                    # Configuración de exclusiones de git
```

---

## Requisitos e Instalación

### Requisitos

- Python `>= 3.10`

### Instalación local

1. Clonar o descargar este repositorio en su máquina local.
2. Crear un entorno virtual e instalar el paquete en modo editable:

```bash
python -m venv .venv
.\.venv\Scripts\activate  # En Windows
pip install -e .
```

---

## Configuración del Servidor MCP

Para usar este servidor con clientes MCP compatibles (como Claude Desktop, VS Code o cursores de IA), configure la conexión stdio.

### Configuración de Claude Desktop

Añada lo siguiente al archivo `claude_desktop_config.json` (usualmente ubicado en `%APPDATA%\Claude\claude_desktop_config.json` en Windows):

```json
{
  "mcpServers": {
    "pptx-editor": {
      "command": "C:\\Users\\moise\\Documents\\014_mcp_ppt\\.venv\\Scripts\\python.exe",
      "args": ["-m", "mcp_pptx_server"],
      "cwd": "C:\\Users\\moise\\Documents\\014_mcp_ppt"
    }
  }
}
```

> **Nota:** Reemplace las rutas con las rutas absolutas correspondientes en su sistema.

---

## Tabla de Herramientas Disponibles

El servidor registra las siguientes herramientas bajo la arquitectura MCP:

### 1. Gestión del Ciclo de Vida

| Herramienta | Parámetros | Descripción |
|---|---|---|
| `create_presentation` | `prs_id` (str), `title` (str, opcional) | Crea una presentación vacía en memoria con un ID. |
| `open_presentation` | `prs_id` (str), `file_path` (str) | Abre una presentación `.pptx` existente desde el disco. |
| `save_presentation` | `prs_id` (str), `file_path` (str, opcional) | Guarda la presentación abierta. Si no se indica ruta, usa la de origen. |
| `close_presentation`| `prs_id` (str) | Cierra y libera de la memoria la presentación especificada. |
| `list_presentations`| Ninguno | Devuelve la lista de presentaciones que están abiertas en memoria. |

### 2. Inspección y Lectura

| Herramienta | Parámetros | Descripción |
|---|---|---|
| `get_presentation_structure` | `prs_id` (str) | Retorna la lista de slides con índice, nombre del layout, título y cantidad de formas. |
| `read_slide` | `prs_id` (str), `slide_index` (int) | Detalle de un slide: formas con tipo, texto, posiciones (cm) y datos de tablas. |
| `read_slide_notes` | `prs_id` (str), `slide_index` (int) | Obtiene el texto de las notas del orador de un slide. |
| `read_full_text` | `prs_id` (str) | Extrae de forma estructurada todo el texto de la presentación. |
| `list_slides` | `prs_id` (str) | Resumen rápido (lista simplificada) de todos los slides. |
| `search_text` | `prs_id` (str), `query` (str), `case_sensitive` (bool) | Busca un término en textos de formas, tablas y notas de la presentación. |

### 3. Edición de Diapositivas

| Herramienta | Parámetros | Descripción |
|---|---|---|
| `edit_shape_text` | `prs_id` (str), `slide_index` (int), `shape_name_or_index` (str/int), `new_text` (str) | Reemplaza el texto de una forma específica (soporta saltos de línea). |
| `edit_slide_title` | `prs_id` (str), `slide_index` (int), `new_title` (str) | Cambia el título de la diapositiva usando detección de marcadores. |
| `edit_slide_notes` | `prs_id` (str), `slide_index` (int), `new_text` (str) | Modifica o añade las notas de presentador de una diapositiva. |
| `replace_text` | `prs_id` (str), `search` (str), `replace` (str), `case_sensitive` (bool) | Realiza un Buscar y Reemplazar global en toda la presentación. |
| `delete_slide` | `prs_id` (str), `slide_index` (int) | Elimina permanentemente una diapositiva reajustando índices. |
| `duplicate_slide` | `prs_id` (str), `slide_index` (int) | Duplica una diapositiva copiando contenido, tablas, imágenes y notas. |
| `move_slide` | `prs_id` (str), `from_index` (int), `to_index` (int) | Reordena la posición de un slide (mueve del índice origen al destino). |

### 4. Creación y Adición de Contenido

| Herramienta | Parámetros | Descripción |
|---|---|---|
| `add_slide` | `prs_id` (str), `layout_index` (int), `title` (str), `content` (str) | Inserta un slide utilizando un layout específico de la plantilla. |
| `add_text_box` | `prs_id` (str), `slide_index` (int), `text` (str), `left_cm` (float), `top_cm` (float), `width_cm` (float), `height_cm` (float) | Añade un cuadro de texto flotante en las coordenadas indicadas (cm). |
| `add_image` | `prs_id` (str), `slide_index` (int), `image_path` (str), `left_cm` (float), `top_cm` (float), `width_cm` (float, opcional), `height_cm` (float, opcional) | Inserta una imagen de disco con opción de auto-escala. |
| `add_table` | `prs_id` (str), `slide_index` (int), `rows` (int), `cols` (int), `left_cm` (float), `top_cm` (float), `width_cm` (float), `height_cm` (float), `data` (list, opcional) | Inserta una tabla matricial opcionalmente pre-poblada con datos. |
| `list_layouts` | `prs_id` (str) | Muestra los diseños de diapositivas disponibles con su índice y nombre. |

### 5. Formato y Apariencia

| Herramienta | Parámetros | Descripción |
|---|---|---|
| `set_shape_font` | `prs_id` (str), `slide_index` (int), `shape_name_or_index` (str/int), `font_name` (str), `font_size` (float), `bold` (bool), `italic` (bool), `color_hex` (str) | Configura familia, tamaño, negrita, cursiva y color de fuente de una forma. |
| `set_slide_background` | `prs_id` (str), `slide_index` (int), `color_hex` (str) | Establece un color de fondo sólido (hexadecimal) para una diapositiva. |

---

## Ejemplos de Flujo de Trabajo

### Crear una presentación corporativa desde cero

1. **Crear la presentación:**
   `create_presentation(prs_id="ventas_2026", title="Resultados Q1 2026")`
2. **Consultar layouts disponibles:**
   `list_layouts(prs_id="ventas_2026")`
3. **Agregar un slide de contenido:**
   `add_slide(prs_id="ventas_2026", layout_index=1, title="Métricas de Crecimiento", content="• Ventas incrementadas un 25% YoY\n• Costo de adquisición de clientes disminuyó 12%\n• Retención de usuarios subió al 94%")`
4. **Dar formato al título del slide corporativo:**
   `set_shape_font(prs_id="ventas_2026", slide_index=0, shape_name_or_index=0, font_name="Georgia", font_size=40, bold=True, color_hex="003366")`
5. **Establecer un color de fondo sólido para el slide de título (Gris claro):**
   `set_slide_background(prs_id="ventas_2026", slide_index=0, color_hex="F4F4F9")`
6. **Guardar la presentación en disco:**
   `save_presentation(prs_id="ventas_2026", file_path="C:\\Presentaciones\\Reporte_Ventas_Q1.pptx")`
7. **Cerrar la sesión de memoria:**
   `close_presentation(prs_id="ventas_2026")`

## Licencia

Este software se proporciona bajo la licencia MIT.
# ppt-mcp
