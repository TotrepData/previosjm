# Generador de Documentos

Aplicación web para la generación automatizada de documentos Word a partir de datos en Excel. La herramienta utiliza plantillas parametrizadas para crear múltiples documentos con datos específicos de manera eficiente.

## Descripción

La aplicación permite automatizar la creación de documentos mediante:

- Carga de un archivo Excel con datos a procesar
- Carga de una plantilla Word con placeholders
- Generación masiva de documentos con reemplazo de variables
- Descarga de documentos generados en formato ZIP

Este flujo es especialmente útil para procesos administrativos que requieren generar grandes volúmenes de documentos personalizados.

## Requisitos

- Python 3.8+
- pip

## Instalación

1. Clona el repositorio:
```bash
git clone <url-del-repo>
cd <directorio-del-proyecto>
```

2. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## Uso

### Ejecución local

```bash
streamlit run app.py
```

La aplicación se abrirá en `http://localhost:8501`

### Flujo de uso

1. **Cargar datos**: Sube un archivo Excel (.xlsx) que contiene los datos a procesar
2. **Cargar plantilla**: Sube una plantilla Word (.docx) con placeholders en formato `{{nombre_columna}}`
3. **Generar documentos**: Haz clic en el botón "Generar Documentos"
4. **Descargar resultados**: Descarga el archivo ZIP con todos los documentos generados

### Formato de placeholders

Los placeholders en la plantilla Word deben seguir el formato:

```
{{nombre_columna}}
```

Donde `nombre_columna` coincide exactamente con los nombres de las columnas en el Excel.

Ejemplo:
- Excel con columnas: `nombre`, `empresa`, `fecha_inicio`
- Plantilla Word: `Estimado {{nombre}}, trabajará en {{empresa}} a partir del {{fecha_inicio}}`

## Características

- Reemplazo de placeholders en párrafos y tablas
- Procesamiento de múltiples documentos
- Manejo de errores con reportes detallados
- Generación de archivos ZIP descargables
- Vista previa de datos antes de procesar
- Advertencias para volúmenes grandes (>1000 registros)

## Estructura del código

- `replace_text_in_paragraph()`: Función que reemplaza placeholders manteniendo el formato original
- `generar_documentos()`: Función principal que procesa Excel y genera documentos
- Interfaz Streamlit: Componentes para carga de archivos y generación de documentos

## Limitaciones y consideraciones

- La aplicación procesa un máximo de 1000 documentos sin advertencia
- Los placeholders deben estar completos en un único run de texto en Word
- Se mantiene el formato original de la plantilla en los documentos generados
- Los errores en filas individuales no detienen el procesamiento general

## Tecnologías utilizadas

- **Streamlit**: Framework para interfaz web
- **pandas**: Procesamiento de datos Excel
- **python-docx**: Manipulación de documentos Word
- **zipfile**: Compresión de archivos generados

## Despliegue en Streamlit Cloud

Para desplegar en Streamlit Cloud:

1. Sube el repositorio a GitHub
2. Accede a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu repositorio GitHub
4. Selecciona `app.py` como archivo principal
5. La aplicación se desplegará automáticamente en cada push

## Licencia

Javier Mondragón
