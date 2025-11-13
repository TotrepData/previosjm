import streamlit as st
import pandas as pd
from docx import Document
import zipfile
import io
from datetime import datetime

st.set_page_config(page_title="Generador de Documentos", layout="wide")
st.title("Generador de documentos - Atenea")
st.markdown("Equipo de transformación digital - Beta")

def replace_text_in_paragraph(paragraph, key, value):
    """Reemplaza placeholders en párrafos manteniendo formato"""
    placeholder = "{{" + key + "}}"
    value_str = str(value) if value is not None else ""
    
    full_text = "".join(run.text for run in paragraph.runs)
    
    if placeholder in full_text:
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, value_str)
                return

def generar_documentos(df, word_file):
    """Genera documentos Word a partir de Excel"""
    try:
        zip_buffer = io.BytesIO()
        documentos_generados = 0
        errores = []
        
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df.iterrows():
                try:
                    word_file.seek(0)
                    doc = Document(word_file)
                    
                    for paragraph in doc.paragraphs:
                        for key in row.index:
                            replace_text_in_paragraph(paragraph, key, row[key])
                    
                    for table in doc.tables:
                        for row_table in table.rows:
                            for cell in row_table.cells:
                                for paragraph in cell.paragraphs:
                                    for key in row.index:
                                        replace_text_in_paragraph(paragraph, key, row[key])
                    
                    doc_bytes = io.BytesIO()
                    doc.save(doc_bytes)
                    zipf.writestr(f"Documento_{idx + 1}.docx", doc_bytes.getvalue())
                    documentos_generados += 1
                    
                except Exception as e:
                    errores.append(f"Fila {idx + 1}: {str(e)}")
                    
        return zip_buffer, documentos_generados, errores
        
    except Exception as e:
        raise Exception(f"Error procesando documentos: {str(e)}")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Carga tu EXCEL (Base.xlsx)", type="xlsx")
with col2:
    word_file = st.file_uploader("Carga tu PLANTILLA (Plantilla.docx)", type="docx")

if excel_file and word_file:
    st.success("Archivos cargados correctamente")
    
    with st.expander("Ver datos del Excel"):
        try:
            df_preview = pd.read_excel(excel_file)
            st.dataframe(df_preview, use_container_width=True)
            st.info(f"Total de filas: {len(df_preview)}")
        except Exception as e:
            st.error(f"Error al leer Excel: {e}")
    
    if st.button("Generar Documentos", use_container_width=True, type="primary"):
        try:
            excel_file.seek(0)
            df = pd.read_excel(excel_file)
            
            if df.empty:
                st.error("El Excel está vacío")
            elif len(df) > 1000:
                st.warning(f"Tienes {len(df)} filas. Esto puede tardar...")
            
            with st.spinner(f"Procesando {len(df)} documentos..."):
                zip_buffer, generados, errores = generar_documentos(df, word_file)
                
                if generados > 0:
                    st.success(f"Se generaron {generados} documentos correctamente")
                    
                    if errores:
                        with st.expander(f"Errores en {len(errores)} documentos"):
                            for error in errores:
                                st.warning(error)
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label=f"Descargar ZIP ({generados} documentos)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Documentos_Generados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                        type="primary"
                    )
                else:
                    st.error("No se pudieron generar documentos. Revisa los errores abajo.")
                    with st.expander("Ver errores"):
                        for error in errores:
                            st.error(error)
                            
        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.info("Verifica que: 1) El Excel tenga datos, 2) La plantilla tenga placeholders {{columna}}")
            
else:
    st.warning("Por favor carga ambos archivos para continuar")
    st.info("Pasos: 1) Carga un Excel con los datos, 2) Carga una plantilla Word con placeholders {{columna}}, 3) Haz clic en Generar")
