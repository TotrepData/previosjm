import streamlit as st
import pandas as pd
from docx import Document
import zipfile
import io
from datetime import datetime

st.set_page_config(page_title="Generador de Documentos", layout="wide")
st.title("Generador de Documentos Word")
st.markdown("Carga tu Excel y tu plantilla, y genera documentos automaticamente")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Carga tu EXCEL (Base.xlsx)", type="xlsx")

with col2:
    word_file = st.file_uploader("Carga tu PLANTILLA (Plantilla.docx)", type="docx")

if excel_file and word_file:
    st.success("Archivos cargados correctamente")
    
    if st.button("Generar Documentos", use_container_width=True):
        df = pd.read_excel(excel_file)
        st.info(f"Procesando {len(df)} filas...")
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df.iterrows():
                fecha = datetime.now().strftime("%d/%m/%Y")
                
                doc = Document(word_file)
                
                for paragraph in doc.paragraphs:
                    for key in row.index:
                        placeholder = "{{" + key + "}}"
                        paragraph.text = paragraph.text.replace(placeholder, str(row[key]))
                
                for table in doc.tables:
                    for row_table in table.rows:
                        for cell in row_table.cells:
                            for paragraph in cell.paragraphs:
                                for key in row.index:
                                    placeholder = "{{" + key + "}}"
                                    paragraph.text = paragraph.text.replace(placeholder, str(row[key]))
                
                doc_bytes = io.BytesIO()
                doc.save(doc_bytes)
                zipf.writestr(f"Documento_{idx + 1}.docx", doc_bytes.getvalue())
        
        zip_buffer.seek(0)
        
        st.success(f"Se generaron {len(df)} documentos correctamente")
        st.download_button(
            label="Descargar ZIP con documentos",
            data=zip_buffer.getvalue(),
            file_name="Documentos_Generados.zip",
            mime="application/zip",
            use_container_width=True
        )
else:
    st.warning("Por favor carga ambos archivos para continuar")
