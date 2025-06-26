import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO

def extract_tables_from_pptx(pptx_file):
    prs = Presentation(pptx_file)
    tables = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                data = []
                for row in table.rows:
                    data.append([cell.text.strip() for cell in row.cells])
                df = pd.DataFrame(data[1:], columns=data[0])
                tables.append((f"Diapositiva {i+1}", df))
    return tables

def extract_tables_from_excel(excel_file):
    xls = pd.ExcelFile(excel_file, engine='openpyxl')
    tables = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')
        tables.append((sheet_name, df))
    return tables

def compare_dataframes(df1, df2):
    try:
        df1_sorted = df1.sort_index(axis=1).sort_values(by=df1.columns[0], ignore_index=True)
        df2_sorted = df2.sort_index(axis=1).sort_values(by=df2.columns[0], ignore_index=True)
        return df1_sorted.equals(df2_sorted)
    except Exception:
        return False

def main():
    st.title("Comparador de Tablas: PowerPoint vs Excel")

    pptx_file = st.file_uploader("Carga tu archivo PowerPoint (.pptx)", type="pptx")
    excel_file = st.file_uploader("Carga tu archivo Excel (.xlsx)", type="xlsx")

    if pptx_file and excel_file:
        st.success("Archivos cargados correctamente. Procesando...")

        pptx_tables = extract_tables_from_pptx(pptx_file)
        excel_tables = extract_tables_from_excel(excel_file)

        if not pptx_tables:
            st.warning("No se encontraron tablas en el archivo PowerPoint.")
            return

        if not excel_tables:
            st.warning("No se encontraron hojas de datos en el archivo Excel.")
            return

        st.header("Resultados de la Comparación")

        for i, (slide_name, ppt_df) in enumerate(pptx_tables):
            match_found = False
            for sheet_name, excel_df in excel_tables:
                if compare_dataframes(ppt_df, excel_df):
                    st.success(f"✅ {slide_name} coincide con la hoja '{sheet_name}' del Excel.")
                    match_found = True
                    break
            if not match_found:
                st.error(f"❌ {slide_name} no coincide con ninguna hoja del Excel.")
            with st.expander(f"Ver tabla de {slide_name}"):
                st.dataframe(ppt_df)

if __name__ == "__main__":
    main()
