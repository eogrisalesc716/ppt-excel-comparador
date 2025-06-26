import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO

def extract_chart_data_from_pptx(pptx_file):
    prs = Presentation(pptx_file)
    charts = []
    for i, slide in enumerate(prs.slides):
        slide_title = None
        for shape in slide.shapes:
            if shape.has_text_frame and not slide_title:
                text = shape.text.strip()
                if text.lower().startswith("indicador"):
                    slide_title = text
            if shape.has_chart:
                chart = shape.chart
                data = []
                categories = [c.label for c in chart.plots[0].categories]
                for series in chart.series:
                    name = series.name
                    values = series.values
                    row = [name] + list(values)
                    data.append(row)
                df = pd.DataFrame(data, columns=["Marca"] + categories)
                charts.append((slide_title or f"Diapositiva {i+1}", df))
    return charts

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
    st.title("Comparador de Gráficos: PowerPoint vs Excel")

    pptx_file = st.file_uploader("Carga tu archivo PowerPoint (.pptx)", type="pptx")
    excel_file = st.file_uploader("Carga tu archivo Excel (.xlsx)", type="xlsx")

    if pptx_file and excel_file:
        st.success("Archivos cargados correctamente. Procesando...")

        pptx_charts = extract_chart_data_from_pptx(pptx_file)
        excel_tables = extract_tables_from_excel(excel_file)

        if not pptx_charts:
            st.warning("No se encontraron gráficos con datos en el archivo PowerPoint.")
            return

        if not excel_tables:
            st.warning("No se encontraron hojas de datos en el archivo Excel.")
            return

        st.header("Resultados de la Comparación")

        for slide_title, chart_df in pptx_charts:
            match_found = False
            for sheet_name, excel_df in excel_tables:
                if slide_title.lower() in sheet_name.lower() or sheet_name.lower() in slide_title.lower():
                    if compare_dataframes(chart_df, excel_df):
                        st.success(f"✅ {slide_title} coincide con la hoja '{sheet_name}' del Excel.")
                    else:
                        st.error(f"❌ {slide_title} no coincide con la hoja '{sheet_name}' del Excel.")
                    match_found = True
                    break
            if not match_found:
                st.warning(f"⚠️ No se encontró una hoja en Excel que coincida con el marcador '{slide_title}'.")
            with st.expander(f"Ver datos del gráfico: {slide_title}"):
                st.dataframe(chart_df)

if __name__ == "__main__":
    main()
