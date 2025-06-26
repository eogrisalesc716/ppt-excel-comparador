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

def compare_dataframes_by_index(df1, df2):
    df1_indexed = df1.set_index(df1.columns[0])
    df2_indexed = df2.set_index(df2.columns[0])
    df1_indexed = df1_indexed.sort_index().sort_index(axis=1)
    df2_indexed = df2_indexed.sort_index().sort_index(axis=1)

    differences = []
    for row_label in df1_indexed.index:
        if row_label in df2_indexed.index:
            for col_label in df1_indexed.columns:
                if col_label in df2_indexed.columns:
                    val1 = df1_indexed.loc[row_label, col_label]
                    val2 = df2_indexed.loc[row_label, col_label]
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    if val1 != val2:
                        differences.append({
                            "Marca": row_label,
                            "Categoría": col_label,
                            "Valor PPT": val1,
                            "Valor Excel": val2
                        })
                else:
                    differences.append({
                        "Marca": row_label,
                        "Categoría": col_label,
                        "Valor PPT": df1_indexed.loc[row_label, col_label],
                        "Valor Excel": "No encontrado"
                    })
        else:
            for col_label in df1_indexed.columns:
                differences.append({
                    "Marca": row_label,
                    "Categoría": col_label,
                    "Valor PPT": df1_indexed.loc[row_label, col_label],
                    "Valor Excel": "Marca no encontrada"
                })
    return differences

def main():
    st.title("Comparador de Gráficos: PowerPoint vs Excel (Indexado por Identificadores)")

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

        all_differences = []

        for slide_title, chart_df in pptx_charts:
            match_found = False
            for sheet_name, excel_df in excel_tables:
                if slide_title.lower() in sheet_name.lower() or sheet_name.lower() in slide_title.lower():
                    differences = compare_dataframes_by_index(chart_df, excel_df)
                    if not differences:
                        st.success(f"✅ {slide_title} coincide con la hoja '{sheet_name}' del Excel.")
                    else:
                        st.error(f"❌ {slide_title} tiene diferencias con la hoja '{sheet_name}':")
                        df_diff = pd.DataFrame(differences)
                        st.dataframe(df_diff)
                        for d in differences:
                            d["Gráfico"] = slide_title
                            d["Hoja Excel"] = sheet_name
                        all_differences.extend(differences)
                    match_found = True
                    break
            if not match_found:
                st.warning(f"⚠️ No se encontró una hoja en Excel que coincida con el marcador '{slide_title}'.")
            with st.expander(f"Ver datos del gráfico: {slide_title}"):
                st.dataframe(chart_df)

        if all_differences:
            df_all = pd.DataFrame(all_differences)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_all.to_excel(writer, index=False, sheet_name="Diferencias")
            st.download_button(
                label="📥 Descargar diferencias como Excel",
                data=output.getvalue(),
                file_name="diferencias_ppt_vs_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
