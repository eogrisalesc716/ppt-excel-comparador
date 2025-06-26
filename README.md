# Comparador de Datos: PowerPoint vs Excel

Esta aplicación permite comparar automáticamente los datos contenidos en **gráficos (charts)** de un archivo PowerPoint (`.pptx`) con bloques de datos en un archivo Excel (`.xlsx`), utilizando **marcadores de texto** como referencia para vincular cada gráfico con su bloque correspondiente.

## Características

- Extrae datos de gráficos (charts) en las diapositivas del PowerPoint.
- Detecta marcadores como "Indicador 1", "Indicador 2", etc., para vincular cada gráfico con su bloque de datos en Excel.
- Compara los datos del gráfico con los datos del Excel, ignorando el orden de filas o columnas.
- Muestra visualmente si los datos coinciden o no.
- Interfaz interactiva construida con Streamlit.

## ¿Cómo usar esta app?

### 1. Clona este repositorio o sube los archivos a GitHub

Asegúrate de que el repositorio contenga los siguientes archivos:
- `comparador.py`
- `requirements.txt`
- `README.md`

### 2. Despliega en Streamlit Cloud

1. Ve a [https://streamlit.io/cloud](https://streamlit.io/cloud)
2. Inicia sesión con tu cuenta de GitHub.
3. Selecciona el repositorio que contiene este proyecto.
4. Asegúrate de que el archivo principal sea `comparador.py`.
5. Haz clic en **Deploy**.

### 3. Usa la app

Una vez desplegada, podrás:
- Cargar un archivo `.pptx` con gráficos.
- Cargar un archivo `.xlsx` con los datos fuente.
- Ver si los datos de cada gráfico coinciden con los datos del Excel.

## Requisitos

Consulta el archivo `requirements.txt` para ver las dependencias necesarias.

## Licencia

MIT
