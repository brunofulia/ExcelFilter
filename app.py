import streamlit as st
import pandas as pd
from io import BytesIO

# Título de la app
st.title("Filtrar y Guardar Tabla de Excel")

# Subida del archivo de Excel
uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])

# Verificar si se ha subido un archivo
if uploaded_file:
    # Leer la hoja de Excel
    df = pd.read_excel(uploaded_file)

    # Mostrar la tabla completa
    st.subheader("Tabla Completa")
    st.dataframe(df)

    # Inicialización de filtros
    filtros = []
    criterios = []

    # Número de filtros
    num_filtros = st.number_input(
        "Número de filtros", min_value=1, step=1, value=1, key="num_filtros"
    )

    for i in range(num_filtros):
        with st.container():
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                column = st.selectbox("Columna", df.columns, key=f"col_{i}")

            with col2:
                if df[column].dtype in ["int64", "float64"]:
                    filter_criteria = [
                        "Mayor que",
                        "Menor que",
                        "Igual a",
                        "Diferente de",
                        "Es nulo",
                        "No es nulo",
                    ]
                else:
                    filter_criteria = [
                        "Contiene",
                        "No contiene",
                        "Empieza con",
                        "Termina con",
                        "Es nulo",
                        "No es nulo",
                    ]

                filter_criterion = st.selectbox(
                    f"Criterio", filter_criteria, key=f"crit_{i}"
                )

            with col3:
                filter_value = (
                    st.text_input(f"Valor", "", key=f"val_{i}")
                    if filter_criterion not in ["Es nulo", "No es nulo"]
                    else None
                )

            with col4:
                if i < num_filtros - 1:
                    criterio_seleccionado = st.radio(
                        "Criterio", ("AND", "OR"), key=f"radio_{i}"
                    )
                    criterios.append(criterio_seleccionado)

            filtro = None

            if filter_criterion in [
                "Mayor que",
                "Menor que",
                "Igual a",
                "Diferente de",
            ]:
                try:
                    filter_value = float(filter_value) if filter_value else None
                except ValueError:
                    st.error(
                        f"Por favor, ingresa un valor numérico válido para {column}."
                    )
                    filter_value = None

            if (filter_value is not None) or (
                filter_criterion in ["Es nulo", "No es nulo"]
            ):
                if df[column].dtype in ["int64", "float64"]:
                    if filter_criterion == "Mayor que":
                        filtro = df[column] > filter_value
                    elif filter_criterion == "Menor que":
                        filtro = df[column] < filter_value
                    elif filter_criterion == "Igual a":
                        filtro = df[column] == filter_value
                    elif filter_criterion == "Diferente de":
                        filtro = df[column] != filter_value
                    elif filter_criterion == "Es nulo":
                        filtro = df[column].isnull()
                    elif filter_criterion == "No es nulo":
                        filtro = df[column].notnull()
                else:
                    if filter_criterion == "Contiene":
                        filtro = df[column].str.contains(
                            filter_value, case=False, na=False
                        )
                    elif filter_criterion == "No contiene":
                        filtro = ~df[column].str.contains(
                            filter_value, case=False, na=False
                        )
                    elif filter_criterion == "Empieza con":
                        filtro = df[column].str.startswith(filter_value, na=False)
                    elif filter_criterion == "Termina con":
                        filtro = df[column].str.endswith(filter_value, na=False)
                    elif filter_criterion == "Es nulo":
                        filtro = df[column].isnull()
                    elif filter_criterion == "No es nulo":
                        filtro = df[column].notnull()

            if filtro is not None:
                filtros.append(filtro)

    if filtros:
        filtro_combinado = filtros[0]
        for i in range(1, len(filtros)):
            if criterios[i - 1] == "AND":
                filtro_combinado &= filtros[i]
            elif criterios[i - 1] == "OR":
                filtro_combinado |= filtros[i]

        filtered_df = df[filtro_combinado]
    else:
        filtered_df = df

    st.subheader("Tabla Filtrada")
    st.dataframe(filtered_df)

    # Input para el nombre del archivo
    output_file_name = st.text_input(
        "Ingresa el nombre del archivo de salida (sin extensión)", "tabla_filtrada"
    )

    # Guardar el archivo filtrado en memoria y crear un enlace de descarga
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    # Convertir el DataFrame a un archivo de Excel
    filtered_df_to_excel = to_excel(filtered_df)

    # Crear enlace de descarga
    st.download_button(
        label="Descargar tabla filtrada como Excel",
        data=filtered_df_to_excel,
        file_name=f"{output_file_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )