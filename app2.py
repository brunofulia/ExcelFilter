import streamlit as st
import pandas as pd
from io import BytesIO

# --- Funciones auxiliares ---
def cargar_archivo(uploaded_file):
    """
    Carga un archivo Excel desde la subida del usuario.
    Args:
        uploaded_file: archivo subido por el usuario.
    Returns:
        DataFrame cargado con pandas.
    """
    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return None


def generar_filtro(df, column, criterion, value):
    """
    Genera un filtro basado en la columna, el criterio y el valor.
    Args:
        df: DataFrame sobre el que se aplica el filtro.
        column: Nombre de la columna.
        criterion: Criterio de filtro seleccionado.
        value: Valor del filtro.
    Returns:
        Filtro pandas compatible o None si hay un error.
    """
    if criterion in ["Es nulo", "No es nulo"]:
        return df[column].isnull() if criterion == "Es nulo" else df[column].notnull()
    try:
        if df[column].dtype in ["int64", "float64"]:
            value = float(value)
            if criterion == "Mayor que":
                return df[column] > value
            elif criterion == "Menor que":
                return df[column] < value
            elif criterion == "Igual a":
                return df[column] == value
            elif criterion == "Diferente de":
                return df[column] != value
        else:
            if criterion == "Contiene":
                return df[column].str.contains(value, case=False, na=False)
            elif criterion == "No contiene":
                return ~df[column].str.contains(value, case=False, na=False)
            elif criterion == "Empieza con":
                return df[column].str.startswith(value, na=False)
            elif criterion == "Termina con":
                return df[column].str.endswith(value, na=False)
    except Exception as e:
        st.error(f"Error al aplicar el filtro: {e}")
    return None


def aplicar_filtros(df, filtros, criterios):
    """
    Combina los filtros aplicados al DataFrame según los criterios.
    Args:
        df: DataFrame original.
        filtros: Lista de filtros pandas.
        criterios: Lista de criterios ("AND", "OR").
    Returns:
        DataFrame filtrado.
    """
    if not filtros:
        return df

    filtro_combinado = filtros[0]
    for i in range(1, len(filtros)):
        if criterios[i - 1] == "AND":
            filtro_combinado &= filtros[i]
        elif criterios[i - 1] == "OR":
            filtro_combinado |= filtros[i]

    return df[filtro_combinado]


def exportar_excel(df):
    """
    Exporta un DataFrame a un archivo Excel en memoria.
    Args:
        df: DataFrame a exportar.
    Returns:
        BytesIO con los datos del archivo Excel.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, max(len(col) + 2, 12))  # Ajusta el tamaño
    return output.getvalue()


# --- Interfaz de usuario ---
st.title("Filtrar y Guardar Tabla de Excel")

# Carga del archivo
uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])

if uploaded_file:
    df = cargar_archivo(uploaded_file)
    if df is not None:
        # Mostrar tabla original
        st.subheader("Tabla Completa")
        st.dataframe(df)

        # Configuración de filtros
        with st.expander("Agregar Filtros"):
            filtros = []
            criterios = []

            num_filtros = st.number_input(
                "Número de filtros", min_value=1, step=1, value=1
            )
            
            #""""""""""""""""""""""""""""""""""
            for i in range(num_filtros):
                st.write(f"Filtro {i + 1}")
                col1, col2, col3 = st.columns(3)

                with col1:
                    column = st.selectbox(f"Columna", df.columns, key=f"col_{i}")

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
                    criterion = st.selectbox(
                        f"Criterio", filter_criteria, key=f"crit_{i}"
                    )

                with col3:
                    value = (
                        st.text_input(f"Valor", key=f"val_{i}")
                        if criterion not in ["Es nulo", "No es nulo"]
                        else None
                    )

                filtro = generar_filtro(df, column, criterion, value)
                if filtro is not None:
                    filtros.append(filtro)

                if i < num_filtros - 1:
                    criterios.append(st.radio("Criterio", ["AND", "OR"], key=f"crit_radio_{i}"))


            #""""""""""""""""""""""""""""""""""

        # Aplicar filtros
        if st.button("Aplicar Filtros"):
            filtered_df = aplicar_filtros(df, filtros, criterios)
            st.subheader("Tabla Filtrada")
            st.dataframe(filtered_df)

            # Exportar tabla filtrada
            output_file_name = st.text_input(
                "Ingresa el nombre del archivo de salida (sin extensión)",
                "tabla_filtrada",
            )
            filtered_df_to_excel = exportar_excel(filtered_df)

            st.download_button(
                label="Descargar tabla filtrada como Excel",
                data=filtered_df_to_excel,
                file_name=f"{output_file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
