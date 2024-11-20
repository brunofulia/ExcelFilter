import streamlit as st
import pandas as pd
from io import BytesIO


# --- Funciones auxiliares ---
def cargar_archivo(uploaded_file):
    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return None


def generar_filtro(df, column, criterion, value):
    try:
        match criterion:
            case "Es nulo":
                return df[column].isnull()
            case "No es nulo":
                return df[column].notnull()
            case "Mayor que" | "Menor que" | "Igual a" | "Diferente de" if df[
                column
            ].dtype in ["int64", "float64"]:
                value = float(value)
                match criterion:
                    case "Mayor que":
                        return df[column] > value
                    case "Menor que":
                        return df[column] < value
                    case "Igual a":
                        return df[column] == value
                    case "Diferente de":
                        return df[column] != value
            case "Contiene" | "No contiene" | "Empieza con" | "Termina con" if df[
                column
            ].dtype == "object":
                match criterion:
                    case "Contiene":
                        return df[column].str.contains(value, case=False, na=False)
                    case "No contiene":
                        return ~df[column].str.contains(value, case=False, na=False)
                    case "Empieza con":
                        return df[column].str.startswith(value, na=False)
                    case "Termina con":
                        return df[column].str.endswith(value, na=False)
            case _:
                st.error(
                    f"Criterio '{criterion}' no válido para la columna seleccionada."
                )
                return None
    except ValueError:
        st.error(f"El valor ingresado no es válido para el criterio '{criterion}'.")
        return None
    except Exception as e:
        st.error(f"Error al aplicar el filtro: {e}")
        return None


def aplicar_filtros(df, filtros, criterios):
    if not filtros:
        return df
    filtro_combinado = filtros[0]
    for i in range(1, len(filtros)):
        match criterios[i - 1]:
            case "AND":
                filtro_combinado &= filtros[i]
            case "OR":
                filtro_combinado |= filtros[i]
    return df[filtro_combinado]


def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, max(len(col) + 2, 12))
    return output.getvalue()


# --- Interfaz de usuario ---
st.title("Filtrar y Guardar Tabla de Excel")

uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])

if uploaded_file:
    df = cargar_archivo(uploaded_file)
    if df is not None:
        # Mostrar tabla original con número de registros
        st.subheader("Tabla Completa")
        st.write(f"Número de registros: {len(df)}")
        st.dataframe(df)

        # Inicializar el estado de los filtros si no existe
        if "filtros" not in st.session_state:
            st.session_state.filtros = []
            st.session_state.criterios = []
            st.session_state.num_filtros = 1

        # Botón para resetear filtros
        if st.button("Resetear Filtros"):
            st.session_state.filtros = []
            st.session_state.criterios = []
            st.session_state.num_filtros = 1
            st.session_state.apply_filters = True
            st.rerun()

        # Configuración de filtros
        with st.expander("Agregar Filtros"):
            st.session_state.num_filtros = st.number_input(
                "Número de filtros",
                min_value=1,
                step=1,
                value=st.session_state.num_filtros,
            )

            for i in range(st.session_state.num_filtros):
                st.write(f"Filtro {i + 1}")
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    column = st.selectbox(f"Columna", df.columns, key=f"col_{i}")

                with col2:
                    filter_criteria = (
                        [
                            "Mayor que",
                            "Menor que",
                            "Igual a",
                            "Diferente de",
                            "Es nulo",
                            "No es nulo",
                        ]
                        if df[column].dtype in ["int64", "float64"]
                        else [
                            "Contiene",
                            "No contiene",
                            "Empieza con",
                            "Termina con",
                            "Es nulo",
                            "No es nulo",
                        ]
                    )
                    criterion = st.selectbox(
                        f"Criterio", filter_criteria, key=f"crit_{i}"
                    )

                with col3:
                    value = (
                        st.text_input(
                            f"Valor",
                            key=f"val_{i}",
                            on_change=lambda: st.session_state.update(
                                {"apply_filters": True}
                            ),
                        )
                        if criterion not in ["Es nulo", "No es nulo"]
                        else None
                    )

                filtro = generar_filtro(df, column, criterion, value)
                if filtro is not None:
                    if i < len(st.session_state.filtros):
                        st.session_state.filtros[i] = filtro
                    else:
                        st.session_state.filtros.append(filtro)

                with col4:
                    if i < st.session_state.num_filtros - 1:
                        if i < len(st.session_state.criterios):
                            st.session_state.criterios[i] = st.radio(
                                "Criterio", ["AND", "OR"], key=f"crit_radio_{i}"
                            )
                        else:
                            st.session_state.criterios.append(
                                st.radio(
                                    "Criterio", ["AND", "OR"], key=f"crit_radio_{i}"
                                )
                            )

        # Aplicar filtros automáticamente
        if "apply_filters" not in st.session_state:
            st.session_state.apply_filters = False

        if st.session_state.apply_filters or (
            len(st.session_state.filtros) > 0
            and any(f is not None for f in st.session_state.filtros)
        ):
            filtered_df = aplicar_filtros(
                df, st.session_state.filtros, st.session_state.criterios
            )
            st.subheader("Tabla Filtrada")
            st.write(f"Número de registros: {len(filtered_df)}")
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

        # Resetear el estado de apply_filters
        st.session_state.apply_filters = False
