import streamlit as st
import pandas as pd
from io import BytesIO


# --- Helper Functions ---
def load_excel_file(uploaded_file):
    try:
        # Load all sheets from the Excel file
        return pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading the file: {e}")
        return None


def generate_filter(df, column, criterion, value):
    try:
        match criterion:
            case "Is null":
                return df[column].isnull()
            case "Is not null":
                return df[column].notnull()
            case "Greater than" | "Less than" | "Equal to" | "Not equal to" if df[
                column
            ].dtype in ["int64", "float64"]:
                value = float(value)
                match criterion:
                    case "Greater than":
                        return df[column] > value
                    case "Less than":
                        return df[column] < value
                    case "Equal to":
                        return df[column] == value
                    case "Not equal to":
                        return df[column] != value
            case "Contains" | "Does not contain" | "Starts with" | "Ends with" if df[
                column
            ].dtype == "object":
                match criterion:
                    case "Contains":
                        return df[column].str.contains(value, case=False, na=False)
                    case "Does not contain":
                        return ~df[column].str.contains(value, case=False, na=False)
                    case "Starts with":
                        return df[column].str.startswith(value, na=False)
                    case "Ends with":
                        return df[column].str.endswith(value, na=False)
            case _:
                st.error(f"Criterion '{criterion}' is not valid for the selected column.")
                return None
    except ValueError:
        st.error(f"The entered value is not valid for the criterion '{criterion}'.")
        return None
    except Exception as e:
        st.error(f"Error applying the filter: {e}")
        return None


def apply_filters(df, filters, conditions):
    if not filters:
        return df
    combined_filter = filters[0]
    for i in range(1, len(filters)):
        match conditions[i - 1]:
            case "AND":
                combined_filter &= filters[i]
            case "OR":
                combined_filter |= filters[i]
    return df[combined_filter]


def export_to_excel(dfs, selected_sheet):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        dfs[selected_sheet].to_excel(writer, index=False, sheet_name="FilteredData")
        worksheet = writer.sheets["FilteredData"]
        for i, col in enumerate(dfs[selected_sheet].columns):
            worksheet.set_column(i, i, max(len(col) + 2, 12))
    return output.getvalue()


# --- User Interface ---
st.title("Filter and Save Excel Workbook")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    sheets = load_excel_file(uploaded_file)
    if sheets is not None:
        # Let the user select the sheet
        sheet_names = list(sheets.keys())
        selected_sheet = st.selectbox("Select a sheet to work with", sheet_names)

        # Load the selected sheet
        df = sheets[selected_sheet]

        # Display the original table with record count
        st.subheader("Full Table")
        st.write(f"Number of records: {len(df)}")
        st.dataframe(df)

        # Initialize filter states if not already present
        if "filters" not in st.session_state:
            st.session_state.filters = []
            st.session_state.conditions = []
            st.session_state.num_filters = 1

        # Button to reset filters
        if st.button("Reset Filters"):
            st.session_state.filters = []
            st.session_state.conditions = []
            st.session_state.num_filters = 1
            st.session_state.apply_filters = True
            st.rerun()

        # Filter configuration
        with st.expander("Add Filters"):
            st.session_state.num_filters = st.number_input(
                "Number of filters",
                min_value=1,
                step=1,
                value=st.session_state.num_filters,
            )

            for i in range(st.session_state.num_filters):
                st.write(f"Filter {i + 1}")
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    column = st.selectbox(f"Column", df.columns, key=f"col_{i}")

                with col2:
                    filter_criteria = (
                        [
                            "Greater than",
                            "Less than",
                            "Equal to",
                            "Not equal to",
                            "Is null",
                            "Is not null",
                        ]
                        if df[column].dtype in ["int64", "float64"]
                        else [
                            "Contains",
                            "Does not contain",
                            "Starts with",
                            "Ends with",
                            "Is null",
                            "Is not null",
                        ]
                    )
                    criterion = st.selectbox(
                        f"Criterion", filter_criteria, key=f"crit_{i}"
                    )

                with col3:
                    value = (
                        st.text_input(
                            f"Value",
                            key=f"val_{i}",
                            on_change=lambda: st.session_state.update(
                                {"apply_filters": True}
                            ),
                        )
                        if criterion not in ["Is null", "Is not null"]
                        else None
                    )

                filter_obj = generate_filter(df, column, criterion, value)
                if filter_obj is not None:
                    if i < len(st.session_state.filters):
                        st.session_state.filters[i] = filter_obj
                    else:
                        st.session_state.filters.append(filter_obj)

                with col4:
                    if i < st.session_state.num_filters - 1:
                        if i < len(st.session_state.conditions):
                            st.session_state.conditions[i] = st.radio(
                                "Condition", ["AND", "OR"], key=f"cond_radio_{i}"
                            )
                        else:
                            st.session_state.conditions.append(
                                st.radio(
                                    "Condition", ["AND", "OR"], key=f"cond_radio_{i}"
                                )
                            )

        # Automatically apply filters
        if "apply_filters" not in st.session_state:
            st.session_state.apply_filters = False

        if st.session_state.apply_filters or (
            len(st.session_state.filters) > 0
            and any(f is not None for f in st.session_state.filters)
        ):
            filtered_df = apply_filters(
                df, st.session_state.filters, st.session_state.conditions
            )
            st.subheader("Filtered Table")
            st.write(f"Number of records: {len(filtered_df)}")
            st.dataframe(filtered_df)

            # Export filtered table
            output_file_name = st.text_input(
                "Enter the output file name (without extension)",
                "filtered_table",
            )
            filtered_df_to_excel = export_to_excel(sheets, selected_sheet)

            st.download_button(
                label="Download filtered table as Excel",
                data=filtered_df_to_excel,
                file_name=f"{output_file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Reset apply_filters state
        st.session_state.apply_filters = False
