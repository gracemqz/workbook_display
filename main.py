import streamlit as st
import pandas as pd
from io import BytesIO


# Custom CSS for styling
def load_excel_to_session(uploaded_file):
    """
    Reads an Excel file from uploaded_file into st.session_state["workbook"].
    """
    excel = pd.ExcelFile(uploaded_file)
    st.session_state["sheet_names"] = excel.sheet_names
    st.session_state["workbook"] = {}

    for sheet_name in excel.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        st.session_state["workbook"][sheet_name] = df


def load_csv_to_session(uploaded_file):
    """
    Reads a CSV file from uploaded_file into st.session_state["workbook"].
    """
    df = pd.read_csv(uploaded_file)
    st.session_state["sheet_names"] = ["Sheet1"]
    st.session_state["workbook"] = {"Sheet1": df}


def download_excel(workbook_dict):
    """
    Converts the workbook_dict (sheet_name -> DataFrame) into an in-memory Excel file
    and returns the buffer.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in workbook_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output


def download_csv(df):
    """
    Converts the single DataFrame to CSV and returns as bytes (or string).
    """
    return df.to_csv(index=False).encode("utf-8")


# ------------------ Streamlit App -------------------
def main():
    st.markdown(
        "<h1 style='color: #006BB8;'>Workbook Editor</h1>", unsafe_allow_html=True
    )

    # 1. File upload
    st.header("1. Workbook Upload")
    uploaded_file = st.file_uploader("", type=["xlsx", "xls", "csv"])

    # Initialize session_state if not already
    if "workbook" not in st.session_state:
        st.session_state["workbook"] = {}
    if "sheet_names" not in st.session_state:
        st.session_state["sheet_names"] = []

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split(".")[-1].lower()

        # 2. Load data into session state if not already loaded or if new upload
        if st.button("Load Workbook"):
            if file_extension in ["xlsx", "xls"]:
                load_excel_to_session(uploaded_file)
            elif file_extension == "csv":
                load_csv_to_session(uploaded_file)
            st.success("Workbook loaded successfully")

    # Only show the rest of the UI if workbook data is loaded
    if st.session_state["workbook"]:
        # 3. Sheet selection (if multiple sheets)
        st.header("2. Sheet Selection and Editing")
        if len(st.session_state["sheet_names"]) > 1:
            selected_sheet = st.selectbox("", st.session_state["sheet_names"])
        else:
            selected_sheet = st.session_state["sheet_names"][0]

        # 4. Display and Edit
        st.write(f"Sheet: {selected_sheet}")
        edited_df = st.data_editor(st.session_state["workbook"][selected_sheet])

        # Every time there's an edit, st.data_editor returns the updated DataFrame
        st.session_state["workbook"][selected_sheet] = edited_df

        # 5. Allow download of updated file
        file_extension = (
            uploaded_file.name.split(".")[-1].lower() if uploaded_file else None
        )

        if file_extension in ["xlsx", "xls"]:
            # Convert to Excel
            excel_data = download_excel(st.session_state["workbook"])
            st.download_button(
                label="Download Workbook",
                data=excel_data,
                file_name="updated_workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        elif file_extension == "csv":
            # Convert to CSV
            csv_data = download_csv(st.session_state["workbook"]["Sheet1"])
            st.download_button(
                label="Download Updated CSV",
                data=csv_data,
                file_name="updated_data.csv",
                mime="text/csv",
            )


if __name__ == "__main__":
    main()
