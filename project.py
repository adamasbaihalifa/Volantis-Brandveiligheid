import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# Updated custom colors (removed purple)
custom_colors = ['#D0DFE6', '#FBCDAB', '#EC6907', '#A6CEE3', '#B2DF8A', '#FDBF6F', '#CAB2D6', '#FF7F00', '#FB9A99']

def add_logo(image_path):
    """Add a logo to the Streamlit app."""
    with open(image_path, "rb") as file:
        contents = file.read()
        encoded_image = base64.b64encode(contents).decode()
    st.markdown(
        f"""
        <div style="text-align: center;">
            <img src="data:image/png;base64,{encoded_image}" width="300">
        </div>
        """,
        unsafe_allow_html=True,
    )

def transform_data(input_data):
    """Transforms the Excel data for processing."""
    # Assume the first row of the Excel should become column headers.
    input_data.columns = input_data.iloc[0]
    input_data = input_data.drop(input_data.index[0])
    
    # Transpose
    transposed = input_data.T
    
    # Reset index so the old row labels become a normal column
    transposed.reset_index(inplace=True)
    
    # Rename columns
    col_count = transposed.shape[1]  # Number of columns
    new_col_names = ['Field'] + [f'Value {i}' for i in range(1, col_count)]
    transposed.columns = new_col_names
    
    return transposed

def style_excel(writer, df):
    """Applies alternating colors and grids in Excel."""
    workbook = writer.book
    worksheet = writer.sheets['Resultaat']
    colors = iter(custom_colors)
    
    # Define formats
    border_format = workbook.add_format({'border': 1})
    current_color = next(colors)
    previous_field = None

    for row_idx, field_value in enumerate(df['Field'], start=1):  # Skip header row
        if field_value != previous_field:
            current_color = next(colors, custom_colors[0])
            previous_field = field_value
        
        # Apply color and grid formatting up to column D
        for col_idx in range(min(4, df.shape[1])):  # Stop at column D (0-indexed: 0, 1, 2, 3)
            cell_value = df.iloc[row_idx - 1, col_idx]
            if pd.isna(cell_value):  # Handle NaN values
                cell_value = ""
            cell_format = workbook.add_format({'bg_color': current_color, 'border': 1, 'font_color': '#000000'})
            worksheet.write(row_idx, col_idx, cell_value, cell_format)

    # Apply grid for all cells
    worksheet.conditional_format(0, 0, df.shape[0], df.shape[1] - 1, {'type': 'no_blanks', 'format': border_format})

def main():
    # Add a logo
    add_logo("bcc_volantis-logo_cmyk.jpg")
    
    # App title
    st.title("Volantis: Excel Transformatie App")
    uploaded_file = st.file_uploader("Kies een Excel-bestand", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            # Read the Excel file
            input_data = pd.read_excel(uploaded_file, header=None)

            st.write("**Originele data (voorvertoning)**")
            st.dataframe(input_data)

            # Find the starting row containing "Omgevingskenmerken - Algemeen"
            start_row = input_data[input_data.apply(lambda row: row.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(), axis=1)].index
            if not start_row.empty:
                start_row_index = start_row[0]
                input_data = input_data.iloc[start_row_index:]
            else:
                st.error("De rij 'Omgevingskenmerken - Algemeen' kon niet worden gevonden.")
                return

            # Transform the data
            transformed_df = transform_data(input_data)

            st.write("**Getransformeerde data (voorvertoning)**")
            st.dataframe(transformed_df)

            # Generate styled Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="Resultaat")
                style_excel(writer, transformed_df)  # Apply styling

            # Download button
            st.download_button(
                label="Download getransformeerde Excel",
                data=output.getvalue(),
                file_name="getransformeerd.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Fout bij het verwerken van het Excel-bestand: {e}")

if __name__ == "__main__":
    main()
