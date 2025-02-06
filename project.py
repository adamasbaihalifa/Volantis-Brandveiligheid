import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import re
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# Custom colors
custom_colors = ['#D0DFE6', '#FBCDAB', '#EC6907', '#A6CEE3', '#B2DF8A', 
                 '#FDBF6F', '#CAB2D6', '#FF7F00', '#FB9A99']

def add_logo(image_path):
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
    input_data.columns = input_data.iloc[0]
    input_data = input_data.drop(input_data.index[0])
    transposed = input_data.T.reset_index()
    
    transposed.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
    if transposed.shape[0] > 22:
        transposed = transposed.iloc[22:].reset_index(drop=True)
    
    processed_rows = []
    skip_next = False
    
    for idx in range(len(transposed)):
        if skip_next:
            skip_next = False
            continue
            
        current_row = transposed.iloc[idx].copy()
        kans_val = str(current_row["Kans"]).strip().lower()
        effect_val = str(current_row["Effect"]).strip().lower()
        
        if "inschatting kans" in kans_val:
            current_row["Kans"] = current_row["Effect"]
            
            if idx + 1 < len(transposed):
                next_row = transposed.iloc[idx + 1]
                if "inschatting effect" in str(next_row["Kans"]).lower():
                    current_row["Effect"] = next_row["Effect"]
                    skip_next = True
        
        elif "inschatting effect" in kans_val:
            continue
        
        def extract_number(value):
            match = re.search(r'\d+', str(value))
            return int(match.group()) if match else 0
        
        current_row["Risico"] = extract_number(current_row["Kans"]) * extract_number(current_row["Effect"])
        processed_rows.append(current_row)
    
    return pd.DataFrame(processed_rows)

def create_risk_matrix(df):
    matrix_labels = ['1', '2', '4', '5']
    matrix = pd.DataFrame(
        0, 
        index=pd.Categorical(matrix_labels, categories=matrix_labels),
        columns=pd.Categorical(matrix_labels, categories=matrix_labels)
    )
    
    def extract_number(value):
        match = re.search(r'\d+', str(value))
        return int(match.group()) if match else 0
    
    for _, row in df.iterrows():
        kans = extract_number(row['Kans'])
        effect = extract_number(row['Effect'])
        if kans in [1,2,4,5] and effect in [1,2,4,5]:
            matrix.loc[str(kans), str(effect)] += 1
    
    return matrix

def create_risk_categories(df):
    bins = [-1, 5, 6, 9, 25]
    labels = ['Safe (0-5)', 'Low (6)', 'Medium (7-9)', 'High (10-25)']
    df['Category'] = pd.cut(df['Risico'], bins=bins, labels=labels)
    return df['Category'].value_counts().reindex(labels, fill_value=0)

def add_matrix_sheet(writer, matrix):
    workbook = writer.book
    worksheet = workbook.add_worksheet('Risk Matrix')
    
    # Write header
    worksheet.write(0, 0, "Kans →\\Effect ↓")

    for col_idx, value in enumerate(matrix.columns, start=1):
        worksheet.write(0, col_idx, value)
    
    # Write index
    for row_idx, value in enumerate(matrix.index, start=1):
        worksheet.write(row_idx, 0, value)
    
    # Write data with product values
    for row_idx, row in enumerate(matrix.index, start=1):
        for col_idx, col in enumerate(matrix.columns, start=1):
            product = int(row) * int(col)
            count = matrix.loc[row, col]
            worksheet.write(row_idx, col_idx, f"{product} ({count})")
    
    # Conditional formatting
    max_val = matrix.max().max()
    for row in range(1, len(matrix)+1):
        for col in range(1, len(matrix.columns)+1):
            cell_ref = xl_rowcol_to_cell(row, col)
            worksheet.conditional_format(
                cell_ref, {
                    'type': '2_color_scale',
                    'min_value': 0,
                    'max_value': max_val,
                    'min_color': '#FFFFFF',
                    'max_color': '#EC6907'
                }
            )

def add_risk_categories_sheet(writer, categories):
    workbook = writer.book
    worksheet = workbook.add_worksheet('Risk Categories')
    
    # Write data
    worksheet.write(0, 0, 'Category')
    worksheet.write(0, 1, 'Count')
    
    for row_idx, (category, count) in enumerate(categories.items(), start=1):
        worksheet.write(row_idx, 0, category)
        worksheet.write(row_idx, 1, count)
    
    # Create pie chart
    chart = workbook.add_chart({'type': 'pie'})
    chart.add_series({
        'name': 'Risk Categories',
        'categories': ['Risk Categories', 1, 0, len(categories), 0],
        'values':     ['Risk Categories', 1, 1, len(categories), 1],
        'points': [
            {'fill': {'color': '#B2DF8A'}},  # Safe
            {'fill': {'color': '#FDBF6F'}},  # Low
            {'fill': {'color': '#EC6907'}},  # Medium
            {'fill': {'color': '#C00000'}},  # High
        ]
    })
    
    chart.set_title({'name': 'Risk Category Distribution'})
    worksheet.insert_chart('D2', chart)

def main():
    add_logo("bcc_volantis-logo_cmyk.jpg")
    st.title("Volantis: Advanced Risk Analysis")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            raw_df = pd.read_excel(uploaded_file, header=None)
            start_row = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(), axis=1)].index
            
            if not start_row.empty:
                processed_df = raw_df.iloc[start_row[0]:].copy()
                transformed_df = transform_data(processed_df)
                
                st.subheader("Processed Data")
                st.dataframe(transformed_df)
                
                st.subheader("Risk Analysis")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Risk Matrix**")
                    risk_matrix = create_risk_matrix(transformed_df)
                    fig, ax = plt.subplots(figsize=(8, 6))
                    sns.heatmap(
                        risk_matrix.astype(int), 
                        annot=True, 
                        fmt="d",
                        cmap="YlOrRd",
                        linewidths=.5,
                        linecolor='white',
                        ax=ax
                    )
                    ax.set_xlabel("Effect", fontsize=12)
                    ax.set_ylabel("Kans", fontsize=12)
                    st.pyplot(fig)
                
                with col2:
                    st.write("**Risk Categories**")
                    categories = create_risk_categories(transformed_df)
                    fig2, ax2 = plt.subplots(figsize=(8, 6))
                    categories.plot(
                        kind='barh', 
                        color=['#B2DF8A','#FDBF6F','#EC6907','#C00000'],
                        ax=ax2
                    )
                    ax2.set_xlabel("Count", fontsize=12)
                    ax2.set_ylabel("")
                    plt.xticks(rotation=45)
                    st.pyplot(fig2)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    transformed_df.to_excel(writer, index=False, sheet_name="Raw Data")
                    add_matrix_sheet(writer, risk_matrix)
                    add_risk_categories_sheet(writer, categories)
                
                st.download_button(
                    label="Download Full Report",
                    data=output.getvalue(),
                    file_name="risk_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            else:
                st.error("Starting row not found in the document")
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()