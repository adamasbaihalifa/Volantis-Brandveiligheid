import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import re
from xlsxwriter.utility import xl_rowcol_to_cell
import matplotlib.pyplot as plt
import seaborn as sns

# Custom colors
custom_colors = ['#D0DFE6', '#FBCDAB', '#EC6907', '#A6CEE3', '#B2DF8A', 
                 '#FDBF6F', '#CAB2D6', '#FF7F00', '#FB9A99']

def add_logo(image_path):
    """Voeg een logo toe aan de Streamlit app."""
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
    """Transformeert de Excel data voor verwerking."""
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
    """Maakt de Tolerantie-Risico matrix."""
    kans_labels = {
        1: "1 - zeer onwaarschijnlijk",
        2: "2 - onwaarschijnlijk",
        3: "3 - mogelijk",
        4: "4 - waarschijnlijk",
        5: "5 - zeer waarschijnlijk"
    }
    
    effect_labels = {
        1: "1 - klein",
        2: "2 - matig",
        3: "3 - hevig",
        4: "4 - ernstig",
        5: "5 - rampzalig"
    }
    
    matrix = pd.DataFrame(
        0, 
        index=pd.Categorical(list(kans_labels.values()), categories=list(kans_labels.values())),
        columns=pd.Categorical(list(effect_labels.values()), categories=list(effect_labels.values()))
    )
    
    def extract_number(value):
        match = re.search(r'\d+', str(value))
        return int(match.group()) if match else 0
    
    for _, row in df.iterrows():
        kans = extract_number(row['Kans'])
        effect = extract_number(row['Effect'])
        if kans in kans_labels and effect in effect_labels:
            matrix.loc[kans_labels[kans], effect_labels[effect]] += 1
    
    return matrix

def add_matrix_sheet(writer, matrix):
    """Voeg matrix worksheet toe met opmaak."""
    workbook = writer.book
    worksheet = workbook.add_worksheet('Tolerantie Matrix')
    
    # Header
    worksheet.write(0, 0, "Tolerantie - Risico Matrix")
    for col_idx, value in enumerate(matrix.columns, start=1):
        worksheet.write(0, col_idx, value)
    
    # Data
    for row_idx, (index, row) in enumerate(matrix.iterrows(), start=1):
        worksheet.write(row_idx, 0, index)
        for col_idx, value in enumerate(row, start=1):
            worksheet.write(row_idx, col_idx, value)
    
    # Opmaak
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

def add_barchart_sheet(writer, df):
    """Voeg risico bar chart toe."""
    workbook = writer.book
    worksheet = workbook.add_worksheet('Top Risicos')
    
    # Filter en sorteer
    filtered = df[df['Risico'] > 0].sort_values('Risico', ascending=False)
    
    # Schrijf data
    filtered[['kenmerken', 'Risico']].to_excel(
        writer, 
        sheet_name='Top Risicos', 
        startrow=1, 
        index=False
    )
    
    # Maak chart
    chart = workbook.add_chart({'type': 'bar'})
    chart.add_series({
        'categories': ['Top Risicos', 1, 0, len(filtered), 0],
        'values':     ['Top Risicos', 1, 1, len(filtered), 1],
        'fill':       {'color': '#EC6907'},
        'name':       'Risico Score'
    })
    
    # Chart instellingen
    chart.set_title({'name': 'Meest Risicovolle Kenmerken'})
    chart.set_x_axis({'name': 'Kenmerk', 'label_position': 'low'})
    chart.set_y_axis({'name': 'Risico Score'})
    chart.set_legend({'position': 'none'})
    chart.set_style(11)
    
    worksheet.insert_chart('D2', chart)

def style_excel(writer, df):
    """Opmaak voor de hoofdsheet."""
    workbook = writer.book
    worksheet = writer.sheets['Resultaat']
    colors = iter(custom_colors)
    border_format = workbook.add_format({'border': 1})
    current_color = next(colors)
    previous_field = None
    
    for row_idx, field_value in enumerate(df['kenmerken'], start=1):
        if field_value != previous_field:
            current_color = next(colors, custom_colors[0])
            previous_field = field_value
        
        for col_idx in range(df.shape[1]):
            cell_value = df.iloc[row_idx - 1, col_idx]
            if pd.isna(cell_value):
                cell_value = ""
            cell_format = workbook.add_format({
                'bg_color': current_color, 
                'border': 1, 
                'font_color': '#000000'
            })
            worksheet.write(row_idx, col_idx, cell_value, cell_format)
    
    worksheet.conditional_format(
        0, 0, df.shape[0], df.shape[1] - 1, 
        {'type': 'no_blanks', 'format': border_format}
    )

def main():
    add_logo("bcc_volantis-logo_cmyk.jpg")
    st.title("Volantis: Excel Transformatie App")
    uploaded_file = st.file_uploader("Kies een Excel-bestand", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            # Sectie 1: Raw data
            st.subheader("1. Originele Data")
            raw_df = pd.read_excel(uploaded_file, header=None)
            with st.expander("Toon volledige dataset"):
                st.dataframe(raw_df)
            
            # Sectie 2: Dataverwerking
            st.subheader("2. Dataverwerking")
            start_row = raw_df[
                raw_df.apply(lambda row: row.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(), axis=1)
            ].index
            
            if not start_row.empty:
                processed_df = raw_df.iloc[start_row[0]:].copy()
                with st.expander("Toon verwerkte input"):
                    st.dataframe(processed_df)
            else:
                st.error("Startrij niet gevonden")
                return
                
            # Sectie 3: Getransformeerde data
            st.subheader("3. Resultaat")
            transformed_df = transform_data(processed_df)
            st.dataframe(transformed_df)
            
            # Sectie 4: Visualisaties
            st.subheader("4. Risico Analyse")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Risico matrix
                risk_matrix = create_risk_matrix(transformed_df)
                st.write("**Tolerantie Matrix**")
                fig, ax = plt.subplots(figsize=(8, 5))
                sns.heatmap(
                    risk_matrix.astype(int), 
                    annot=True, 
                    fmt="d", 
                    cmap="YlOrRd", 
                    linewidths=.5,
                    ax=ax
                )
                ax.set_xlabel("Effect")
                ax.set_ylabel("Kans")
                st.pyplot(fig)
            
            with col2:
                # Bar chart
                st.write("**Risico Verdeling**")
                filtered = transformed_df[transformed_df['Risico'] > 0]
                if not filtered.empty:
                    fig2, ax2 = plt.subplots(figsize=(8, 5))
                    sns.barplot(
                        data=filtered.sort_values('Risico', ascending=False),
                        y='kenmerken',
                        x='Risico',
                        palette="YlOrRd",
                        ax=ax2
                    )
                    ax2.set_xlabel("Risico Score")
                    ax2.set_ylabel("")
                    plt.tight_layout()
                    st.pyplot(fig2)
                else:
                    st.info("Geen risico's gevonden")
            
            # Sectie 5: Download
            st.subheader("5. Download Rapport")
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="Resultaat")
                style_excel(writer, transformed_df)
                add_matrix_sheet(writer, risk_matrix)
                add_barchart_sheet(writer, transformed_df)
            
            st.download_button(
                label="📥 Download Excel Rapport",
                data=output.getvalue(),
                file_name="risico_analyse.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Klik om het volledige rapport te downloaden"
            )
            
        except Exception as e:
            st.error(f"Fout opgetreden: {str(e)}")

if __name__ == "__main__":
    main()