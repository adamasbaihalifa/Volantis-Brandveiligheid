import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from fpdf import FPDF
from io import BytesIO
import base64
import matplotlib.pyplot as plt
import seaborn as sns
import tempfile
from xlsxwriter.utility import xl_rowcol_to_cell
import math
import matplotlib.colors as mcolors
from matplotlib.colors import ListedColormap

################################################################################
# Streamlit Config en Kleuren
################################################################################
st.set_page_config(layout="wide", page_title="Risico Analyse 📊")

aangepaste_kleuren = [
    '#D0DFE6', '#FBCDAB',
    '#EC6907', '#A6CEE3',
    '#B2DF8A', '#FDBF6F',
    '#CAB2D6', '#FF7F00',
    '#FB9A99'
]

categorie_palet = {
    'Laag': '#A6CEE3',
    'Midden': '#BFD8D2',
    'Hoog': '#FFA54F',
    'Zeer hoog': '#E24E1B'
}

################################################################################
# Logo Tonen (Streamlit)
################################################################################
def voeg_logo_toe(afbeeldingspad):
    with open(afbeeldingspad, "rb") as bestand:
        inhoud = bestand.read()
        gecodeerde_afbeelding = base64.b64encode(inhoud).decode()
    st.markdown(
        f"""
        <div style="text-align: center;">
            <img src="data:image/png;base64,{gecodeerde_afbeelding}" width="600">
        </div>
        """,
        unsafe_allow_html=True,
    )

################################################################################
# Data Transformatie
################################################################################
def transformeer_gegevens(invoer_gegevens):
    """
    Transformeert en herstructureert de ingelezen DataFrame, zodat we
    de benodigde kolommen (Kenmerken, Kans, Effect) hebben.
    """
    # Eerste rij als kolomnamen
    invoer_gegevens.columns = invoer_gegevens.iloc[0]
    invoer_gegevens = invoer_gegevens.drop(invoer_gegevens.index[0])

    # Transponeren
    getransponeerd = invoer_gegevens.T.reset_index()
    getransponeerd.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
    # Als er meer dan 22 rijen zijn, knip vanaf rij 22 (afhankelijk van je Excel)
    if getransponeerd.shape[0] > 22:
        getransponeerd = getransponeerd.iloc[22:].reset_index(drop=True)

    # Vullen van lege 'kenmerken'
    getransponeerd['kenmerken'] = getransponeerd['kenmerken'].ffill()

    verwerkte_rijen = []
    sla_volgende_over = False

    for idx in range(len(getransponeerd)):
        if sla_volgende_over:
            sla_volgende_over = False
            continue
        
        huidige_rij = getransponeerd.iloc[idx].copy()
        kans_waarde = str(huidige_rij["Kans"]).strip().lower()
        effect_waarde = str(huidige_rij["Effect"]).strip().lower()

        if "inschatting kans" in kans_waarde:
            huidige_rij["Kans"] = huidige_rij["Effect"]
            if idx + 1 < len(getransponeerd):
                volgende_rij = getransponeerd.iloc[idx + 1]
                if "inschatting effect" in str(volgende_rij["Kans"]).lower():
                    huidige_rij["Effect"] = volgende_rij["Effect"]
                    sla_volgende_over = True
        elif "inschatting effect" in kans_waarde:
            continue
        
        def extraheer_nummer(waarde):
            match = re.search(r'\d+', str(waarde))
            return int(match.group()) if match else 0
        
        huidige_rij["Risico"] = extraheer_nummer(huidige_rij["Kans"]) * extraheer_nummer(huidige_rij["Effect"])
        verwerkte_rijen.append(huidige_rij)

    return pd.DataFrame(verwerkte_rijen)


################################################################################
# Hulpfunctie: extracteer int uit string
################################################################################
def extraheer_nummer(waarde):
    if isinstance(waarde, str):
        match = re.search(r'\((\d+)\)', waarde)
        if match:
            return int(match.group(1))
        else:
            match = re.search(r'\d+', waarde)
            return int(match.group()) if match else 0
    return 0

################################################################################
# Dynamische (optionele) risicomatrix
################################################################################
def maak_risicomatrix(df):
    # Aangepaste labels voor Effect (y-as) en Kans (x-as)
    y_labels = [
        'Zeer groot (5) - een dode, zeer grote schade', 
        'Groot (4) - een (zwaar) gewonde, grote schade',
        'Middel (3) - een (licht) gewonde, schade',
        'Klein (2) - geen slachtoffers, wel materiele schade', 
        'Beperkt (1) - beperkte materiele schade'
    ]
    
    x_labels = [
        'Zeer waarschijnlijk (1)',
        'Waarschijnlijk (2)',
        'Mogelijk (3)',
        'Onwaarschijnlijk (4)',
        'Zeer onwaarschijnlijk (5)'
    ]

    matrix = pd.DataFrame(
        0,
        index=pd.Categorical(y_labels, categories=y_labels[::-1], ordered=True),
        columns=pd.Categorical(x_labels, categories=x_labels, ordered=True)
    )
    
    for _, rij in df.iterrows():
        kans = extraheer_nummer(rij['Kans'])
        effect = extraheer_nummer(rij['Effect'])

        if 1 <= kans <= 5 and 1 <= effect <= 5:
            matrix.loc[y_labels[5-effect], x_labels[kans-1]] += 1
    
    return matrix

def maak_dynamische_matrix_afbeelding(matrix: pd.DataFrame) -> str:
    """
    Bouwt en slaat een risicomatrix op waarbij de celkleuren correct worden weergegeven in PDF.
    Gebruikt puur matplotlib in plaats van seaborn.
    """
    import matplotlib.pyplot as plt
    import matplotlib.colors as mcolors
    import numpy as np
    import tempfile

    def get_num_from_label(label):
        match = re.search(r"\((\d+)\)", str(label))
        return int(match.group(1)) if match else 0

    color_map = {
        (5,1): "green",  (5,2): "red",    (5,3): "red",    (5,4): "red",    (5,5): "red",
        (4,1): "green",  (4,2): "yellow", (4,3): "red",    (4,4): "red",    (4,5): "red",
        (3,1): "green",  (3,2): "yellow", (3,3): "yellow", (3,4): "red",    (3,5): "red",
        (2,1): "green",  (2,2): "green",  (2,3): "yellow", (2,4): "yellow", (2,5): "red",
        (1,1): "green",  (1,2): "green",  (1,3): "green",  (1,4): "green",  (1,5): "green"
    }

    n_rows, n_cols = matrix.shape
    colors = np.empty((n_rows, n_cols), dtype=object)

    for i, y_label in enumerate(matrix.index):
        effect_val = get_num_from_label(y_label)
        for j, x_label in enumerate(matrix.columns):
            chance_val = get_num_from_label(x_label)
            val = matrix.iloc[i, j]
            if val == 0:
                colors[i, j] = "#E0E0E0"
            else:
                colors[i, j] = color_map.get((effect_val, chance_val), "white")

    fig, ax = plt.subplots(figsize=(12, 8))

    for i in range(n_rows):
        for j in range(n_cols):
            rect = plt.Rectangle((j, i), 1, 1, facecolor=colors[i, j], edgecolor='black')
            ax.add_patch(rect)
            text = str(matrix.iloc[i, j])
            ax.text(j + 0.5, i + 0.5, text, ha='center', va='center', fontsize=14, fontweight='bold')

    ax.set_xlim(0, n_cols)
    ax.set_ylim(0, n_rows)
    ax.set_xticks(np.arange(n_cols) + 0.5)
    ax.set_yticks(np.arange(n_rows) + 0.5)
    ax.set_xticklabels(matrix.columns, rotation=45, ha='right', fontsize=12)
    ax.set_yticklabels(matrix.index, fontsize=12)
    ax.invert_yaxis()
    ax.set_xlabel("Kans", fontsize=14, fontweight='bold')
    ax.set_ylabel("Effect", fontsize=14, fontweight='bold')
    ax.set_aspect('equal')
    ax.tick_params(left=False, bottom=False)

    plt.tight_layout()

    tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    fig.savefig(tmpfile.name, dpi=300)
    plt.close(fig)

    return tmpfile.name

def color_cell(val):
    """
    Kleurentabel op basis van screenshot:
    <=4 => groen
    5-9 => geel
    >=10 => rood
    """
    if val <= 4:
        return "background-color: limegreen; color:black;"
    elif val <= 9:
        return "background-color: gold; color:black;"
    else:
        return "background-color: red; color:white;"


################################################################################
# Bepaal Tolerantie - Risico Analyse (tellingen per categorie)
################################################################################
def maak_risico_categorieen(df):
    """
    Verdeelt de 'Risico'-waarden in categorieën en geeft tellingen.
    """
    bins = [-1, 5, 10, 20, 30]  # grenzen voor categorieën
    labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
    df['Categorie'] = pd.cut(df['Risico'], bins=bins, labels=labels, right=False)
    return df['Categorie'].value_counts().reindex(labels, fill_value=0)


################################################################################
# PDF Generator
################################################################################
def pdf_tussenlijn(pdf, extra_space=4):
    """
    Maakt een subtiele lijn over de pagina, en voegt optioneel extra ruimte eronder toe.
    """
    pdf.set_line_width(0.2)
    pdf.set_draw_color(150, 150, 150)
    x_start = 10
    x_end = 200  # past bij A4 marges van FPDF
    y = pdf.get_y()
    pdf.line(x_start, y, x_end, y)
    pdf.ln(extra_space)

def pdf_kop(pdf, tekst, size=16, bold=True, align='L'):
    """ 
    Maakt een kop in de PDF met een bepaalde lettergrootte en dikte. 
    """
    if bold:
        pdf.set_font("Arial", 'B', size)
    else:
        pdf.set_font("Arial", '', size)
    pdf.cell(0, 10, tekst, ln=True, align=align)
    pdf.ln(2)

def pdf_subkop(pdf, tekst, size=12, bold=True, align='L'):
    """
    Maakt een subkop (iets kleiner dan kop) met horizontale lijn eronder.
    """
    if bold:
        pdf.set_font("Arial", 'B', size)
    else:
        pdf.set_font("Arial", '', size)
    pdf.cell(0, 8, tekst, ln=True, align=align)
    pdf_tussenlijn(pdf, extra_space=5)

def vervang_unicode_apostrof(tekst):
    return tekst.replace("’", "'")

def style_tolerantie_matrix(df):
    """
    Kleur de cellen precies zoals in de Tolerantiematrix-screenshot:
    """
    def get_num_from_label(label):
        match = re.search(r"\((\d+)\)", str(label))
        if match:
            return int(match.group(1))
        return None

    color_map = {
        (5,1): "green",  (5,2): "red",    (5,3): "red",    (5,4): "red",    (5,5): "red",
        (4,1): "green",  (4,2): "yellow", (4,3): "red",    (4,4): "red",    (4,5): "red",
        (3,1): "green",  (3,2): "yellow",(3,3): "yellow", (3,4): "red",    (3,5): "red",
        (2,1): "green",  (2,2): "green", (2,3): "yellow", (2,4): "yellow", (2,5): "red",
        (1,1): "green",  (1,2): "green", (1,3): "green",  (1,4): "green",  (1,5): "green",
    }

    styled = pd.DataFrame("", index=df.index, columns=df.columns)

    for row_i, row_label in enumerate(df.index):
        effect_val = get_num_from_label(row_label)  
        for col_j, col_label in enumerate(df.columns):
            chance_val = get_num_from_label(col_label) 
            count_in_cell = df.iloc[row_i, col_j]

            if count_in_cell == 0:
                styled.iloc[row_i, col_j] = "background-color: grey; color: black;"
            else:
                kleur = color_map.get((effect_val, chance_val), "green")
                if kleur == "red":
                    styled.iloc[row_i, col_j] = "background-color: red; color: white;"
                elif kleur == "yellow":
                    styled.iloc[row_i, col_j] = "background-color: yellow; color: black;"
                else:
                    styled.iloc[row_i, col_j] = "background-color: limegreen; color: black;"

    return styled


class CustomPDF(FPDF):
    pass

# def print_table_autofit(pdf, data, col_widths):
#     line_height = 6
#     for row_cells in data:
#         # (1) Bepaal aantal regels op basis van tekstbreedte
#         line_counts = []
#         for i, text in enumerate(row_cells):
#             txt = str(text)
#             col_w = col_widths[i]
#             txt_width = pdf.get_string_width(txt)
#             usable_width = max(col_w - 4, 1)
#             lines_needed = math.ceil(txt_width / usable_width)
#             line_counts.append(lines_needed if lines_needed > 0 else 1)

#         max_lines = max(line_counts)
#         row_height = max_lines * line_height

#         x_start = pdf.get_x()
#         y_start = pdf.get_y()

#         # (2) Print elke kolom
#         for i, text in enumerate(row_cells):
#             cell_x = pdf.get_x()
#             cell_y = pdf.get_y()
#             pdf.multi_cell(col_widths[i], line_height, str(text), border=1, align='L')
#             pdf.set_xy(cell_x + col_widths[i], cell_y)

#         # (3) Naar de volgende rij
#         pdf.set_xy(x_start, y_start + row_height)


def print_table_autofit(pdf, data, col_widths):
    line_height = 5  # Iets compacter
    font_size = 10
    pdf.set_font("Arial", size=font_size)
    
    # Zet kop in vet
    header = data[0]
    pdf.set_font("Arial", 'B', font_size)
    for i, header_cell in enumerate(header):
        pdf.multi_cell(col_widths[i], line_height, str(header_cell), border=1, align='C', ln=3, max_line_height=pdf.font_size)
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.set_xy(x + col_widths[i], y - line_height)
    pdf.ln(line_height * 2)

    pdf.set_font("Arial", '', font_size)

    for row in data[1:]:
        # Bepaal het maximale aantal regels nodig in deze rij
        line_counts = []
        for i, cell in enumerate(row):
            text_width = pdf.get_string_width(str(cell))
            usable_width = col_widths[i] - 2  # kleine marge
            wrapped_lines = math.ceil(text_width / usable_width)
            line_counts.append(max(1, wrapped_lines))

        max_lines = max(line_counts)
        row_height = line_height * max_lines

        x_start = pdf.get_x()
        y_start = pdf.get_y()

        for i, cell in enumerate(row):
            x = pdf.get_x()
            y = pdf.get_y()
            pdf.multi_cell(col_widths[i], line_height, str(cell), border=1, align='L')
            pdf.set_xy(x + col_widths[i], y)
        pdf.set_xy(x_start, y_start + row_height)


def genereer_pdf(
    top_10_risico, 
    tolerantie_df,
    risicokenmerken_path,
    top10_kenmerken_path,
    risicomatrix
):
    # Hieronder kun je de oriëntering op Landscape zetten als je wilt:
    # pdf = CustomPDF('L', 'mm', 'A4')
    pdf = CustomPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # PAGINA 1
    pdf.add_page()
    
    # -- LOGO in rechterbovenhoek --
    # Pas x en y eventueel iets aan zodat het logo mooier op de pagina staat
    try:
        pdf.image("bcc_volantis-logo_cmyk.jpg", x=165, y=10, w=30)
    except:
        pass  # als het pad niet klopt, slaan we het logo over

    # Kop + kleine inleiding
    pdf_kop(pdf, "Volantis: RGB Tool Risicoanalyse", size=18)
    pdf.set_font("Arial", '', 11)
    pdf.multi_cell(
        0, 6, vervang_unicode_apostrof(
            "Inleiding:\n\n"
            "Deze tool is ontwikkeld om snel en eenvoudig de belangrijkste risico's "
            "van een bouwwerk te identificeren aan de hand van Kans en Effect. "
            "Hieronder ziet u de Dynamische Tolerantiematrix, gevolgd door "
            "meer details over de belangrijkste risico's."
        )
    )
    pdf.ln(5)

    pdf_subkop(pdf, "Dynamische Tolerantiematrix", size=14)
    static_matrix_img = maak_dynamische_matrix_afbeelding(risicomatrix)
    # Vergroot de width (w=190) zodat hij bijna de hele A4-breedte vult
    pdf.image(static_matrix_img, x=10, y=pdf.get_y(), w=190)
    pdf.ln(150)  # Extra ruimte na de matrix
    pdf_subkop(pdf, "Tolerantie - Risico Analyse", size=14)
    pdf.set_font("Arial", '', 10)
    col_widths = [40, 40, 40]
    headers = tolerantie_df.columns.tolist()
    pdf.set_fill_color(200, 220, 255)
    # Tabelkop
    for i, head in enumerate(headers):
        pdf.cell(col_widths[i], 7, head, border=1, align='C', fill=True)
    pdf.ln()
    # Tabel-rijen
    for _, row in tolerantie_df.iterrows():
        pdf.cell(col_widths[0], 7, str(row[headers[0]]), border=1, align='C')
        pdf.cell(col_widths[1], 7, str(row[headers[1]]), border=1, align='C')
        pdf.cell(col_widths[2], 7, str(row[headers[2]]), border=1, align='C')
        pdf.ln()
    # PAGINA 2
    pdf.add_page()
    pdf_subkop(pdf, "Risicokenmerken (Som)", size=14)
    pdf.image(risicokenmerken_path, x=15, y=pdf.get_y(), w=180)
    # PAGINA 3
    pdf.add_page()
    pdf_subkop(pdf, "Top 10 Hoogste Risico's", size=14)
    pdf.set_font("Arial", '', 11)

    for _, row in top_10_risico.iterrows():
        kenmerk = str(row['kenmerken']).strip()
        omschrijving = str(row['Kenmerken van het bouwwerk']).strip()
        score = str(row['Risico']).strip()

        # Titel
        pdf.set_font("Arial", 'B', 11)
        pdf.multi_cell(col_widths[i], line_height, str(header_cell), border=1, align='C', ln=3, max_line_height=pdf.font_size)

        # Omschrijving
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(0, 5, omschrijving)

        # Score rechts onderaan de paragraaf
        pdf.set_font("Arial", 'I', 9)
        pdf.cell(0, 6, f"Score: {score}", align='R')
        pdf.ln(6)  # Extra ruimte tussen de risico's
        
    # PAGINA 4: Grafiek
    pdf.add_page()
    pdf_subkop(pdf, "Visualisatie: Risicokenmerken Top 10", size=14)
    if top10_kenmerken_path:
        pdf.image(top10_kenmerken_path, x=15, y=pdf.get_y(), w=180)
    
    pdf_result = pdf.output(dest="S")
    pdf_bytes = bytes(pdf_result)
    
    return pdf_bytes


################################################################################
# Excel Helpers
################################################################################
def voeg_matrix_toe(schrijver, df):
    wb = schrijver.book
    ws = wb.add_worksheet('Risicomatrix')
    ws.write(0, 0, "Dynamische Risicomatrix")
    for j, col in enumerate(df.columns, start=1):
        ws.write(1, j, col)
    for i, row_name in enumerate(df.index, start=2):
        ws.write(i, 0, row_name)
        for j, col_name in enumerate(df.columns, start=1):
            val = df.loc[row_name, col_name]
            ws.write(i, j, val)
def voeg_tolerantie_toe(schrijver, df_tolerantie):
    df_tolerantie.to_excel(schrijver, sheet_name="Tolerantie Analyse", index=False)
def voeg_top10_toe(schrijver, top_10):
    wb = schrijver.book
    ws = wb.add_worksheet("Hoogste Risicos")
    top_10.to_excel(schrijver, sheet_name="Hoogste Risicos", index=False)
    diagram = wb.add_chart({'type': 'bar'})
    diagram.add_series({
        'categories': ['Hoogste Risicos', 1, 0, len(top_10), 0],
        'values': ['Hoogste Risicos', 1, 2, len(top_10), 2],
        'fill': {'color': '#EC6907'},
        'name': 'Risicoscore'
    })
    diagram.set_title({'name': 'Top 10 Hoogste Risicos'})
    ws.insert_chart('F2', diagram)
def voeg_staafdiagram_toe(schrijver, df):
    """
    Zet in 'Risicokenmerken' sheet de som van Risico per groep + bar chart
    """
    wb = schrijver.book
    ws = wb.add_worksheet('Risicokenmerken')
    df['Groep'] = df['kenmerken'].str.split(' - ').str[0]
    gegroepeerd = df.groupby('Groep', as_index=False)['Risico'].sum()
    gegroepeerd = gegroepeerd.sort_values(by='Risico', ascending=True)
    gegroepeerd.to_excel(schrijver, sheet_name="Risicokenmerken", startrow=1, index=False)

    diagram = wb.add_chart({'type': 'bar'})
    diagram.add_series({
        'categories': ['Risicokenmerken', 2, 0, len(gegroepeerd)+1, 0],
        'values': ['Risicokenmerken', 2, 1, len(gegroepeerd)+1, 1],
        'fill': {'color': '#EC6907'},
        'name': 'Risicoscore'
    })
    diagram.set_title({'name': 'Meest Risicovolle Kenmerken'})
    ws.insert_chart('F2', diagram)
################################################################################
# Hoofdfunctie (Streamlit)
################################################################################
def hoofd():
    voeg_logo_toe("bcc_volantis-logo_cmyk.jpg")
    st.title("Volantis: RGB Tool Risicoanalyse 📊")
    geupload_bestand = st.file_uploader("Upload Excel-bestand", type=["xlsx", "xls"])
    if geupload_bestand:
        try:
            # Lees ruwe data
            ruwe_df = pd.read_excel(geupload_bestand, header=None)
            # Zoek de start_rij
            start_rij = ruwe_df[
                ruwe_df.apply(
                    lambda rij: rij.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(),
                    axis=1
                )
            ].index
            if not start_rij.empty:
                verwerkte_df = ruwe_df.iloc[start_rij[0]:].copy()
                getransformeerde_df = transformeer_gegevens(verwerkte_df)
                # Data voor staafdiagram (som per groep, etc.)
                getransformeerde_df['Groep'] = getransformeerde_df['kenmerken'].str.split(' - ').str[0]
                gegroepeerde_risicos = getransformeerde_df.groupby('Groep')['Risico'].sum().reset_index()
                kol1, kol2, kol3 = st.columns(3)
                with kol1:
                    risico_drempel = st.select_slider(
                        "Toon risico's met een score boven:",
                        options=list(range(0, 31)),
                        value=0
                    )
                options = ['Alle groepen'] + list(getransformeerde_df['Groep'].unique())
                with kol2:
                    selected_group = st.selectbox("Kies de Groep", options=options)
                options_kenmerken = ['Alle kenmerken'] + list(getransformeerde_df['kenmerken'].unique())
                with kol3:
                    selected_kenmerken = st.selectbox("Kies de kenmerken", options=options_kenmerken)
                df_filtered = getransformeerde_df.copy()
                if selected_group != 'Alle groepen':
                    df_filtered = df_filtered[df_filtered['Groep'] == selected_group]
                if selected_kenmerken != 'Alle kenmerken':
                    df_filtered = df_filtered[df_filtered['kenmerken'] == selected_kenmerken]
                gefilterde_risicos = df_filtered[df_filtered["Risico"] >= risico_drempel]
                st.write(f"**Risico’s boven drempel ({risico_drempel})**")
                st.dataframe(gefilterde_risicos, hide_index=True)
                st.subheader("Risicoanalyse")
                kol_links, kol_rechts = st.columns(2)
                with kol_links:
                    st.write("**Dynamische Risicomatrix**")
                    risicomatrix = maak_risicomatrix(getransformeerde_df)
                    st.dataframe(
                        risicomatrix.style.apply(style_tolerantie_matrix, axis=None),
                        height=(len(risicomatrix) + 1) * 35 + 3
                    )
                    st.write("**Tolerantie - Risico Analyse**")
                    werkelijke_waarden = maak_risico_categorieen(getransformeerde_df)
                    df_risicomatrix = pd.DataFrame({
                        "Risc SC": ["Veilig", "Laag", "Medium", "Hoog"],
                        "Risc Level SC": ["0 > 5", "6", "7 > 9", "10 > 25"],
                        "Aantal": [
                            werkelijke_waarden.get('Laag', 0), 
                            werkelijke_waarden.get('Midden', 0), 
                            werkelijke_waarden.get('Hoog', 0), 
                            werkelijke_waarden.get('Zeer hoog', 0)
                        ]
                    })
                    st.dataframe(df_risicomatrix, hide_index=True)
                with kol_rechts:
                    st.write("**Risicokenmerken**")
                    if not gegroepeerde_risicos.empty:
                        fig2, ax2 = plt.subplots(figsize=(8,6))
                        gegroepeerde_som = getransformeerde_df.groupby('Groep', as_index=False)['Risico'].sum()
                        gegroepeerde_som.columns = ['Groep', 'Risico']
                        
                        sns.barplot(
                            data=gegroepeerde_som.sort_values('Risico', ascending=False),
                            y='Groep',
                            x='Risico',
                            palette=aangepaste_kleuren,
                            ax=ax2
                        )
                        ax2.set_xlabel("Totaalscore Risico", fontsize=12)
                        ax2.set_title("Risicokenmerken (Som)")
                        plt.tight_layout()
                        st.pyplot(fig2)
                    else:
                        st.info("Geen risico's gevonden")
                kol1_top10, kol2_top10 = st.columns(2)
                with kol1_top10:
                    st.write("**Top 10 Hoogste Risico's**")
                    if len(getransformeerde_df) >= 10:
                        top_10_risico = getransformeerde_df.nlargest(10, "Risico")[
                            ["kenmerken", "Kenmerken van het bouwwerk", "Risico"]
                        ]
                        top_10_risico['Categorie'] = pd.cut(
                            top_10_risico['Risico'], 
                            bins=[0, 5, 10, 20, 30],
                            labels=['Laag', 'Midden', 'Hoog', 'Zeer hoog'],
                            right=False
                        )
                        st.dataframe(top_10_risico, hide_index=True)
                    else:
                        top_10_risico = pd.DataFrame()
                        st.info("Niet genoeg gegevens om de top 10 te bepalen.")
                with kol2_top10:
                    st.write("**Risicokenmerken top 10**")
                    if not top_10_risico.empty:
                        fig3, ax3 = plt.subplots(figsize=(8,9))
                        sns.barplot(
                            data=top_10_risico.sort_values('Risico', ascending=False),
                            y='Kenmerken van het bouwwerk',
                            x='Risico',
                            hue='Categorie',
                            palette=categorie_palet,
                            dodge=False,
                            ax=ax3
                        )
                        ax3.set_xlabel("Risicoscore", fontsize=12)
                        ax3.set_ylabel("")
                        plt.tight_layout()
                        st.pyplot(fig3)
                    else:
                        st.info("Geen risico's gevonden.")
                # === Excel-Export ===
                uitvoer = BytesIO()
                with pd.ExcelWriter(uitvoer, engine="xlsxwriter") as schrijver:
                    getransformeerde_df.to_excel(schrijver, index=False, sheet_name="RuweData")
                    voeg_matrix_toe(schrijver, risicomatrix)
                    voeg_staafdiagram_toe(schrijver, getransformeerde_df)
                    voeg_tolerantie_toe(schrijver, df_risicomatrix)
                    if not top_10_risico.empty:
                        voeg_top10_toe(schrijver, top_10_risico)
                    
                st.download_button(
                    label="📥 Download Volledig Rapport (Excel)",
                    data=uitvoer.getvalue(),
                    file_name="risico_analyse.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                # === PDF-Export ===
                fig2_pdf, ax2_pdf = plt.subplots(figsize=(8,6))
                gegroepeerde_som = getransformeerde_df.groupby('Groep', as_index=False)['Risico'].sum()
                sns.barplot(
                    data=gegroepeerde_som.sort_values('Risico', ascending=False),
                    y='Groep',
                    x='Risico',
                    palette=aangepaste_kleuren,
                    ax=ax2_pdf
                )
                ax2_pdf.set_xlabel("Totaalscore Risico", fontsize=12)
                ax2_pdf.set_title("Risicokenmerken (Som)")
                plt.tight_layout()
                tmp_kenmerken_som = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                fig2_pdf.savefig(tmp_kenmerken_som.name, bbox_inches='tight')
                plt.close(fig2_pdf)
                tmp_top10_kenmerken = None
                if not top_10_risico.empty:
                    fig_top10, ax_top10 = plt.subplots(figsize=(8,9))
                    sns.barplot(
                        data=top_10_risico.sort_values('Risico', ascending=False),
                        y='Kenmerken van het bouwwerk',
                        x='Risico',
                        hue='Categorie',
                        palette=categorie_palet,
                        dodge=False,
                        ax=ax_top10
                    )
                    ax_top10.set_xlabel("Risicoscore", fontsize=12)
                    ax_top10.set_ylabel("")
                    plt.tight_layout()
                    tmp_top10_kenmerken = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    fig_top10.savefig(tmp_top10_kenmerken.name, bbox_inches='tight')
                    plt.close(fig_top10)
                    
                pdf_bytes = genereer_pdf(
                    top_10_risico=top_10_risico,
                    tolerantie_df=df_risicomatrix,
                    risicokenmerken_path=tmp_kenmerken_som.name,
                    top10_kenmerken_path=tmp_top10_kenmerken.name if tmp_top10_kenmerken else None,
                    risicomatrix=risicomatrix
                )
                if pdf_bytes:
                    st.download_button(
                        "📄 Download PDF-rapport",
                        data=pdf_bytes,
                        file_name="Risicoanalyse.pdf",
                        mime="application/pdf"
                    )   
                else:
                    st.info("Geen top 10 risico's om in het PDF-rapport te zetten.")
            else:
                st.error("Startrij niet gevonden in het document (kan 'Omgevingskenmerken - Algemeen' niet vinden).")
                
        except Exception as e:
            st.error(f"Fout opgetreden: {str(e)}")
if __name__ == "__main__":
    hoofd()