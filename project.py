from fpdf import FPDF
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import re
from xlsxwriter.utility import xl_rowcol_to_cell
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide", page_title="Risico Analyse 📊")

# Aangepaste kleuren
aangepaste_kleuren = [
    '#D0DFE6', '#FBCDAB', 
    '#EC6907', '#A6CEE3', 
    '#B2DF8A', '#FDBF6F', 
    '#CAB2D6', '#FF7F00', 
    '#FB9A99']

categorie_palet = {
    'Laag': '#A6CEE3',  
    'Midden': '#BFD8D2',  
    'Hoog': '#FFA54F', 
    'Zeer hoog': '#E24E1B'
}

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

def transformeer_gegevens(invoer_gegevens):
    invoer_gegevens.columns = invoer_gegevens.iloc[0]
    invoer_gegevens = invoer_gegevens.drop(invoer_gegevens.index[0])

    getransponeerd = invoer_gegevens.T.reset_index()
    getransponeerd.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
    if getransponeerd.shape[0] > 22:
        getransponeerd = getransponeerd.iloc[22:].reset_index(drop=True)

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

def extraheer_nummer(waarde):
    if isinstance(waarde, str):
        match = re.search(r'\((\d+)\)', waarde)
        if match:
            return int(match.group(1))
        else:
            match = re.search(r'\d+', waarde)
            return int(match.group()) if match else 0
    return 0

def maak_risicomatrix(df):
    y_labels = ['Catastrofaal (5)', 'Zeer ernstig (4)', 'Ernstig (3)', 'Matig (2)', 'Licht (1)']
    x_labels = ['Zeer laag (1)', 'Laag (2)', 'Gemiddeld (3)', 'Hoog (4)', 'Zeer hoog (5)']

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

def maak_risico_categorieen(df):
    bins = [-1, 5, 10, 20, 30]
    labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
    df['Categorie'] = pd.cut(df['Risico'], bins=bins, labels=labels, right=False)
    return df['Categorie'].value_counts().reindex(labels, fill_value=0)

def voeg_matrix_toe(schrijver, matrix):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Risicomatrix')
    
    werkblad.write(0, 0, "Kans →\Effect ↓")
    for kolom_idx, waarde in enumerate(matrix.columns, start=1):
        werkblad.write(0, kolom_idx, waarde)
    
    for rij_idx, waarde in enumerate(matrix.index, start=1):
        werkblad.write(rij_idx, 0, waarde)
    
    for rij_idx, kans_label in enumerate(matrix.index, start=1):
        for kolom_idx, effect_label in enumerate(matrix.columns, start=1):
            kans = extraheer_nummer(kans_label)
            effect = extraheer_nummer(effect_label)
            product = kans * effect
            telling = matrix.loc[kans_label, effect_label]
            werkblad.write(rij_idx, kolom_idx, f"{product} ({telling})")
    
    max_waarde = matrix.max().max()
    for rij in range(1, len(matrix)+1):
        for kolom in range(1, len(matrix.columns)+1):
            cel_ref = xl_rowcol_to_cell(rij, kolom)
            werkblad.conditional_format(
                cel_ref, {
                    'type': '2_color_scale',
                    'min_value': 0,
                    'max_value': max_waarde,
                    'min_color': '#FFFFFF',
                    'max_color': '#EC6907'
                }
            )

def voeg_tolerantie_toe(schrijver, df_tolerantie):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Tolerantie Analyse')
    df_tolerantie.to_excel(schrijver, sheet_name='Tolerantie Analyse', index=False)

def voeg_top10_toe(schrijver, top_10):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Hoogste Risicos')
    top_10.to_excel(schrijver, sheet_name='Hoogste Risicos', index=False)
    
    diagram = werkboek.add_chart({'type': 'bar'})
    diagram.add_series({
        'categories': ['Hoogste Risicos', 1, 0, len(top_10), 0],
        'values':     ['Hoogste Risicos', 1, 2, len(top_10), 2],
        'fill':       {'color': '#EC6907'},
        'name':       'Risicoscore'
    })
    diagram.set_title({'name': 'Top 10 Hoogste Risicos'})
    werkblad.insert_chart('F2', diagram)

def voeg_staafdiagram_toe(schrijver, df):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Risicokenmerken')
    
    df['Groep'] = df['kenmerken'].str.split(' - ').str[0]
    gegroepeerd = df.groupby('Groep', as_index=False)['Risico'].sum()
    gegroepeerd = gegroepeerd.sort_values(by='Risico', ascending=True)

    bins = [0, 25, 50, 75, 100]
    labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
    gegroepeerd['Categorie'] = pd.cut(gegroepeerd['Risico'], bins=bins, labels=labels, right=False)

    gegroepeerd[['Groep', 'Risico', 'Categorie']].to_excel(
        schrijver, 
        sheet_name='Risicokenmerken', 
        startrow=1, 
        index=False
    )

    diagram = werkboek.add_chart({'type': 'bar'})
    diagram.add_series({
        'categories': ['Risicokenmerken', 2, 0, len(gegroepeerd) + 1, 0],
        'values':     ['Risicokenmerken', 2, 1, len(gegroepeerd) + 1, 1],
        'fill':       {'color': '#EC6907'},
        'name':       'Risicoscore'
    })
    
    diagram.set_title({'name': 'Meest Risicovolle Kenmerken'})
    werkblad.insert_chart('E2', diagram)

def hoofd():
    voeg_logo_toe("bcc_volantis-logo_cmyk.jpg")
    st.title("Volantis: Geavanceerde Risicoanalyse 📊")
    geupload_bestand = st.file_uploader("Upload Excel-bestand", type=["xlsx", "xls"])
    
    if geupload_bestand:
        try:
            ruwe_df = pd.read_excel(geupload_bestand, header=None)
            start_rij = ruwe_df[ruwe_df.apply(lambda rij: rij.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(), axis=1)].index
            
            if not start_rij.empty:
                verwerkte_df = ruwe_df.iloc[start_rij[0]:].copy()
                getransformeerde_df = transformeer_gegevens(verwerkte_df)
                
                # Data voor staafdiagram
                getransformeerde_df['Groep'] = getransformeerde_df['kenmerken'].str.split(' - ').str[0]
                gegroepeerde_risicos = getransformeerde_df.groupby('Groep')['Risico'].sum().reset_index()

                bins = [0, 25, 50, 75, 100]
                labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
                gegroepeerde_risicos['Categorie'] = pd.cut(gegroepeerde_risicos['Risico'], bins=bins, labels=labels, right=False)
                
                kol1, kol2, kol3 = st.columns(3)
                
                with kol1:
                    risico_drempel = st.select_slider(
                        "Toon risico's met een score boven:",
                        options=list(range(0, 26)),  # Een lijst met waarden van 0 t/m 30
                        value=0  # Standaard geselecteerde waarde
                    )
                    # risico_drempel = st.select_slider(
                    #     "Toon risico's met een score tussen:",
                    #     options=list(range(0, 31)),  # Een lijst met waarden van 0 t/m 30
                    #     value=(5, 20)  # Standaard geselecteerde bereik van 5 tot 20
                    # )

                options = ['Alle groepen'] + list(getransformeerde_df['Groep'])
                
                # Opties genereren inclusief 'Alles'
                options = ['Alle groepen'] + list(getransformeerde_df['Groep'].unique())
                options_kenmerken = ['Alle kenmerken'] + list(getransformeerde_df['kenmerken'].unique())
                
                with kol2:
                # Selectie dropdown
                    selected_group = st.selectbox("Kies de Groep", options=options)
                    
                with kol3:
                    selected_kenmerken = st.selectbox("Kies de kenmerken", options=options_kenmerken)

                # Filter de dataset op de geselecteerde groep (indien niet 'Alles')
                if selected_group != 'Alle groepen':
                    getransformeerde_df = getransformeerde_df[getransformeerde_df['Groep'] == selected_group]

                if selected_kenmerken != 'Alle kenmerken':
                    getransformeerde_df = getransformeerde_df[getransformeerde_df['kenmerken'] == selected_kenmerken]
                gefilerde_risicos = getransformeerde_df[getransformeerde_df["Risico"] >= risico_drempel]
                st.write(f"**Risico’s boven drempel ({risico_drempel})**")
                st.dataframe(gefilerde_risicos, hide_index=True)
                st.subheader("Risicoanalyse")
                kol1, kol2 = st.columns(2)
                
                with kol1:
                    st.write("**Risicomatrix**")
                    risicomatrix = maak_risicomatrix(getransformeerde_df)
                    fig, ax = plt.subplots(figsize=(8, 6))
                    sns.heatmap(
                        risicomatrix.astype(int), 
                        annot=True, 
                        fmt="d",
                        cmap="Oranges",
                        linewidths=.5,
                        linecolor='white',
                        ax=ax
                    )
                    ax.set_xlabel("Effect", fontsize=12)
                    ax.set_ylabel("Kans", fontsize=12)
                    st.pyplot(fig)
                
                    st.write("**Tolerantie - Risico Analyse**")
                    werkelijke_waarden = maak_risico_categorieen(getransformeerde_df)
                    
                    df_risicomatrix = pd.DataFrame(
                        {
                            "Risc SC": ["Safe", "Low", "Medium", "High"],
                            "Risc Level SC": ["0 > 5", "6", "7 > 9", "10 > 25"],
                            "Aantal": [
                                werkelijke_waarden.get('Laag', 0), 
                                werkelijke_waarden.get('Midden', 0), 
                                werkelijke_waarden.get('Hoog', 0), 
                                werkelijke_waarden.get('Zeer hoog', 0)
                                ],
                            "Gewenste Acties": ["Geen direct actie", "Herstel binnen 3 maanden", "Herstel binnen 1 maand", "Onmiddelijke herstelactie"]
                        }
                    )
                    st.dataframe(df_risicomatrix, hide_index=True)
                    

                with kol2:
                    st.write("**Risicokenmerken**")
                    if not gegroepeerde_risicos.empty:
                        fig2, ax2 = plt.subplots(figsize=(8,6))
                        
                        sns.barplot(
                            data=gegroepeerde_risicos.sort_values('Risico', ascending=False),
                            y='Groep',
                            x='Risico',
                            hue='Categorie',
                            palette=categorie_palet,
                            dodge=False,
                            ax=ax2
                        )
                        ax2.set_xlabel("Risicoscore", fontsize=12)
                        ax2.set_ylabel("")
                        plt.tight_layout()
                        st.pyplot(fig2)
                    else:
                        st.info("Geen risico's gevonden")
                    
                
                kol1, kol2 = st.columns(2)
                with kol1:               
                    st.write("**Top 10 Hoogste Risico's**")

                    # Zorg ervoor dat de DataFrame minimaal 10 rijen heeft voordat nlargest() wordt uitgevoerd
                    if len(getransformeerde_df) >= 10:
                        top_10_risico = getransformeerde_df.nlargest(10, "Risico")[["kenmerken", "Kenmerken van het bouwwerk", "Risico"]]
                        
                        # Voeg de risicocategorie toe voor de bar chart
                        bins = [0, 5, 10, 20, 30]
                        labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
                        top_10_risico['Categorie'] = pd.cut(top_10_risico['Risico'], bins=bins, labels=labels, right=False)
                        
                        st.dataframe(top_10_risico, hide_index=True)
                    
                    else:
                        st.info("Niet genoeg gegevens om de top 10 te bepalen.")

                with kol2:
                    st.write("**Risicokenmerken top 10**")

                    if 'top_10_risico' in locals() and not top_10_risico.empty:
                        fig3, ax2 = plt.subplots(figsize=(8,9))

                        sns.barplot(
                            data=top_10_risico.sort_values('Risico', ascending=False),
                            y='Kenmerken van het bouwwerk',
                            x='Risico',
                            hue='Categorie',  # Nu veilig om te gebruiken omdat de categorieën zijn toegevoegd
                            palette=categorie_palet,
                            dodge=False,
                            ax=ax2
                        )
                        ax2.set_xlabel("Risicoscore", fontsize=12)
                        ax2.set_ylabel("")
                        plt.tight_layout()
                        st.pyplot(fig3)
                    else:
                        st.info("Geen risico's gevonden.")

                
                uitvoer = BytesIO()
                with pd.ExcelWriter(uitvoer, engine="xlsxwriter") as schrijver:
                    getransformeerde_df.to_excel(schrijver, index=False, sheet_name="RuweData")
                    getransformeerde_df.to_excel(schrijver, index=False, sheet_name="RuweData")
                    voeg_matrix_toe(schrijver, risicomatrix)
                    voeg_staafdiagram_toe(schrijver, getransformeerde_df)
                    voeg_tolerantie_toe(schrijver, df_risicomatrix)
                    voeg_top10_toe(schrijver, top_10_risico)
                    
                st.download_button(
                    label="📥 Download Volledig Rapport",
                    data=uitvoer.getvalue(),
                    file_name="risico_analyse.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            else:
                st.error("Startrij niet gevonden in het document")
                
        except Exception as e:
            st.error(f"Fout opgetreden: {str(e)}")

    def genereer_pdf():
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.set_title("test")
        pdf.cell(200, 10, "Risicoanalyse Rapport", ln=True, align="C")
        pdf.ln(10)
        
        for index, row in top_10_risico.iterrows():
            pdf.cell(200, 10, f"{row['Risico']}: {row['kenmerken']}", ln=True)
        
        # Return PDF bytes directly
        return pdf.output(dest='S')

    st.download_button("📄 Download PDF-rapport", data=genereer_pdf(), file_name="Risicoanalyse.pdf", mime="application/pdf")

if __name__ == "__main__":
    hoofd()