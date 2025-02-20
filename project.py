# import streamlit as st
# import pandas as pd
# import numpy as np
# from io import BytesIO
# import base64
# import re
# from xlsxwriter.utility import xl_rowcol_to_cell
# import matplotlib.pyplot as plt
# import seaborn as sns

# # Aangepaste kleuren
# aangepaste_kleuren = ['#D0DFE6', '#FBCDAB', '#EC6907', '#A6CEE3', '#B2DF8A', 
#                      '#FDBF6F', '#CAB2D6', '#FF7F00', '#FB9A99']

# def voeg_logo_toe(afbeeldingspad):
#     with open(afbeeldingspad, "rb") as bestand:
#         inhoud = bestand.read()
#         gecodeerde_afbeelding = base64.b64encode(inhoud).decode()
#     st.markdown(
#         f"""
#         <div style="text-align: center;">
#             <img src="data:image/png;base64,{gecodeerde_afbeelding}" width="300">
#         </div>
#         """,
#         unsafe_allow_html=True,
#     )

# def transformeer_gegevens(invoer_gegevens):
#     invoer_gegevens.columns = invoer_gegevens.iloc[0]
#     invoer_gegevens = invoer_gegevens.drop(invoer_gegevens.index[0])
#     getransponeerd = invoer_gegevens.T.reset_index()
    
#     getransponeerd.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
#     if getransponeerd.shape[0] > 22:
#         getransponeerd = getransponeerd.iloc[22:].reset_index(drop=True)
    
#     verwerkte_rijen = []
#     sla_volgende_over = False
    
#     for idx in range(len(getransponeerd)):
#         if sla_volgende_over:
#             sla_volgende_over = False
#             continue
            
#         huidige_rij = getransponeerd.iloc[idx].copy()
#         kans_waarde = str(huidige_rij["Kans"]).strip().lower()
#         effect_waarde = str(huidige_rij["Effect"]).strip().lower()
        
#         if "inschatting kans" in kans_waarde:
#             huidige_rij["Kans"] = huidige_rij["Effect"]
            
#             if idx + 1 < len(getransponeerd):
#                 volgende_rij = getransponeerd.iloc[idx + 1]
#                 if "inschatting effect" in str(volgende_rij["Kans"]).lower():
#                     huidige_rij["Effect"] = volgende_rij["Effect"]
#                     sla_volgende_over = True
        
#         elif "inschatting effect" in kans_waarde:
#             continue
        
#         def extraheer_nummer(waarde):
#             match = re.search(r'\d+', str(waarde))
#             return int(match.group()) if match else 0
        
#         huidige_rij["Risico"] = extraheer_nummer(huidige_rij["Kans"]) * extraheer_nummer(huidige_rij["Effect"])
#         verwerkte_rijen.append(huidige_rij)
    
#     return pd.DataFrame(verwerkte_rijen)

# def maak_risicomatrix(df):
#     matrix_labels = ['1', '2', '3', '4', '5']
#     matrix = pd.DataFrame(
#         0, 
#         index=pd.Categorical(matrix_labels, categories=matrix_labels),
#         columns=pd.Categorical(matrix_labels, categories=matrix_labels)
#     )
    
#     def extraheer_nummer(waarde):
#         match = re.search(r'\d+', str(waarde))
#         return int(match.group()) if match else 0
    
#     for _, rij in df.iterrows():
#         kans = extraheer_nummer(rij['Kans'])
#         effect = extraheer_nummer(rij['Effect'])
#         if kans in [1,2,4,5] and effect in [1,2,4,5]:
#             matrix.loc[str(kans), str(effect)] += 1
    
#     return matrix

# def maak_risico_categorieen(df):
#     bins = [-1, 5, 6, 9, 25]
#     labels = ['Veilig (0-5)', 'Laag (6)', 'Medium (7-9)', 'Hoog (10-25)']
#     df['Categorie'] = pd.cut(df['Risico'], bins=bins, labels=labels)
#     return df['Categorie'].value_counts().reindex(labels, fill_value=0)

# def voeg_matrix_toe(schrijver, matrix):
#     werkboek = schrijver.book
#     werkblad = werkboek.add_worksheet('Risicomatrix')
    
#     # Header
#     werkblad.write(0, 0, "Kans →\Effect ↓")
#     for kolom_idx, waarde in enumerate(matrix.columns, start=1):
#         werkblad.write(0, kolom_idx, waarde)
    
#     # Index
#     for rij_idx, waarde in enumerate(matrix.index, start=1):
#         werkblad.write(rij_idx, 0, waarde)
    
#     # Data met product en telling
#     for rij_idx, kans in enumerate(matrix.index, start=1):
#         for kolom_idx, effect in enumerate(matrix.columns, start=1):
#             product = int(kans) * int(effect)
#             telling = matrix.loc[kans, effect]
#             werkblad.write(rij_idx, kolom_idx, f"{product} ({telling})")
    
#     # Opmaak
#     max_waarde = matrix.max().max()
#     for rij in range(1, len(matrix)+1):
#         for kolom in range(1, len(matrix.columns)+1):
#             cel_ref = xl_rowcol_to_cell(rij, kolom)
#             werkblad.conditional_format(
#                 cel_ref, {
#                     'type': '2_color_scale',
#                     'min_value': 0,
#                     'max_value': max_waarde,
#                     'min_color': '#FFFFFF',
#                     'max_color': '#EC6907'
#                 }
#             )

# def voeg_staafdiagram_toe(schrijver, df):
#     werkboek = schrijver.book
#     werkblad = werkboek.add_worksheet('Risicokenmerken')
    
#     # Filter en sorteer
#     gefilterd = df[df['Risico'] > 0].sort_values('Risico', ascending=False)
    
#     # Schrijf data
#     gefilterd[['kenmerken', 'Risico']].to_excel(
#         schrijver, 
#         sheet_name='Risicokenmerken', 
#         startrow=1, 
#         index=False
#     )
    
#     # Maak diagram
#     diagram = werkboek.add_chart({'type': 'bar'})
#     diagram.add_series({
#         'categories': ['Risicokenmerken', 1, 0, len(gefilterd), 0],
#         'values':     ['Risicokenmerken', 1, 1, len(gefilterd), 1],
#         'fill':       {'color': '#EC6907'},
#         'name':       'Risicoscore'
#     })
    
#     # Diagraminstellingen
#     diagram.set_title({'name': 'Meest Risicovolle Kenmerken'})
#     diagram.set_y_axis({'name': 'Kenmerk', 'label_position': 'laag'})
#     diagram.set_x_axis({'name': 'Risicoscore'})
#     diagram.set_legend({'position': 'geen'})
    
#     werkblad.insert_chart('D2', diagram)

# def hoofd():
#     voeg_logo_toe("bcc_volantis-logo_cmyk.jpg")
#     st.title("Volantis: Geavanceerde Risicoanalyse")
#     geupload_bestand = st.file_uploader("Upload Excel-bestand", type=["xlsx", "xls"])
    
#     if geupload_bestand:
#         try:
#             ruwe_df = pd.read_excel(geupload_bestand, header=None)
#             start_rij = ruwe_df[ruwe_df.apply(lambda rij: rij.astype(str).str.contains("Omgevingskenmerken - Algemeen").any(), axis=1)].index
            
#             if not start_rij.empty:
#                 verwerkte_df = ruwe_df.iloc[start_rij[0]:].copy()
#                 getransformeerde_df = transformeer_gegevens(verwerkte_df)
                
#                 st.subheader("Verwerkte Gegevens")
#                 st.dataframe(getransformeerde_df)
                
#                 st.subheader("Risicoanalyse")
                
#                 kol1, kol2 = st.columns(2)
                
#                 with kol1:
#                     st.write("**Risicomatrix**")
#                     risicomatrix = maak_risicomatrix(getransformeerde_df)
#                     fig, ax = plt.subplots(figsize=(8, 6))
#                     sns.heatmap(
#                         risicomatrix.astype(int), 
#                         annot=True, 
#                         fmt="d",
#                         cmap="YlOrRd",
#                         linewidths=.5,
#                         linecolor='white',
#                         ax=ax
#                     )
#                     ax.set_xlabel("Kans", fontsize=12)
#                     ax.set_ylabel("Effect", fontsize=12)
#                     st.pyplot(fig)
                
#                 with kol2:
#                     st.write("**Risicokenmerken**")
#                     gefilterd = getransformeerde_df[getransformeerde_df['Risico'] > 0]
#                     if not gefilterd.empty:
#                         fig2, ax2 = plt.subplots(figsize=(8, 6))
#                         sns.barplot(
#                             data=gefilterd.sort_values('Risico', ascending=False),
#                             y='kenmerken',
#                             x='Risico',
#                             palette="YlOrRd",
#                             ax=ax2
#                         )
#                         ax2.set_xlabel("Risicoscore", fontsize=12)
#                         ax2.set_ylabel("")
#                         plt.tight_layout()
#                         st.pyplot(fig2)
#                     else:
#                         st.info("Geen risico's gevonden")
                
#                 uitvoer = BytesIO()
#                 with pd.ExcelWriter(uitvoer, engine="xlsxwriter") as schrijver:
#                     getransformeerde_df.to_excel(schrijver, index=False, sheet_name="RuweData")
#                     voeg_matrix_toe(schrijver, risicomatrix)
#                     voeg_staafdiagram_toe(schrijver, getransformeerde_df)
                
#                 st.download_button(
#                     label="📥 Download Volledig Rapport",
#                     data=uitvoer.getvalue(),
#                     file_name="risico_analyse.xlsx",
#                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                 )
            
#             else:
#                 st.error("Startrij niet gevonden in het document")
                
#         except Exception as e:
#             st.error(f"Fout opgetreden: {str(e)}")

# if __name__ == "__main__":
#     hoofd()

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import re
from xlsxwriter.utility import xl_rowcol_to_cell
import matplotlib.pyplot as plt
import seaborn as sns

# Aangepaste kleuren
aangepaste_kleuren = ['#D0DFE6', '#FBCDAB', '#EC6907', '#A6CEE3', '#B2DF8A', 
                     '#FDBF6F', '#CAB2D6', '#FF7F00', '#FB9A99']

# Categoriekleuren voor risiconiveaus
categorie_palet = {
    'Laag': '#D0DFE6',
    'Midden': '#FBCDAB',
    'Hoog': '#EC6907',
    'Zeer hoog': '#A6CEE3'
}

def voeg_logo_toe(afbeeldingspad):
    with open(afbeeldingspad, "rb") as bestand:
        inhoud = bestand.read()
        gecodeerde_afbeelding = base64.b64encode(inhoud).decode()
    st.markdown(
        f"""
        <div style="text-align: center;">
            <img src="data:image/png;base64,{gecodeerde_afbeelding}" width="300">
        </div>
        """,
        unsafe_allow_html=True,
    )



# def transformeer_gegevens(invoer_gegevens):
#     invoer_gegevens.columns = invoer_gegevens.iloc[0]
#     invoer_gegevens = invoer_gegevens.drop(invoer_gegevens.index[0])
#     getransponeerd = invoer_gegevens.T.reset_index()
    
#     getransponeerd.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
#     if getransponeerd.shape[0] > 22:
#         getransponeerd = getransponeerd.iloc[22:].reset_index(drop=True)
    
#     verwerkte_rijen = []
#     sla_volgende_over = False
    
#     for idx in range(len(getransponeerd)):
#         if sla_volgende_over:
#             sla_volgende_over = False
#             continue
            
#         huidige_rij = getransponeerd.iloc[idx].copy()
#         kans_waarde = str(huidige_rij["Kans"]).strip().lower()
#         effect_waarde = str(huidige_rij["Effect"]).strip().lower()
        
#         if "inschatting kans" in kans_waarde:
#             huidige_rij["Kans"] = huidige_rij["Effect"]
            
#             if idx + 1 < len(getransponeerd):
#                 volgende_rij = getransponeerd.iloc[idx + 1]
#                 if "inschatting effect" in str(volgende_rij["Kans"]).lower():
#                     huidige_rij["Effect"] = volgende_rij["Effect"]
#                     sla_volgende_over = True
        
#         elif "inschatting effect" in kans_waarde:
#             continue
        
#         def extraheer_nummer(waarde):
#             match = re.search(r'\d+', str(waarde))
#             return int(match.group()) if match else 0
        
#         huidige_rij["Risico"] = extraheer_nummer(huidige_rij["Kans"]) * extraheer_nummer(huidige_rij["Effect"])
#         verwerkte_rijen.append(huidige_rij)
    
#     return pd.DataFrame(verwerkte_rijen)

def transformeer_gegevens(invoer_gegevens):
    # Stap 1: Gebruik de bovenste rij als headers
    invoer_gegevens.columns = invoer_gegevens.iloc[0]
    invoer_gegevens = invoer_gegevens.drop(invoer_gegevens.index[0])

    # Stap 2: Transponeer en zet index in kolom
    getransponeerd = invoer_gegevens.T.reset_index()
    getransponeerd.columns = ['kenmerken', 'Kenmerken van het bouwwerk', 'Kans', 'Effect']
    
    # Stap 3: Indien er meer dan 22 rijen zijn, snijd ze af
    if getransponeerd.shape[0] > 22:
        getransponeerd = getransponeerd.iloc[22:].reset_index(drop=True)

    # >>> Nieuw toegevoegd: vul lege cellen in 'kenmerken' met de laatst bekende niet-lege waarde
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
            # Schuif de 'Effect' cel naar de 'Kans' kolom
            huidige_rij["Kans"] = huidige_rij["Effect"]

            # Kijk of de volgende rij de "inschatting effect" bevat
            if idx + 1 < len(getransponeerd):
                volgende_rij = getransponeerd.iloc[idx + 1]
                if "inschatting effect" in str(volgende_rij["Kans"]).lower():
                    huidige_rij["Effect"] = volgende_rij["Effect"]
                    sla_volgende_over = True
        
        elif "inschatting effect" in kans_waarde:
            # Deze rij alleen overslaan
            continue
        
        # Bepaal risicoscore
        def extraheer_nummer(waarde):
            match = re.search(r'\d+', str(waarde))
            return int(match.group()) if match else 0
        
        huidige_rij["Risico"] = extraheer_nummer(huidige_rij["Kans"]) * extraheer_nummer(huidige_rij["Effect"])
        verwerkte_rijen.append(huidige_rij)

    return pd.DataFrame(verwerkte_rijen)


def maak_risicomatrix(df):
    y_labels = ['5', '4', '3', '2', '1']  # Omgekeerde volgorde voor y-as
    x_labels = ['1', '2', '3', '4', '5']
    matrix = pd.DataFrame(
        0, 
        index=pd.Categorical(y_labels, categories=y_labels, ordered=True),
        columns=pd.Categorical(x_labels, categories=x_labels, ordered=True)
    )
    
    def extraheer_nummer(waarde):
        match = re.search(r'\d+', str(waarde))
        return int(match.group()) if match else 0
    
    for _, rij in df.iterrows():
        kans = extraheer_nummer(rij['Kans'])
        effect = extraheer_nummer(rij['Effect'])
        if 1 <= kans <=5 and 1 <= effect <=5:
            matrix.loc[str(kans), str(effect)] += 1
    
    return matrix

def maak_risico_categorieen(df):
    bins = [-1, 5, 10, 20, 30]
    labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
    df['Categorie'] = pd.cut(df['Risico'], bins=bins, labels=labels, right=False)
    return df['Categorie'].value_counts().reindex(labels, fill_value=0)

def voeg_matrix_toe(schrijver, matrix):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Risicomatrix')
    
    # Header
    werkblad.write(0, 0, "Kans →\Effect ↓")
    for kolom_idx, waarde in enumerate(matrix.columns, start=1):
        werkblad.write(0, kolom_idx, waarde)
    
    # Index
    for rij_idx, waarde in enumerate(matrix.index, start=1):
        werkblad.write(rij_idx, 0, waarde)
    
    # Data met product en telling
    for rij_idx, kans in enumerate(matrix.index, start=1):
        for kolom_idx, effect in enumerate(matrix.columns, start=1):
            product = int(kans) * int(effect)
            telling = matrix.loc[kans, effect]
            werkblad.write(rij_idx, kolom_idx, f"{product} ({telling})")
    
    # Opmaak
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

def voeg_staafdiagram_toe(schrijver, df):
    werkboek = schrijver.book
    werkblad = werkboek.add_worksheet('Risicokenmerken')
    
    # Groepeer en sommeer
    df['Groep'] = df['kenmerken'].str.split(' - ').str[0]
    gegroepeerd = df.groupby('Groep', as_index=False)['Risico'].sum()

    # Categoriseer
    bins = [-1, 5, 10, 20, 30]
    labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
    gegroepeerd['Categorie'] = pd.cut(gegroepeerd['Risico'], bins=bins, labels=labels, right=False)
    
    # Schrijf data
    gegroepeerd[['Groep', 'Risico', 'Categorie']].to_excel(
        schrijver, 
        sheet_name='Risicokenmerken', 
        startrow=1, 
        index=False
    )
    
    # Maak diagram
    diagram = werkboek.add_chart({'type': 'bar'})
    diagram.add_series({
        'categories': ['Risicokenmerken', 1, 0, len(gegroepeerd), 0],
        'values':     ['Risicokenmerken', 1, 1, len(gegroepeerd), 1],
        'fill':       {'color': '#EC6907'},
        'name':       'Risicoscore'
    })
    
    # Diagraminstellingen
    diagram.set_title({'name': 'Meest Risicovolle Kenmerken'})
    diagram.set_y_axis({'name': 'Kenmerk', 'label_position': 'laag'})
    diagram.set_x_axis({'name': 'Risicoscore'})
    diagram.set_legend({'position': 'geen'})
    
    werkblad.insert_chart('D2', diagram)

def hoofd():
    voeg_logo_toe("bcc_volantis-logo_cmyk.jpg")
    st.title("Volantis: Geavanceerde Risicoanalyse")
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
                # gegroepeerde_risicos = getransformeerde_df.groupby('Groep', as_index=False)['Risico'].sum()
                gegroepeerde_risicos = getransformeerde_df.groupby('Groep')['Risico'].sum().reset_index()

                # gegroepeerde_groepen = getransformeerde_df.groupby('Groep', as_index=False)['Groep'].count()
                print('groep', gegroepeerde_risicos)
                bins = [0, 25, 50, 75, 100]
                labels = ['Laag', 'Midden', 'Hoog', 'Zeer hoog']
                gegroepeerde_risicos['Categorie'] = pd.cut(gegroepeerde_risicos['Risico'], bins=bins, labels=labels, right=False)
                
                st.subheader("Verwerkte Gegevens")
                st.dataframe(getransformeerde_df)
                
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
                        cmap="YlOrRd",
                        linewidths=.5,
                        linecolor='white',
                        ax=ax
                    )
                    ax.set_xlabel("Effect", fontsize=12)
                    ax.set_ylabel("Kans", fontsize=12)
                    st.pyplot(fig)
                
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

                
                uitvoer = BytesIO()
                with pd.ExcelWriter(uitvoer, engine="xlsxwriter") as schrijver:
                    getransformeerde_df.to_excel(schrijver, index=False, sheet_name="RuweData")
                    voeg_matrix_toe(schrijver, risicomatrix)
                    voeg_staafdiagram_toe(schrijver, getransformeerde_df)
                
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

if __name__ == "__main__":
    hoofd()