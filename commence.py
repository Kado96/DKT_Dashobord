import streamlit as st
import pandas as pd
import folium
from folium.plugins import MarkerCluster, HeatMap, Fullscreen, Draw
from streamlit_folium import folium_static
import plotly.express as px
from datetime import date, timedelta
from UI import *
from add_data import *
import plotly.graph_objects as go
import io
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver

# Configuration de la page
st.set_page_config(page_title="DKT2", page_icon="🌍", layout="wide")
st.header(":bar_chart: DKT2 Dashboard")



# Nom du fichier
file_name = 'CampagneDKT001_-_all_versions_-_labels_-_2024-11-27-10-51-31.xlsx'

# Charger les feuilles du fichier Excel
df_unilever = pd.read_excel(file_name, sheet_name='CampagneDKT001')
df_gpi = pd.read_excel(file_name, sheet_name='GPI')
df_sondage = pd.read_excel(file_name, sheet_name='Sondage')

print("Fichiers chargés avec succès.")

# Sélection des colonnes spécifiques
df_unilever_cols = ["_index", "_submission_time", "Nom de Point De Vente","Nom et prénom du proprietaire?","Numéro de téléphone","Type du PDV", "Province", "Commune", "Quartier", 
                    "Adresse", "Y a-t-il eu un achat?", "Nom et prénom",
                    "Gestion de commandes et crédits", "Entrez la date et l'heure de livraison du commande :", "Entrez la date et l'heure du paiment du crédit :",
                    "_Prendre les coordonnées du point de vente_latitude", "Le PDV a t- il été recruté?", "Quels sont vos commentaires généraux ou ceux du vendeur sur le point de vente?", 
                    "_Prendre les coordonnées du point de vente_longitude"]
df_gpi_cols = ["_index", "_submission__submission_time"]
df_sondage_cols = ["_index", "Sorte_caracteristic", "Combien de ${Sorte_caracteristic} avez-vous vendus?", 
                   "Montant de la vente", "_submission__submission_time"]


# Extraire seulement les colonnes nécessaires pour réduire la taille des DataFrames
df_unilever = df_unilever[df_unilever_cols]
df_gpi = df_gpi[df_gpi_cols]
df_sondage = df_sondage[df_sondage_cols]

# Fusionner les DataFrames (corriger l'identifiant de fusion si nécessaire)
df_merged = pd.merge(df_unilever, df_gpi, on='_index', how='left')
df_merged = pd.merge(df_merged, df_sondage, on='_index', how='left')

# Filtrage par date
date1 = st.sidebar.date_input("Choose a start date")
date2 = st.sidebar.date_input("Choose an end date")
date1 = pd.to_datetime(date1)
date2 = pd.to_datetime(date2) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
df_filtered = df_merged[(df_merged["_submission_time"] >= date1) & (df_merged["_submission_time"] <= date2)]

# Filtres supplémentaires
st.sidebar.header("Additional filters :")
filters = {
    "Commune": st.sidebar.multiselect("Commune", sorted(df_filtered["Commune"].unique())),
    "Quartier": st.sidebar.multiselect("Quartier", sorted(df_filtered["Quartier"].unique())),
    "Nom et prénom": st.sidebar.multiselect("Agent", sorted(df_filtered["Nom et prénom"].unique())),
    "Nom et prénom du proprietaire?": st.sidebar.multiselect("Proprietaire", sorted(df_filtered["Nom et prénom du proprietaire?"].fillna("").unique())),
    "Gestion de commandes et crédits": st.sidebar.multiselect("Commandes et crédits", sorted(df_filtered["Gestion de commandes et crédits"].astype(str).unique()))  # Conversion en str
}

for col, selection in filters.items():
    if selection:
        df_filtered = df_filtered[df_filtered[col].isin(selection)]

# Bloc analytique
if df_filtered is not None and not df_filtered.empty:
    with st.expander("VIEW EXCEL DATASET"):
        showData = st.multiselect('Filter: ', df_filtered.columns, default=df_unilever_cols)
        st.dataframe(df_filtered[showData], use_container_width=True)

        # Exporter les données filtrées au format Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_filtered.to_excel(writer, index=False, sheet_name="Données filtrées")

        processed_data = output.getvalue()

        st.download_button(
            label="📥 Télécharger les données filtrées en format Excel",
            data=processed_data,
            file_name="données_filtrées.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.warning("Aucune donnée n'a été récupérée. Veuillez vérifier votre URL ou votre clé API.")



# Affichage de la carte
with st.expander("Mapping"):
    if df_filtered['_Prendre les coordonnées du point de vente_latitude'].isnull().all() or \
       df_filtered['_Prendre les coordonnées du point de vente_longitude'].isnull().all():
        st.error("Les coordonnées de localisation sont toutes manquantes.")
    else:
        latitude_mean = df_filtered['_Prendre les coordonnées du point de vente_latitude'].mean()
        longitude_mean = df_filtered['_Prendre les coordonnées du point de vente_longitude'].mean()
        m = folium.Map(location=[latitude_mean, longitude_mean], zoom_start=4)
        marker_cluster = MarkerCluster().add_to(m)

        for _, row in df_filtered.iterrows():
            if pd.notnull(row['_Prendre les coordonnées du point de vente_latitude']) and \
               pd.notnull(row['_Prendre les coordonnées du point de vente_longitude']):
                
                # Lien Google Maps pour obtenir l'itinéraire vers le point de vente (nous utiliserons des coordonnées statiques pour l'exemple)
                # Vous pourrez remplacer ces coordonnées avec celles de l'utilisateur une fois la géolocalisation récupérée
                google_maps_url = f"https://www.google.com/maps/dir/?api=1&origin=YOUR_LATITUDE,YOUR_LONGITUDE&destination={row['_Prendre les coordonnées du point de vente_latitude']},{row['_Prendre les coordonnées du point de vente_longitude']}&travelmode=driving"

                popup_content = f"""
                    <h3>Informations sur {row['Nom de Point De Vente']}</h3>
                    <div style='color:gray; font-size:14px;'>
                        <b>Nom de l'agent :</b> {row['Nom et prénom']}<br>
                        <b>Nom et prénom du proprietaire? :</b> {row['Nom et prénom du proprietaire?']}<br>
                        <b>Type du PDV :</b> {row['Type du PDV']}<br>
                        <b>Commune :</b> {row['Commune']}<br>
                        <b>Adresse :</b> {row['Adresse']}<br>
                        <b>Gestion de commandes et crédits :</b> {row['Gestion de commandes et crédits']}<br>
                        <b>Date de livraison de Commandes :</b> {row.get("Entrez la date et l'heure de livraison du commande :", "Non spécifié")}<br>
                        <b>Date de paiement de crédit :</b> {row.get("Entrez la date et l'heure du paiment du crédit :", "Non spécifié")}<br>
                        <b>Numéro de téléphone :</b> {row['Numéro de téléphone']}<br>
                        <b>Date d'enregistrement :</b> {row['_submission_time']}<br>
                        <b>Voir sur la carte :</b> <a href="{google_maps_url}" target="_blank">Cliquer ici pour obtenir l'itinéraire vers ce point de vente</a>
                    </div>
                """

                folium.Marker(
                    location=[row['_Prendre les coordonnées du point de vente_latitude'], 
                              row['_Prendre les coordonnées du point de vente_longitude']],
                    tooltip=row['Nom de Point De Vente'],
                    icon=folium.Icon(color='red', icon='fa-dollar-sign', prefix='fa')
                ).add_to(marker_cluster).add_child(folium.Popup(popup_content, max_width=600))

        # Ajout de la heatmap
        heat_data = [[row['_Prendre les coordonnées du point de vente_latitude'], 
                      row['_Prendre les coordonnées du point de vente_longitude']] 
                     for _, row in df_filtered.iterrows()
                     if pd.notnull(row['_Prendre les coordonnées du point de vente_latitude']) and 
                        pd.notnull(row['_Prendre les coordonnées du point de vente_longitude'])]
        if heat_data:
            HeatMap(heat_data).add_to(m)
        Fullscreen(position='topright').add_to(m)
        Draw(export=True).add_to(m)

        # Affichage de la carte
        folium_static(m)

