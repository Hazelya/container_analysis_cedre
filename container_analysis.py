import streamlit as st
import pandas as pd
import io

def process_data(input_file):
    # On charge le fichier à partor de l'en-tête
    df = pd.read_excel(input_file, skiprows=4)

    # On enlève les espaces inutiles
    df.columns = [str(c).strip() for c in df.columns]

    # On récupère les marchandises dangereuses 
    if 'Class' in df.columns:
        df_dg = df[df['Class'].notna()].copy()
    else:
        st.error("La colonne 'Class' est introuvable dans le fichier.") # Affiche si une erreur
        return None

    # Les colonnes que l'on compte utiliser
    mapping = {
        'Slot': 'Localisation',
        'Declared goods': 'Nom',
        'Class': 'IMD',
        'Weight': 'Quantité'
    }

    # On garde les colonne qu'on veut
    cols_to_use = [k for k in mapping.keys() if k in df_dg.columns]
    resultat = df_dg[cols_to_use].rename(columns=mapping)
    
    return resultat

st.set_page_config(page_title="Extracteur Manifeste Cargo", layout="wide")
st.title("Analyse des marchandises")
st.write("Téléchargez votre fichier Excel...")

# On récupère le fichier
uploaded_file = st.file_uploader("Téléchargez le manifeste (Excel)", type=["xlsx", "xls"])

if uploaded_file: # si un fichier
    try:
        with st.spinner('Traitement des données...'): 
            df_final = process_data(uploaded_file) # On appel la fonction

        if df_final is not None and not df_final.empty:
            st.success(f"Extraction terminée : {len(df_final)} marchandises dangereuses trouvées.") # Afficher le nombre de marchandises
            
            st.dataframe(df_final, use_container_width=True) # Affichage des données

            output = io.BytesIO() # Chemin vers les fichiers 
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='IMDG_Extract') # Ecrire le fichier excel en direct
            
            # boutton pour permettre le téléchargement
            st.download_button(
                label="Télécharger l'Excel final",
                data=output.getvalue(),
                file_name="dg_manifeste_extrait.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif df_final is not None:
            st.warning("Aucune marchandise dangereuse (colonne 'Class' remplie) n'a été trouvée.")
            
    except Exception as e: # Prendre en compte les erreurs
        st.error(f"Une erreur est survenue lors de la lecture du fichier : {e}")
