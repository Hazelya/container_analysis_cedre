import streamlit as st
import pandas as pd
import io

def process_data(input_file):
    # 1. Chargement du fichier Excel
    # On saute les 4 premières lignes
    df = pd.read_excel(input_file, skiprows=4)

    # 2. Nettoyage des noms de colonnes
    df.columns = [str(c).strip() for c in df.columns]

    # 3. Filtrage : uniquement les lignes avec une 'Class' (IMDG)
    if 'Class' in df.columns:
        df_dg = df[df['Class'].notna()].copy()
    else:
        st.error("La colonne 'Class' est introuvable dans le fichier.")
        return None

    # 4. Sélection et renommage
    mapping = {
        'Slot': 'Localisation',
        'Declared goods': 'Nom',
        'Class': 'IMD',
        'Weight': 'Quantité'
    }

    # On ne garde que ce qui existe réellement dans le fichier
    cols_to_use = [k for k in mapping.keys() if k in df_dg.columns]
    resultat = df_dg[cols_to_use].rename(columns=mapping)
    
    return resultat

# --- Interface Streamlit ---
st.set_page_config(page_title="Extracteur Manifeste Cargo", layout="wide", page_icon="🚢")
st.title("🚢 Extracteur de Manifeste Marchandises Dangereuses")
st.write("Téléchargez votre fichier Excel pour extraire uniquement les lignes IMDG.")

# Correction : on demande un Excel car pd.read_excel est utilisé
uploaded_file = st.file_uploader("Téléchargez le manifeste (Excel)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        with st.spinner('Traitement des données...'):
            df_final = process_data(uploaded_file)

        if df_final is not None and not df_final.empty:
            st.success(f"Extraction terminée : {len(df_final)} marchandises dangereuses trouvées.")
            
            # Aperçu des données
            st.dataframe(df_final, use_container_width=True)

            # Export Excel vers la mémoire
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='IMDG_Extract')
            
            st.download_button(
                label="📥 Télécharger l'extraction Excel",
                data=output.getvalue(),
                file_name="dg_manifeste_extrait.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif df_final is not None:
            st.warning("Aucune marchandise dangereuse (colonne 'Class' remplie) n'a été trouvée.")
            
    except Exception as e:
        st.error(f"Une erreur est survenue lors de la lecture du fichier : {e}")
