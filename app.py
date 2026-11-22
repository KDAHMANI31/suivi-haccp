import streamlit as st
import pandas as pd
from datetime import date
import base64

# Configuration de la page
st.set_page_config(page_title="Suivi des Non-Conformités HACCP", layout="wide")

st.title("📋 Tableau de Suivi des Non-Conformités (HACCP)")
st.markdown("Ajoutez vos constats d'inspection, insérez des preuves photos et exportez votre plan d'action.")

# Initialisation de la mémoire (session_state) pour conserver les données
if "donnees" not in st.session_state:
    st.session_state.donnees = []

# --- FORMULAIRE DE SAISIE ---
with st.expander("➕ Ajouter une nouvelle non-conformité", expanded=True):
    with st.form("formulaire_nc", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            zone = st.selectbox("Zone concernée", [
                "Cuisine Centrale", "Pâtisserie", "Boucherie", "Économat",
                "Cantine", "Vestiaires Hommes", "Cuisine Canastel", 
                "Restaurant Le Canastel", "Cuisine Banquet", "Plonge / Laverie", "Autre"
            ])
            anomalie = st.text_area("Anomalie constatée", placeholder="Ex: Siphon bouché, néon grillé...")
            action = st.text_input("Action corrective demandée", placeholder="Ex: Débouchage, remplacement néon...")
            
        with col2:
            responsable = st.text_input("Responsable", placeholder="Ex: Service Technique, Chef Steward...")
            date_echeance = st.date_input("Date d'échéance", value=date.today())
            remarque = st.text_input("Remarque / Statut", placeholder="Ex: Urgent, En cours...")
            photo_upload = st.file_uploader("Insérer une photo depuis le PC", type=["png", "jpg", "jpeg"])

        bouton_ajouter = st.form_submit_button("✅ Enregistrer dans le tableau")

        if bouton_ajouter:
            if anomalie.strip() == "":
                st.error("⚠️ Veuillez décrire l'anomalie avant d'enregistrer.")
            else:
                # Conversion de l'image en format lisible par le tableau
                photo_url = None
                nom_photo = "Aucune photo"
                if photo_upload is not None:
                    bytes_data = photo_upload.getvalue()
                    b64 = base64.b64encode(bytes_data).decode()
                    photo_url = f"data:{photo_upload.type};base64,{b64}"
                    nom_photo = photo_upload.name

                # Ajout de la ligne
                nouvelle_ligne = {
                    "Zone": zone,
                    "Anomalie": anomalie,
                    "Action": action,
                    "Responsable": responsable,
                    "Photo": photo_url,
                    "Fichier Photo": nom_photo,
                    "Date d'échéance": date_echeance.strftime("%d/%m/%Y"),
                    "Remarque": remarque
                }
                st.session_state.donnees.append(nouvelle_ligne)
                st.success("Non-conformité ajoutée avec succès !")

# --- AFFICHAGE DU TABLEAU ---
st.divider()
st.subheader("📊 Résultat du tableau d'inspection")

if len(st.session_state.donnees) > 0:
    df = pd.DataFrame(st.session_state.donnees)
    
    # Affichage interactif avec support d'image
    st.data_editor(
        df[["Zone", "Anomalie", "Action", "Responsable", "Photo", "Date d'échéance", "Remarque"]],
        column_config={
            "Photo": st.column_config.ImageColumn("Preuve Photo", help="Aperçu de la photo importée")
        },
        use_container_width=True,
        num_rows="dynamic"
    )

    # --- BOUTONS D'EXPORT ET DE RÉINITIALISATION ---
    col_exp, col_reset = st.columns([1, 4])
    
    with col_exp:
        # Export CSV compatible Excel (séparateur point-virgule et encodage UTF-8)
        df_export = df.drop(columns=["Photo"])
        csv = df_export.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            label="📥 Télécharger le tableau (Excel/CSV)",
            data=csv,
            file_name=f"Suivi_HACCP_{date.today().strftime('%d_%m_%Y')}.csv",
            mime="text/csv"
        )
        
    with col_reset:
        if st.button("🗑️ Vider tout le tableau"):
            st.session_state.donnees = []
            st.rerun()
else:
    st.info("Aucune anomalie enregistrée pour le moment. Utilisez le formulaire ci-dessus pour commencer.")
