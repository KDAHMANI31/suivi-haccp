import streamlit as st
import pandas as pd
from datetime import date
import base64
import io
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

# Configuration de la page
st.set_page_config(page_title="Suivi des Non-Conformités HACCP", layout="wide")

st.title("📋 Tableau de Suivi des Non-Conformités (HACCP)")
st.markdown("Ajoutez vos constats d'inspection, insérez des preuves photos et générez un **rapport Excel professionnel**.")

if "donnees" not in st.session_state:
    st.session_state.donnees = []

# --- FORMULAIRE DE SAISIE ---
with st.expander("➕ Ajouter une nouvelle non-conformité", expanded=True):
    with st.form("formulaire_nc", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            zone = st.selectbox("Zone concernée", (
                "Cuisine Centrale", "Pâtisserie", "Boucherie", "Économat",
                "Cantine", "Vestiaires Hommes", "Cuisine Canastel", 
                "Restaurant Le Canastel", "Cuisine Banquet", "Plonge / Laverie", "Autre"
            ))
            anomalie = st.text_area("Anomalie constatée", placeholder="Ex: Chambre froide 107 non opérationnelle...")
            action = st.text_input("Action corrective demandée", placeholder="Ex: Réparation compresseur, transfert marchandises...")
            
        with col2:
            responsable = st.text_input("Responsable", placeholder="Ex: Technique, Chef Steward...")
            date_echeance = st.date_input("Date d'échéance", value=date.today())
            remarque = st.text_input("Remarque / Statut", placeholder="Ex: Urgent, En cours...")
            photo_upload = st.file_uploader("Insérer une photo depuis le PC", type=("png", "jpg", "jpeg"))

        bouton_ajouter = st.form_submit_button("✅ Enregistrer dans le tableau")

        if bouton_ajouter:
            if anomalie.strip() == "":
                st.error("⚠️ Veuillez décrire l'anomalie avant d'enregistrer.")
            else:
                photo_url = None
                photo_bytes = None
                if photo_upload is not None:
                    photo_bytes = photo_upload.getvalue()
                    b64 = base64.b64encode(photo_bytes).decode()
                    photo_url = f"data:{photo_upload.type};base64,{b64}"

                nouvelle_ligne = {
                    "Zone": zone,
                    "Anomalie": anomalie,
                    "Action": action,
                    "Responsable": responsable,
                    "Photo": photo_url,
                    "Photo_Bytes": photo_bytes,
                    "Date d'échéance": date_echeance.strftime("%d/%m/%Y"),
                    "Remarque": remarque
                }
                st.session_state.donnees.append(nouvelle_ligne)
                st.success("Non-conformité ajoutée avec succès !")

# --- FONCTION DE CRÉATION DU FICHIER EXCEL PROFESSIONNEL ---
def generer_excel_pro(donnees):
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi HACCP"

    # Titre du document (Ligne 1)
    ws.merge_cells("A1:G1")
    ws["A1"] = f"RAPPORT D'INSPECTION HYGIÈNE & SÉCURITÉ ALIMENTAIRE (HACCP) - {date.today().strftime('%d/%m/%Y')}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions.setdefault(1).height = 35

    # En-têtes des colonnes (Ligne 3)
    entetes = ("Zone", "Anomalie constatée", "Action corrective", "Responsable", "Preuve Photo", "Date d'échéance", "Remarque / Statut")
    header_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E0'),
        right=Side(style='thin', color='CBD5E0'),
        top=Side(style='thin', color='CBD5E0'),
        bottom=Side(style='thin', color='CBD5E0')
    )

    ws.row_dimensions.setdefault(3).height = 28
    for col_num, titre in enumerate(entetes, 1):
        cell = ws.cell(row=3, column=col_num, value=titre)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Largeur des colonnes bien ordonnées
    largeurs = {"A": 22, "B": 38, "C": 35, "D": 20, "E": 22, "F": 18, "G": 22}
    for col_lettre, largeur in largeurs.items():
        ws.column_dimensions[col_lettre].width = largeur

    # Remplissage des données
    fill_pair = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    fill_impair = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for idx, item in enumerate(donnees, start=4):
        row_fill = fill_pair if idx % 2 == 0 else fill_impair
        
        valeurs = (
            item["Zone"],
            item["Anomalie"],
            item["Action"],
            item["Responsable"],
            "",
            item["Date d'échéance"],
            item["Remarque"]
        )

        for col_num, val in enumerate(valeurs, 1):
            cell = ws.cell(row=idx, column=col_num, value=val)
            cell.fill = row_fill
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=11)
            
            # Alignement centré pour Zone (1), Responsable (4) et Date (6)
            if col_num in (1, 4, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            # Mise en évidence en rouge si "Urgent" dans Remarque
            if col_num == 7 and "urgent" in str(val).lower():
                cell.font = Font(name="Calibri", size=11, bold=True, color="C53030")
                cell.fill = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")

        # Insertion de la photo dans la cellule E si présente
        if item["Photo_Bytes"] is not None:
            ws.row_dimensions[idx].height = 95
            pil_img = PILImage.open(io.BytesIO(item["Photo_Bytes"]))
            pil_img.thumbnail((130, 115))
            img_buffer = io.BytesIO()
            pil_img.save(img_buffer, format="PNG")
            img_buffer.seek(0)
            
            xl_img = XLImage(img_buffer)
            ws.add_image(xl_img, f"E{idx}")
        else:
            ws.cell(row=idx, column=5, value="Aucune photo").alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[idx].height = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- AFFICHAGE DU TABLEAU ---
st.divider()
st.subheader("📊 Aperçu du tableau d'inspection")

if len(st.session_state.donnees) > 0:
    df = pd.DataFrame(st.session_state.donnees)
    
    st.data_editor(
        df[["Zone", "Anomalie", "Action", "Responsable", "Photo", "Date d'échéance", "Remarque"]],
        column_config={
            "Photo": st.column_config.ImageColumn("Preuve Photo", help="Aperçu de la photo importée")
        },
        use_container_width=True,
        num_rows="dynamic"
    )

    col_exp, col_reset = st.columns(2)
    
    with col_exp:
        fichier_excel = generer_excel_pro(st.session_state.donnees)
        st.download_button(
            label="📥 Télécharger le Tableau Excel Pro (.xlsx)",
            data=fichier_excel,
            file_name=f"Rapport_Suivi_HACCP_{date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_reset:
        if st.button("🗑️ Vider tout le tableau"):
            st.session_state.donnees = []
            st.rerun()
else:
    st.info("Aucune anomalie enregistrée pour le moment. Utilisez le formulaire ci-dessus pour commencer.")
