import streamlit as st
import pandas as pd
import datetime
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# -----------------------------------------------------------------------------
# CONFIGURATION DE LA PAGE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Adhésions - Café du Coin",
    page_icon="☕",
    layout="wide"
)

st.title("☕ Café du Coin - Gestion des Adhésions")

# -----------------------------------------------------------------------------
# GESTION DES DONNÉES ET CHARGEMENT DU FICHIER
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Données adhérents")
fichier_charge = st.sidebar.file_uploader(
    "Importer le fichier Excel des adhérents", 
    type=["xlsx", "xls"]
)

# Initialisation des données dans la session Streamlit
if "df_db" not in st.session_state:
    st.session_state.df_db = pd.DataFrame(columns=[
        "nom", "prenom", "email", "commune", "adherent_2025", "adherent_2026", "adherent_2027", "date_derniere_adhesion", "bad_email"
    ])

# Si l'utilisateur charge un fichier Excel via l'uploader
if fichier_charge is not None:
    try:
        # Tente de lire les feuilles du fichier Excel déposé
        wb = openpyxl.load_workbook(fichier_charge)
        
        # Exemple de lecture si le fichier a la structure complète
        if 'Adhérent.es 2025 et 2026' in wb.sheetnames:
            ws1 = wb['Adhérent.es 2025 et 2026']
            ws2 = wb.get('Adhérent.es 2026', ws1)

            membres_2026 = set()
            for r in range(8, ws2.max_row + 1):
                nom_val = ws2.cell(r, 1).value
                prenom_val = ws2.cell(r, 2).value
                email_val = ws2.cell(r, 3).value
                if nom_val or prenom_val or email_val:
                    k = f"{str(nom_val).strip().upper()}_{str(prenom_val).strip().title()}_{str(email_val).strip().lower()}"
                    membres_2026.add(k)

            data_master = []
            for r in range(9, ws1.max_row + 1):
                nom_v = ws1.cell(r, 1).value
                prenom_v = ws1.cell(r, 2).value
                email_v = ws1.cell(r, 3).value
                commune_v = ws1.cell(r, 4).value
                
                c_email = ws1.cell(r, 3)
                fill = c_email.fill.fgColor.value if c_email.fill else None
                bad_email = True if fill == 'FFFF0000' else False

                if nom_v or prenom_v or email_v or commune_v:
                    nom_str = str(nom_v).strip().upper() if nom_v else ''
                    prenom_str = str(prenom_v).strip().title() if prenom_v else ''
                    email_str = str(email_v).strip().lower() if email_v else ''
                    commune_str = str(commune_v).strip().title() if commune_v else ''
                    
                    key = f"{nom_str}_{prenom_str}_{email_str}"
                    is_2026 = key in membres_2026
                    
                    data_master.append({
                        "nom": nom_str,
                        "prenom": prenom_str,
                        "email": email_str,
                        "commune": commune_str,
                        "adherent_2025": True,
                        "adherent_2026": is_2026,
                        "adherent_2027": False,
                        "date_derniere_adhesion": "2026-10-04" if is_2026 else "2025-10-04",
                        "bad_email": bad_email
                    })
            st.session_state.df_db = pd.DataFrame(data_master)
        else:
            # Importation d'un tableau Excel simple
            df_simple = pd.read_excel(fichier_charge)
            st.session_state.df_db = df_simple

        st.sidebar.success("✅ Fichier Excel chargé avec succès !")

    except Exception as e:
        st.sidebar.error(f"Erreur de lecture : {e}")

df_db = st.session_state.df_db

# CRÉATION DES ONGLETS PRINCIPAUX
tab1, tab2 = st.tabs(["📝 Adhésion & Renouvellement", "⚙️ Espace Admin & Exports"])

# -----------------------------------------------------------------------------
# TAB 1 : ACCUEIL, RECHERCHE INTELLIGENTE & FORMULAIRE
# -----------------------------------------------------------------------------
with tab1:
    st.subheader("Adhésion & Renouvellement 2027")
    st.info("💡 **Déjà adhérent(e) ?** Tapez votre nom ci-dessous pour retrouver vos coordonnées et valider en 1 clic !")

    noms_existants = sorted(list(set(df_db["nom"].dropna().astype(str).str.upper()))) if not df_db.empty else []
    
    mode_saisie = st.radio("Sélectionnez votre situation :", ["Rechercher mon nom (Déjà adhérent)", "Nouveau / Je ne trouve pas mon nom"], horizontal=True)

    nom_pref, prenom_pref, email_pref, commune_pref = "", "", "", ""

    if mode_saisie == "Rechercher mon nom (Déjà adhérent)" and noms_existants:
        nom_selectionne = st.selectbox("Tapez ou choisissez votre Nom :", [""] + noms_existants)
        if nom_selectionne:
            fiches = df_db[df_db["nom"].str.upper() == nom_selectionne]
            if len(fiches) == 1:
                fiche = fiches.iloc[0]
                nom_pref = fiche["nom"]
                prenom_pref = fiche["prenom"]
                email_pref = fiche["email"]
                commune_pref = str(fiche["commune"]) if pd.notna(fiche["commune"]) else ""
            elif len(fiches) > 1:
                prenom_sel = st.selectbox("Plusieurs personnes portent ce nom, choisissez votre prénom :", fiches["prenom"].tolist())
                fiche = fiches[fiches["prenom"] == prenom_sel].iloc[0]
                nom_pref = fiche["nom"]
                prenom_pref = fiche["prenom"]
                email_pref = fiche["email"]
                commune_pref = str(fiche["commune"]) if pd.notna(fiche["commune"]) else ""

    # Formulaire principal
    with st.form("form_adhesion", clear_on_submit=False):
        st.write("---")
        nom = st.text_input("Nom *", value=nom_pref).strip().upper()
        prenom = st.text_input("Prénom *", value=prenom_pref).strip().title()
        email = st.text_input("Adresse e-mail *", value=email_pref).strip().lower()
        commune = st.text_input("Commune de résidence *", value=commune_pref).strip().title()

        submitted = st.form_submit_button("✅ Valider mon adhésion 2027")

        if submitted:
            if not nom or not prenom or not email or not commune:
                st.error("⚠️ Veuillez remplir tous les champs obligatoires.")
            elif "@" not in email or "." not in email:
                st.error("⚠️ Veuillez saisir une adresse e-mail valide.")
            else:
                mask = (df_db["email"].str.lower() == email) | ((df_db["nom"].str.upper() == nom) & (df_db["prenom"].str.title() == prenom)) if not df_db.empty else pd.Series([False])
                today_str = datetime.date.today().strftime("%Y-%m-%d")

                if not df_db.empty and mask.any():
                    df_db.loc[mask, "nom"] = nom
                    df_db.loc[mask, "prenom"] = prenom
                    df_db.loc[mask, "email"] = email
                    df_db.loc[mask, "commune"] = commune
                    df_db.loc[mask, "adherent_2027"] = True
                    df_db.loc[mask, "date_derniere_adhesion"] = today_str
                    st.success(f"🎉 Merci {prenom} ! Votre adhésion 2027 est renouvelée.")
                else:
                    nouvelle_fiche = {
                        "nom": nom,
                        "prenom": prenom,
                        "email": email,
                        "commune": commune,
                        "adherent_2025": False,
                        "adherent_2026": False,
                        "adherent_2027": True,
                        "date_derniere_adhesion": today_str,
                        "bad_email": False
                    }
                    df_db = pd.concat([df_db, pd.DataFrame([nouvelle_fiche])], ignore_index=True)
                    st.session_state.df_db = df_db
                    st.success(f"🎉 Bienvenue au Café du Coin, {prenom} ! Votre adhésion 2027 est enregistrée.")

# -----------------------------------------------------------------------------
# TAB 2 : ESPACE ADMIN & EXPORTS EXCEL / CSV LILO
# -----------------------------------------------------------------------------
with tab2:
    st.title("⚙️ Gestion & Exports administratifs")
    
    password = st.text_input("Mot de passe Bureau :", type="password")
    
    if password == "cafe2027":
        st.success("Accès autorisé au bureau.")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Base Historique", len(df_db))
        col2.metric("Adhérents Actifs 2027", len(df_db[df_db["adherent_2027"] == True]) if "adherent_2027" in df_db.columns else 0)
        col3.metric("Relances à faire (2026 non réadhérés)", len(df_db[(df_db["adherent_2026"] == True) & (df_db["adherent_2027"] != True)]) if "adherent_2026" in df_db.columns else 0)

        st.divider()

        # --- EXPORTS LILO ---
        st.subheader("📧 Exports Mailing pour Lilo.org")
        type_mailing = st.selectbox("Choisir la liste d'envoi à générer :", [
            "Tous les adhérents 2027 (Actifs)",
            "Nouveaux adhérents 2027 uniquement",
            "Relances : Adhérents 2026 n'ayant pas encore réadhéré en 2027",
            "Base globale historique complète"
        ])

        if not df_db.empty and "adherent_2027" in df_db.columns:
            if type_mailing == "Tous les adhérents 2027 (Actifs)":
                df_mail = df_db[df_db["adherent_2027"] == True]
            elif type_mailing == "Nouveaux adhérents 2027 uniquement":
                df_mail = df_db[(df_db["adherent_2027"] == True) & (df_db["adherent_2026"] == False)]
            elif type_mailing == "Relances : Adhérents 2026 n'ayant pas encore réadhéré en 2027":
                df_mail = df_db[(df_db["adherent_2026"] == True) & (df_db["adherent_2027"] != True)]
            else:
                df_mail = df_db.copy()

            df_lilo_export = df_mail[["email", "prenom", "nom", "commune"]] if not df_mail.empty else pd.DataFrame(columns=["email", "prenom", "nom", "commune"])
            csv_bytes = df_lilo_export.to_csv(index=False, encoding='utf-8-sig')

            st.download_button(
                label=f"📥 Télécharger le CSV Lilo ({len(df_lilo_export)} contacts)",
                data=csv_bytes,
                file_name=f"export_lilo_{type_mailing.split()[0].lower()}.csv",
                mime="text/csv"
            )

        st.divider()

        # --- EXPORT EXCEL OFFICIEL ---
        st.subheader("📊 Générer le classeur Excel Officiel (.xlsx)")

        def generer_excel_officiel(df_master):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            align_center = Alignment(horizontal="center", vertical="center")
            border_light = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                                  top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

            # Onglet 1 : Adhérents 2027
            ws1 = wb.active
            ws1.title = "Adhérent.es 2027"
            ws1.cell(row=1, column=1, value="Liste des Adhérent.es du Café du Coin").font = Font(size=14, bold=True)
            ws1.cell(row=2, column=4, value=f"Généré le : {datetime.date.today().strftime('%d/%m/%Y')}")
            
            headers = ["Nom", "Prénom", "Courriel", "Commune"]
            for col_num, h in enumerate(headers, 1):
                cell = ws1.cell(row=7, column=col_num, value=h)
                cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center

            df_2027 = df_master[df_master["adherent_2027"] == True].sort_values(by="nom") if "adherent_2027" in df_master.columns else pd.DataFrame()
            for r_idx, row in enumerate(df_2027.itertuples(), start=8):
                ws1.cell(row=r_idx, column=1, value=row.nom).border = border_light
                ws1.cell(row=r_idx, column=2, value=row.prenom).border = border_light
                ws1.cell(row=r_idx, column=3, value=row.email).border = border_light
                ws1.cell(row=r_idx, column=4, value=row.commune).border = border_light

            # Onglet 2 : Historique
            ws2 = wb.create_sheet(title="Adhérent.es 2026 et 2027")
            ws2.cell(row=1, column=1, value="Liste Historique des Adhérent.es").font = Font(size=14, bold=True)
            
            headers_hist = ["Nom", "Prénom", "Courriel", "Commune", "Statut 2027"]
            for col_num, h in enumerate(headers_hist, 1):
                cell = ws2.cell(row=7, column=col_num, value=h)
                cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center

            df_hist = df_master.sort_values(by="nom") if "nom" in df_master.columns else pd.DataFrame()
            for r_idx, row in enumerate(df_hist.itertuples(), start=8):
                statut = "Adhérent 2027" if getattr(row, 'adherent_2027', False) else ("Adhérent 2026" if getattr(row, 'adherent_2026', False) else "Ancien")
                ws2.cell(row=r_idx, column=1, value=row.nom).border = border_light
                ws2.cell(row=r_idx, column=2, value=row.prenom).border = border_light
                ws2.cell(row=r_idx, column=3, value=row.email).border = border_light
                ws2.cell(row=r_idx, column=4, value=row.commune).border = border_light
                ws2.cell(row=r_idx, column=5, value=statut).border = border_light

            # Largeurs automatiques
            for ws in [ws1, ws2]:
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(output)
            return output.getvalue()

        excel_bytes = generer_excel_officiel(df_db)

        st.download_button(
            label="📊 Télécharger le Fichier Excel Officiel (.xlsx)",
            data=excel_bytes,
            file_name=f"liste_adherents_cafe_du_coin_{datetime.date.today().year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )