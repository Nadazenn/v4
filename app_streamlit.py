import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import xlwings as xw
import tempfile
import unicodedata, re
import math
# Imports backend
import donnees as do
import parametrage as pa
import database as daba
import entrainement_modele as entmod
import pilotage_excel as pex
import database as daba
from openpyxl import load_workbook
import io


import pandas as pd
import plotly.express as px

import shutil
import tempfile

# ===================== PARAMÈTRES CCC — VARIABLES GLOBALES =====================

st.set_page_config(layout="wide")

st.markdown(
    "<h1 style='text-align: center;'> Logicast - Outil de chiffrage </h1>",
    unsafe_allow_html=True
)

st.markdown("""
⚠️ **Information importante :**  
Cet outil est une **première version de test**.  
Les fonctionnalités sont en cours de développement et peuvent évoluer.
""")

# Menu principal
menu = st.sidebar.radio(
    "Navigation",
    ["Paramétrage", "Données", "Pilotage Excel", "Dashboard", "Entraînement modèles", "Base de données"]
)


if menu == "Paramétrage":
    st.header("⚙️ Paramétrage")

    # --- Persistance de la page entière ---
    if "parametrage_page" not in st.session_state:
        st.session_state["parametrage_page"] = {}
    p = st.session_state["parametrage_page"]

    # --- Explication générale ---


    # --- Sélections principales ---
    entreprises = [f.replace("_logo.png", "") for f in os.listdir("images/logos_entreprises") if f.endswith(".png")]
    # Modèles disponibles
    models = [m.replace(".pkl", "") for m in os.listdir("models") if m.endswith(".pkl")]

    # Ajouter le modèle global
    if "GLOBAL" not in models:
        models.insert(0, "GLOBAL")

    p["entreprise_choice"] = st.selectbox(
        "Sélectionnez votre entreprise",
        entreprises,
        index=entreprises.index(p.get("entreprise_choice")) if "entreprise_choice" in p else 0,
    )

    # Modèles disponibles
    models = [m.replace(".pkl", "") for m in os.listdir("models") if m.endswith(".pkl")]

    # Ajouter le modèle global
    if "GLOBAL" not in models:
        models.insert(0, "GLOBAL")

    p["model_choice"] = st.selectbox(
        "Sélectionnez le modèle (GLOBAL ou spécifique)",
        models,
        index=models.index(p.get("model_choice")) if "model_choice" in p and p["model_choice"] in models else 0,
    )




    # --- Caractéristiques du bâtiment ---
    st.subheader("Caractéristiques du bâtiment")
    col1, col2, col3 = st.columns(3)
    with col1:
        p["nombre_etages"] = st.number_input("Nombre d'étages", min_value=1, value=p.get("nombre_etages", 1))
    with col2:
        p["zones_par_etage_defaut"] = st.number_input("Zones par étage par défaut", min_value=1, value=p.get("zones_par_etage_defaut", 1))
    with col3:
        p["numero_etage_inf"] = st.number_input("Numéro étage inférieur", value=p.get("numero_etage_inf", 0))

    # --- Planning ---
    st.subheader("Planning")
    if st.button("➕ Insérer un indice de planning"):
        p["show_popup_planning"] = True

    if p.get("show_popup_planning", False):
        p["planning_indice"] = st.text_input(
            "👉 Indiquez l'indice du planning :",
            value=p.get("planning_indice", "")
        )
        if p["planning_indice"]:
            st.success(f"Indice enregistré : {p['planning_indice']}")
            p["show_popup_planning"] = False

    st.markdown("""
    ℹ️ **Explication :**
    - **Date début Production** = début des travaux techniques de base (réseaux, câblage, équipements centraux).
    - **Date début Terminaux** = début de la pose des terminaux (prises, luminaires, capteurs, etc.), plus tard dans le chantier.
    """)

    col1, col2, col3 = st.columns(3)
    with col1:
        p["date_debut_prod"] = st.text_input("Début Production (JJ/MM/AAAA)", p.get("date_debut_prod", "01/01/2025"))
    with col2:
        p["date_debut_term"] = st.text_input("Début Terminaux (JJ/MM/AAAA)", p.get("date_debut_term", "01/05/2025"))
    with col3:
        p["intervalle_par_etage"] = st.number_input("Intervalle entre étages (jours)", min_value=0, value=p.get("intervalle_par_etage", 14))

    col1, col2, col3 = st.columns(3)
    with col1:
        p["delai_livraison"] = st.number_input("Délai livraison avant travaux (jours)", min_value=0, value=p.get("delai_livraison", 0))
    with col2:
        p["duree_prodmoyen_paretage"] = st.number_input("Durée moyenne Production (jours)", min_value=0, value=p.get("duree_prodmoyen_paretage", 30))
    with col3:
        p["duree_termmoyen_paretage"] = st.number_input("Durée moyenne Terminaux (jours)", min_value=0, value=p.get("duree_termmoyen_paretage", 30))

    # --- CCC ---
    # --- Activation CCC ---
    st.subheader("Utilisation de la CCC")

    use_ccc = st.radio(
        "Souhaitez-vous utiliser une CCC ?",
        ["Oui", "Non"],
        index=1 if p.get("use_ccc") is False else 0
    )

    p["use_ccc"] = (use_ccc == "Oui")

    if p.get("use_ccc", False):

        st.subheader("Caractéristiques du CCC")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            p["duree_stockage"] = st.number_input(
                "Durée stockage (mois)", min_value=0, value=p.get("duree_stockage", 2)
            )
        with col2:
            p["tarif_stockage"] = st.number_input(
                "Tarif stockage (€/mois)", value=p.get("tarif_stockage", 19)
            )
        with col3:
            p["frais_palette"] = st.number_input(
                "Frais palette (€)", value=p.get("frais_palette", 10)
            )
        with col4:
            p["frais_livraison"] = st.number_input(
                "Frais livraison (€)", value=p.get("frais_livraison", 175)
            )


    # --- Logistique du chantier ---
    st.subheader("Logistique du chantier")
    st.image("images/camions.png", caption="Camions disponibles")
    p["choix_camions"] = st.multiselect("Sélectionner les camions possibles", daba.liste_camions, default=p.get("choix_camions", daba.liste_camions))
    st.image("images/conditionnements.png", caption="Conditionnements disponibles")
    p["choix_conditionnement"] = st.multiselect("Sélectionner les conditionnements possibles", daba.liste_conditionnement, default=p.get("choix_conditionnement", daba.liste_conditionnement))

    # --- Étape 1 : Génération Étages / Zones ---
    if st.button("Valider Étages / Zones"):
        df1 = pa.generate_table(p["nombre_etages"], p["zones_par_etage_defaut"], p["numero_etage_inf"])
        p["output_table"] = df1
        st.success("✅ Étages / Zones générés")

    if "output_table" in p:
        st.dataframe(p["output_table"], use_container_width=True)

        if st.button("Créer le planning détaillé"):
            etages_zones = p["output_table"]["Numéro étage (pas de lettres)"].tolist()
            zones_per_etage = p["output_table"]["Nombre de zones"].tolist()
            df2 = pa.generate_details_table(
                etages_zones, zones_per_etage,
                p["delai_livraison"], p["date_debut_prod"], p["date_debut_term"],
                p["intervalle_par_etage"],
                p["duree_prodmoyen_paretage"], p["duree_termmoyen_paretage"]
            )
            p["output_details_table"] = df2
            st.success("✅ Planning généré")

    # --- Étape 2 : Tableau Détails (modifiable) ---
    if "output_details_table" in p:
        st.subheader("🗓️ Tableau Détails (modifiable)")
        df_base = p["output_details_table"]

        gb = GridOptionsBuilder.from_dataframe(df_base)
        gb.configure_pagination(enabled=True)
        gb.configure_default_column(editable=True, wrapText=True, autoHeight=True)
        grid_options = gb.build()

        grid_response = AgGrid(
            df_base,
            gridOptions=grid_options,
            data_return_mode='AS_INPUT',
            update_mode=GridUpdateMode.NO_UPDATE,
            fit_columns_on_grid_load=True,
            allow_unsafe_jscode=True,
            key="planning_grid"
        )

        if st.button("💾 Enregistrer le planning"):
            p["output_details_table"] = pd.DataFrame(grid_response["data"])
            st.success("✅ Planning enregistré avec succès")

        if st.button("✅ Valider le paramétrage"):
            df_final = p["output_details_table"]
            _, msg = pa.validate_parametrage()

            st.session_state["parametrage"] = {
                "entreprise": p["entreprise_choice"],
                "lot": p["model_choice"],
                "nombre_etages": p["nombre_etages"],
                "duree_stockage": p["duree_stockage"],
                "tarif_stockage": p["tarif_stockage"],
                "frais_palette": p["frais_palette"],
                "frais_livraison": p["frais_livraison"],
                "df1": p.get("output_table"),
                "param_details": df_final,
                "camions": p["choix_camions"],
                "conditionnements": p["choix_conditionnement"],
                "date_debut_prod": p["date_debut_prod"],
                "date_debut_term": p["date_debut_term"],
                "intervalle_par_etage": p["intervalle_par_etage"],
                "delai_livraison": p["delai_livraison"],
                "duree_prodmoyen_paretage": p["duree_prodmoyen_paretage"],
                "duree_termmoyen_paretage": p["duree_termmoyen_paretage"],
                "planning_indice": p.get("planning_indice", ""),
                "use_ccc": p.get("use_ccc", False)

            }

            st.success(msg)

# Onglet 2 : Données
elif menu == "Données":
    st.header("📊 Classification des bordereaux")

    params = st.session_state.get("parametrage", None)
    if not params:
        st.warning("⚠️ Veuillez d'abord compléter et valider l'onglet Paramétrage avant de continuer.")
        st.stop()

    
    #  CHOIX MODE A / MODE B
    
    mode = st.radio(
        "Méthode de travail :",
        ["Travailler directement dans l'application", "Télécharger et ré-uploader après modification"]
    )

    
    #  MODE B — Télécharger → Modifier dans Excel → Ré-uploader 
    
    if mode == "Télécharger et ré-uploader après modification":

        st.subheader("📥 Mode Édition Externe")

        st.info(
            "1️⃣ Déposez votre bordereau initial\n"
            "2️⃣ Générez un bordereau classé\n"
            "3️⃣ Téléchargez-le et modifiez-le dans Excel\n"
            "4️⃣ Ré-uploadez-le → il remplacera le tableau modifiable"
        )

        uploaded_init = st.file_uploader("📂 Bordereau Excel initial", type=["xlsx"], key="init_upload_B")

        if st.button("Créer le Bordereau classé (Mode B)") and uploaded_init:
            message, temp_path = do.process_file(params["lot"], uploaded_init)
            st.text_area("Message", value=message, height=80)

            if temp_path:
                with open(temp_path, "rb") as f:
                    st.download_button(
                        "📥 Télécharger pour modification dans Excel",
                        f,
                        file_name="bordereau_classé.xlsx"
                    )

        uploaded_modified = st.file_uploader(
            "📤 Ré-uploadez votre fichier modifié",
            type=["xlsx"],
            key="upload_modified_B"
        )

        if uploaded_modified:
            df_uploaded = pd.read_excel(uploaded_modified)
            st.session_state["bordereau_modifie"] = df_uploaded
            st.success("📌 Bordereau modifié chargé et prêt pour la génération finale.")


    
    # mode  A : Direct app 
    
    if mode == "Travailler directement dans l'application":

        # ------------------------------
        # 📂 Upload du fichier Excel source
        # ------------------------------
        uploaded_file = st.file_uploader("📂 Déposez le Bordereau Excel", type=["xlsx"])

        # ------------------------------
        # 🚀 Création du Bordereau classé
        # ------------------------------
        if st.button("Créer le Bordereau classé") and uploaded_file is not None:
            message, temp_path = do.process_file(params["lot"], uploaded_file)

            st.session_state["process_message"] = message
            st.text_area("Message", value=message, height=80)

            if temp_path:
                df_classed = pd.read_excel(temp_path)
                if "Taux de Confiance" in df_classed.columns:
                    df_classed["Taux de Confiance"] = df_classed["Taux de Confiance"].apply(lambda x: f"{x*100:.1f}%")

                st.session_state["bordereau_table"] = df_classed
                st.session_state["bordereau_modifie"] = df_classed.copy()

                with open(temp_path, "rb") as f:
                    st.download_button(
                        "📥 Télécharger le Bordereau classé (non regroupé)",
                        f,
                        file_name="bordereau_classé_non_regroupé.xlsx"
                    )
            else:
                st.error("Le bordereau n'a pas pu être généré. Corrige le fichier source puis réessaie.")

        # ------------------------------
        # 📝 Tableau modifiable (AgGrid)
        # ------------------------------
        if "bordereau_modifie" in st.session_state:
            st.subheader("📑 Bordereau classé (modifiable)")

            df_base = st.session_state["bordereau_modifie"]

            gb = GridOptionsBuilder.from_dataframe(df_base)
            gb.configure_pagination(enabled=True)
            gb.configure_default_column(editable=True, wrapText=True, autoHeight=True)

            if "Catégorie Prédite" in df_base.columns:
                lot_actuel = params["lot"]
                try:
                    if lot_actuel.upper() == "GLOBAL":
                        query = "SELECT DISTINCT nom FROM materiel"
                        df_cat = pd.read_sql_query(query, daba.conn)
                    else:
                        query = "SELECT DISTINCT nom FROM materiel WHERE lot = ?"
                        df_cat = pd.read_sql_query(query, daba.conn, params=(lot_actuel,))
                    categories_completes = (
                        df_cat["nom"].dropna().astype(str).unique().tolist()
                    )
                except:
                    categories_completes = (
                        df_base["Catégorie Prédite"].dropna().astype(str).unique().tolist()
                    )

                gb.configure_column(
                    "Catégorie Prédite",
                    editable=True,
                    cellEditor="agSelectCellEditor",
                    cellEditorParams={"values": sorted(categories_completes)}
                )

            grid_options = gb.build()
            grid_response = AgGrid(
                df_base,
                gridOptions=grid_options,
                data_return_mode='AS_INPUT',
                update_mode=GridUpdateMode.NO_UPDATE,
                fit_columns_on_grid_load=True,
                allow_unsafe_jscode=True,
                key="aggrid_bordereau"
            )

            if st.button("💾 Enregistrer le bordereau"):
                st.session_state["bordereau_modifie"] = pd.DataFrame(grid_response["data"])
                st.success("✅ Bordereau enregistré. Vous pouvez maintenant créer le fichier final.")


    if "bordereau_modifie" in st.session_state:
        st.subheader("📦 Génération du fichier final")

        if st.button("Valider et créer le fichier final"):
            msg, path = do.finalize_wrapper(
                st.session_state["bordereau_modifie"],
                params["lot"],
                params["nombre_etages"],
                params["duree_stockage"],
                params["tarif_stockage"],
                params["frais_palette"],
                params["frais_livraison"],
                params["df1"],
                params["param_details"],
                params["camions"],
                params["conditionnements"],
                params["entreprise"],
                params["planning_indice"],
                params["use_ccc"]
            )

            st.text_area("Message final", value=msg, height=80)

            if path:
                st.session_state["pilotage_file"] = path
                with open(path, "rb") as f:
                    st.download_button(
                        "📥 Télécharger le fichier final",
                        f,
                        file_name="Outil_Logistique.xlsm"
                    )

# onglet 2 : pilotag execel 

#onglet 2 pilotage excel : 
elif menu == "Pilotage Excel":
    st.header("Pilotage Excel")

    # === FEUILLE DONNEES ===
    st.subheader("Feuille Donnees")

    params = st.session_state.get("parametrage", None)
    if not params:
        st.warning("Veuillez d'abord completer l'onglet Parametrage.")
        st.stop()

    if "bordereau_modifie" not in st.session_state:
        st.warning("Veuillez d'abord completer l'onglet Donnees.")
        st.stop()

    planning = params.get("param_details")
    if planning is None or planning.empty:
        st.error("Planning detaille manquant ou vide.")
        st.stop()

    if "pilotage" not in st.session_state:
        st.session_state["pilotage"] = {}
    if "donnees_grid" not in st.session_state["pilotage"]:
        st.session_state["pilotage"]["donnees_grid"] = pex.build_donnees_grid(
            st.session_state["bordereau_modifie"],
            planning,
            params.get("lot", "")
        )

    df_donnees = st.session_state["pilotage"]["donnees_grid"]
    edited_donnees = st.data_editor(
        df_donnees,
        use_container_width=True,
        num_rows="dynamic",
        key="donnees_editor"
    )

    if st.button("Enregistrer Donnees", key="save_donnees"):
        st.session_state["pilotage"]["donnees_grid"] = edited_donnees.copy()
        st.success("Feuille Donnees mise a jour en session.")

    fichier_excel = st.session_state.get("pilotage_file")

    # === FEUILLE TABLEAU SOURCE ===
    st.subheader("Feuille Tableau Source")

    # Empêcher l'affichage automatique du tableau
    if "df_source" not in st.session_state:
        st.session_state["df_source"] = None
    if "df_source_modif" not in st.session_state:
        st.session_state["df_source_modif"] = None

    # --- Bouton creer tableau source ---
    if st.button("Creer Tableau Source", key="create_source"):
        try:
            st.session_state["df_source"] = pex.build_tableau_source(
                st.session_state["pilotage"]["donnees_grid"],
                params.get("lot", "")
            )
            st.session_state["df_source_modif"] = None
        except Exception as e:
            st.error(f"Erreur creation Tableau Source : {e}")

    # --- Bouton afficher tableau uniquement si df_source existe ---
    if st.session_state["df_source"] is not None or st.session_state["df_source_modif"] is not None:
        df_source = (
            st.session_state["df_source_modif"]
            if st.session_state["df_source_modif"] is not None
            else st.session_state["df_source"]
        )

        gb = GridOptionsBuilder.from_dataframe(df_source)
        gb.configure_default_column(editable=True)

        if "Type de camion requis" in df_source.columns:
            values = sorted(set(df_source["Type de camion requis"].dropna().unique()) | set(daba.liste_camions))
            gb.configure_column(
                "Type de camion requis",
                editable=True,
                cellEditor="agSelectCellEditor",
                cellEditorParams={"values": values}
            )

        if "Conditionnement" in df_source.columns:
            values = sorted(set(df_source["Conditionnement"].dropna().unique()) | set(daba.liste_conditionnement))
            gb.configure_column(
                "Conditionnement",
                editable=True,
                cellEditor="agSelectCellEditor",
                cellEditorParams={"values": values}
            )

        if "Utilisation d'un CCC" in df_source.columns:
            values = sorted(set(df_source["Utilisation d'un CCC"].dropna().unique()) | {"Oui", "Non"})
            gb.configure_column(
                "Utilisation d'un CCC",
                editable=True,
                cellEditor="agSelectCellEditor",
                cellEditorParams={"values": values}
            )

        grid_options = gb.build()
        grid_response = AgGrid(
            df_source,
            gridOptions=grid_options,
            update_on=["cellValueChanged"],
            fit_columns_on_grid_load=True,
            allow_unsafe_jscode=True,
            key="tableau_source"
        )

        st.session_state["df_source_modif"] = pd.DataFrame(grid_response["data"])

        if st.button("Enregistrer Tableau Source", key="save_source"):
            st.session_state["df_source"] = st.session_state["df_source_modif"].copy()
            st.success("Feuille Tableau Source mise a jour en session.")

    # === BILAN / LIVRABLE ===
    st.subheader("Creer le Livrable")
    if st.button("Creer Bilan & Livrable"):
        if not fichier_excel:
            st.info("Cette section utilise encore le fichier Excel.")
        else:
            msg = pex.lancer_macro_bilan(fichier_excel)
            if "Erreur" in msg:
                st.error(msg)
            else:
                st.success(msg)


# Onglet 3 : Dashboard :
elif menu == "Dashboard":
    import io
    st.header("📊 Études logistiques ")
    use_ccc = st.session_state.get("parametrage", {}).get("use_ccc", False)

    # =====================================================
    # 1) Chargement des dataframes internes (sans Excel)
    # =====================================================
    params = st.session_state.get("parametrage", None)
    if not params:
        st.warning("Veuillez d'abord compléter l'onglet Paramétrage.")
        st.stop()

    file_state = st.session_state.get("pilotage_file")
    file_bytes = file_state if isinstance(file_state, bytes) else None

    src = st.session_state.get("df_source_modif")
    if src is None:
        src = st.session_state.get("df_source")
    if src is None or src.empty:
        st.warning("Veuillez d'abord créer le Tableau Source dans l'onglet Pilotage Excel.")
        st.stop()

    planning = params.get("param_details")
    if planning is None or planning.empty:
        st.warning("Planning détaillé manquant ou vide.")
        st.stop()

    def _norm_col(name: str) -> str:
        value = "" if name is None else str(name)
        value = unicodedata.normalize("NFKD", value)
        value = "".join(c for c in value if not unicodedata.combining(c))
        return value.lower().strip()

    def _find_col(columns, target_name: str):
        target_norm = _norm_col(target_name)
        for col in columns:
            if _norm_col(col) == target_norm:
                return col
        for col in columns:
            if target_norm in _norm_col(col):
                return col
        return None

    def _find_col_any(columns, *names):
        for name in names:
            col = _find_col(columns, name)
            if col:
                return col
        return None

    def _is_true(val) -> bool:
        return str(val).strip().lower() in {"oui", "yes", "y", "1", "true"}

    def _safe_num(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    lot_col_src = _find_col(src.columns, "Lot")
    if not lot_col_src:
        src = src.copy()
        src["Lot"] = params.get("lot", "GLOBAL")
        lot_col_src = "Lot"

    lot_options = sorted(
        [str(x) for x in src[lot_col_src].dropna().unique() if str(x).strip() != ""]
    )
    if not lot_options:
        lot_options = [params.get("lot", "GLOBAL")]
        src[lot_col_src] = lot_options[0]
    lot_mode = st.radio(
        "Mode lots",
        ["Un lot", "Superposition"],
        horizontal=True,
        key="dashboard_lot_mode",
    )
    if lot_mode == "Un lot":
        lot_choice = st.selectbox(
            "Lot a afficher",
            lot_options,
            index=0,
            key="dashboard_lot_single",
        )
        selected_lots = [lot_choice] if lot_choice else []
    else:
        selected_lots = st.multiselect(
            "Lots a afficher",
            lot_options,
            default=lot_options,
            key="dashboard_lots",
        )
    if not selected_lots:
        st.info("Aucun lot sélectionné.")
        st.stop()

    src = src[src[lot_col_src].isin(selected_lots)].copy()

    # Matériel depuis la base SQL (avec filtre lots)
    try:
        materiel = daba.afficher_donnees("Matériel", "GLOBAL")
    except Exception:
        materiel = pd.DataFrame()
    if not materiel.empty:
        if "lot" in materiel.columns:
            materiel = materiel[materiel["lot"].isin(selected_lots)]
        if "nom" in materiel.columns and "Nom" not in materiel.columns:
            materiel = materiel.rename(columns={"nom": "Nom"})

    # Paramétrage au format "feuille"
    param_label = params.get("lot", "GLOBAL")
    param = pd.DataFrame(
        {
            "Lot": [
                "Nombre étage :",
                "Durée de stockage CCC (en mois)",
                "Tarif mois de stockage (en €)",
                "Frais supplémentaires/palette (en €)",
                "Frais de livraison par camion",
            ],
            param_label: [
                params.get("nombre_etages", ""),
                params.get("duree_stockage", ""),
                params.get("tarif_stockage", ""),
                params.get("frais_palette", ""),
                params.get("frais_livraison", ""),
            ],
        }
    )

    def _round_up(value, decimals=2):
        try:
            v = float(value)
        except Exception:
            return 0.0
        if v == 0:
            return 0.0
        factor = 10 ** decimals
        return math.ceil(v * factor) / factor

    def _build_bilan_graphique_df(df_source: pd.DataFrame, planning_df: pd.DataFrame):
        if df_source is None or df_source.empty:
            return pd.DataFrame(), {}

        # Colonnes tableau source
        col_etage = _find_col(df_source.columns, "Etage")
        col_zone = _find_col(df_source.columns, "Zone")
        col_lot = _find_col(df_source.columns, "Lot") or "Lot"
        col_phase = _find_col_any(
            df_source.columns,
            "Phase de traveaux",
            "Phase de travaux",
        )
        col_nom = _find_col_any(
            df_source.columns,
            "Nom de l'?l?ment",
            "Nom de l'element",
            "Nom",
        )
        col_qty = _find_col(df_source.columns, "Quantit?")
        col_pal = _find_col(df_source.columns, "Nombre palettes equivalent total")
        col_cam = _find_col_any(
            df_source.columns,
            "Nombre de camions n?cessaires",
            "Nombre de camions necessaires",
            "Nombre de Camions n?cessaires",
            "Nombre de Camions necessaires",
            "Nombre de camions nÃ©cessaires",
        )
        col_cam_plein = _find_col_any(
            df_source.columns,
            "Dont camions pleins",
            "Dont Camions pleins",
        )
        col_rempl = _find_col_any(
            df_source.columns,
            "Remplissage camion non plein",
            "Remplissage Camion non plein",
        )
        col_ccc = _find_col_any(
            df_source.columns,
            "Utilisation d'une CCC",
            "Utilisation d'un CCC",
            "Utilisation CCC",
        )
        col_type = _find_col(df_source.columns, "Type de camion requis")

        df = df_source.copy()
        if col_lot not in df.columns:
            df[col_lot] = "GLOBAL"
            col_lot = "Lot"
        if col_lot != "Lot":
            df = df.rename(columns={col_lot: "Lot"})
            col_lot = "Lot"

        if not col_etage or not col_zone:
            return pd.DataFrame(), {}

        def _not_blank(val):
            return val is not None and str(val).strip() != ""

        def _is_oui(val):
            return str(val).strip().lower() == "oui"

        def _is_non(val):
            return str(val).strip().lower() == "non"

        # Numeric conversion
        for c in [col_qty, col_pal, col_cam, col_cam_plein, col_rempl]:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        def _sum_mask(df_loc, mask, col):
            if not col or col not in df_loc.columns:
                return 0.0
            return float(df_loc.loc[mask, col].sum())

        # Colonnes planning
        p_etage = _find_col(planning_df.columns, "Etage") or _find_col(planning_df.columns, "?tage")
        p_zone = _find_col(planning_df.columns, "Zone")
        p_date_prod = _find_col(planning_df.columns, "Date d?but phase production")
        p_date_term = _find_col(planning_df.columns, "Date d?but phase terminaux")
        p_delai = _find_col(planning_df.columns, "D?lai de livraison avant travaux (jours)")
        p_duree_prod = _find_col(planning_df.columns, "Dur?e phase production")
        p_duree_term = _find_col(planning_df.columns, "Dur?e phase terminaux")

        planning_rows = []
        if all([p_etage, p_zone, p_date_prod, p_date_term, p_delai, p_duree_prod, p_duree_term]):
            for _, r in planning_df.iterrows():
                planning_rows.append({
                    "etage": r[p_etage],
                    "zone": r[p_zone],
                    "date_prod": r[p_date_prod],
                    "date_term": r[p_date_term],
                    "delai": r[p_delai],
                    "duree_prod": r[p_duree_prod],
                    "duree_term": r[p_duree_term],
                })

        lot_values = selected_lots if selected_lots else df[col_lot].dropna().astype(str).unique().tolist()

        palettes_rows = []
        camions_rows = []
        flux_rows = []
        cam_type_rows = []
        cam_type_ccc_rows = []
        mat_ccc_rows = []

        for lot in lot_values:
            df_lot = df[df[col_lot].astype(str) == str(lot)].copy()
            if df_lot.empty:
                continue

            # Zones order
            if planning_rows:
                zone_order = [(r["etage"], r["zone"], r) for r in planning_rows]
            else:
                zone_order = [
                    (r[col_etage], r[col_zone], None)
                    for _, r in df_lot[[col_etage, col_zone]].drop_duplicates().iterrows()
                ]

            # Palettes + camions par zone
            for etage, zone, prow in zone_order:
                mask_zone = (df_lot[col_etage] == etage) & (df_lot[col_zone] == zone)

                mask_util = mask_zone & df_lot[col_ccc].apply(_not_blank) if col_ccc else mask_zone
                if col_phase and col_phase in df_lot.columns:
                    mask_prod = mask_zone & (df_lot[col_phase] == "Production")
                    mask_term = mask_zone & (df_lot[col_phase] == "Terminaux")
                else:
                    mask_prod = mask_zone
                    mask_term = mask_zone & False

                pal_prod = _sum_mask(df_lot, mask_prod, col_pal)
                pal_term = _sum_mask(df_lot, mask_term, col_pal)

                palettes_rows.append({
                    "Lot": lot,
                    "?tage - Zone": f"{etage} - {zone}",
                    "Production": pal_prod,
                    "Terminaux": pal_term,
                })

                # Camions sans CCC
                cam_prod_sans = _sum_mask(df_lot, mask_prod, col_cam)
                cam_term_sans = _sum_mask(df_lot, mask_term, col_cam)

                denom_sans = _sum_mask(df_lot, mask_util, col_cam)
                numer_sans = _sum_mask(df_lot, mask_util, col_rempl) + _sum_mask(df_lot, mask_util, col_cam_plein)
                rempl_sans = _round_up(numer_sans / denom_sans, 2) if denom_sans else 0.0

                # Stock CCC lignes
                if col_nom and col_nom in df_lot.columns:
                    stock_prod = df_lot[mask_zone & df_lot[col_nom].astype(str).str.strip().eq("Stock CCC Production")]
                    stock_term = df_lot[mask_zone & df_lot[col_nom].astype(str).str.strip().eq("Stock CCC Terminaux")]
                else:
                    stock_prod = df_lot.iloc[0:0]
                    stock_term = df_lot.iloc[0:0]

                stock_prod_m = _sum_mask(stock_prod, slice(None), col_cam)
                stock_term_m = _sum_mask(stock_term, slice(None), col_cam)
                stock_prod_n = _sum_mask(stock_prod, slice(None), col_cam_plein)
                stock_term_n = _sum_mask(stock_term, slice(None), col_cam_plein)
                stock_prod_o = _sum_mask(stock_prod, slice(None), col_rempl)
                stock_term_o = _sum_mask(stock_term, slice(None), col_rempl)

                mask_non = mask_zone & df_lot[col_ccc].apply(_is_non) if col_ccc else mask_zone
                if col_phase and col_phase in df_lot.columns:
                    mask_prod_non = mask_non & (df_lot[col_phase] == "Production")
                    mask_term_non = mask_non & (df_lot[col_phase] == "Terminaux")
                else:
                    mask_prod_non = mask_non
                    mask_term_non = mask_non & False

                cam_prod_avec = _sum_mask(df_lot, mask_prod_non, col_cam) + stock_prod_m
                cam_term_avec = _sum_mask(df_lot, mask_term_non, col_cam) + stock_term_m

                denom_avec = _sum_mask(df_lot, mask_non, col_cam) + stock_prod_m + stock_term_m
                numer_avec = _sum_mask(df_lot, mask_non, col_rempl) + _sum_mask(df_lot, mask_non, col_cam_plein)
                numer_avec += stock_prod_o + stock_term_o + stock_prod_n + stock_term_n
                rempl_avec = _round_up(numer_avec / denom_avec, 2) if denom_avec else 0.0

                camions_rows.append({
                    "Lot": lot,
                    "?tage - Zone": f"{etage} - {zone}",
                    "Camions Production sans CCC": cam_prod_sans,
                    "Camions Terminaux sans CCC": cam_term_sans,
                    "Camions Production avec CCC": cam_prod_avec,
                    "Camions Terminaux avec CCC": cam_term_avec,
                    "Remplissage camions sans CCC": rempl_sans * 100,
                    "Remplissage camions avec CCC": rempl_avec * 100,
                })

                # Flux mensuel (r?gle 50% + reste)
                if prow:
                    try:
                        delai = float(prow["delai"]) if prow["delai"] not in ("", None) else 0.0
                    except Exception:
                        delai = 0.0

                    for phase_name, volume, cam_sans, cam_avec, date_key, duree_key in [
                        ("Production", pal_prod, cam_prod_sans, cam_prod_avec, "date_prod", "duree_prod"),
                        ("Terminaux", pal_term, cam_term_sans, cam_term_avec, "date_term", "duree_term"),
                    ]:
                        date_raw = prow.get(date_key)
                        duree_raw = prow.get(duree_key)
                        if date_raw in (None, "") or duree_raw in (None, ""):
                            continue
                        try:
                            date_base = pd.to_datetime(date_raw, dayfirst=True, errors="coerce")
                        except Exception:
                            date_base = None
                        if pd.isna(date_base):
                            continue

                        try:
                            duree = float(duree_raw)
                        except Exception:
                            duree = 0.0
                        if duree <= 0:
                            continue

                        date_debut = date_base - pd.to_timedelta(delai, unit="D")
                        nb_mois = int(math.ceil(duree / 30.0))
                        if nb_mois <= 0:
                            continue

                        if nb_mois == 1:
                            repart_reste = 1.0
                        else:
                            repart_reste = 0.5 / (nb_mois - 1)

                        for k in range(nb_mois):
                            if nb_mois == 1:
                                repart = 1.0
                            elif k == 0:
                                repart = 0.5
                            else:
                                repart = repart_reste

                            mois = (date_debut + pd.DateOffset(months=k)).to_period("M").strftime("%Y-%m")
                            flux_rows.append({
                                "Lot": lot,
                                "Mois": mois,
                                "Volume (nombre de palettes ?quivalentes)": volume * repart,
                                "Nombre de Camions": cam_sans * repart,
                                "Nombre de Camions CCC": cam_avec * repart,
                            })

            # Typologie camions
            if col_type and col_cam:
                mask_cam = df_lot[col_ccc].apply(_not_blank) if col_ccc else pd.Series(True, index=df_lot.index)
                cam_group = (
                    df_lot[mask_cam]
                    .groupby([col_etage, col_zone, col_type], as_index=False)[col_cam]
                    .sum()
                )
                for _, r in cam_group.iterrows():
                    cam_type_rows.append({
                        "Lot": lot,
                        "?tage": r[col_etage],
                        "Zone": r[col_zone],
                        "Type de Camion": r[col_type],
                        "Nombre de Camions": r[col_cam],
                    })

                mask_cam_ccc = df_lot[col_ccc].apply(lambda x: str(x).strip().lower() != "oui") if col_ccc else pd.Series(True, index=df_lot.index)
                cam_group_ccc = (
                    df_lot[mask_cam_ccc]
                    .groupby([col_type], as_index=False)[col_cam]
                    .sum()
                )
                for _, r in cam_group_ccc.iterrows():
                    cam_type_ccc_rows.append({
                        "Lot": lot,
                        "?tage": "",
                        "Type de Camion": r[col_type],
                        "Nombre de Camions avec CCC": r[col_cam],
                    })

            # Mat?riel CCC
            if col_nom and col_qty and col_ccc:
                mat_ccc = (
                    df_lot[df_lot[col_ccc].apply(_is_oui)]
                    .groupby(col_nom, as_index=False)[col_qty]
                    .sum()
                )
                for _, r in mat_ccc.iterrows():
                    mat_ccc_rows.append({
                        "Lot": lot,
                        "Mat?riel CCC": r[col_nom],
                        "Nombre de mat?riels CCC": r[col_qty],
                    })

        # Flux mensuel agr?g? par mois/lot
        if flux_rows:
            flux_df = (
                pd.DataFrame(flux_rows)
                .groupby(["Lot", "Mois"], as_index=False)
                .sum()
            )
        else:
            flux_df = pd.DataFrame(columns=["Lot", "Mois", "Volume (nombre de palettes ?quivalentes)", "Nombre de Camions", "Nombre de Camions CCC"])

        palettes_df = pd.DataFrame(palettes_rows)
        camions_df = pd.DataFrame(camions_rows)
        cam_type_df = pd.DataFrame(cam_type_rows)
        cam_type_ccc_df = pd.DataFrame(cam_type_ccc_rows)
        mat_ccc_df = pd.DataFrame(mat_ccc_rows)

        # Mat?riel complet depuis bordereau (optionnel)
        mat_complete_df = pd.DataFrame(columns=["Mat?riel complet", "Nombre total de mat?riels"])
        df_bord = st.session_state.get("bordereau_modifie")
        if isinstance(df_bord, pd.DataFrame) and not df_bord.empty:
            col_cat = _find_col(df_bord.columns, "Categorie Predite") or _find_col(df_bord.columns, "Cat?gorie Pr?dite")
            col_q = _find_col(df_bord.columns, "Quantite") or _find_col(df_bord.columns, "Quantit?")
            if col_cat and col_q:
                mat_complete_df = (
                    df_bord.groupby(col_cat, as_index=False)[col_q]
                    .sum()
                    .rename(columns={col_cat: "Mat?riel complet", col_q: "Nombre total de mat?riels"})
                )

        bg = pd.concat(
            [
                palettes_df,
                camions_df,
                flux_df,
                cam_type_df,
                cam_type_ccc_df,
                mat_ccc_df,
                mat_complete_df,
            ],
            ignore_index=True,
        )

        extras = {
            "mat_v0": mat_complete_df.rename(columns={"Mat?riel complet": "Mat?riel complet", "Nombre total de mat?riels": "Nombre total de mat?riels"}),
            "mat_ccc": mat_ccc_df.rename(columns={"Mat?riel CCC": "Mat?riel CCC", "Nombre de mat?riels CCC": "Nombre de mat?riels CCC"}),
            "cam_type_v0": cam_type_df,
            "cam_type_ccc": cam_type_ccc_df,
        }
        return bg, extras

    bg, extras = _build_bilan_graphique_df(src, planning)
    if bg is None or bg.empty:
        st.error("Données insuffisantes pour générer le dashboard.")
        st.stop()

    def _lot_color(df, default_color=None):
        if isinstance(df, pd.DataFrame) and "Lot" in df.columns and df["Lot"].nunique() > 1:
            return "Lot"
        return default_color

    # =====================================================
    # 3) Pipelines unifiés
    #    - pipeline_sans_ccc : traitement V0
    #    - pipeline_avec_ccc : traitement V1
    # =====================================================

    def pipeline_sans_ccc(bg_df: pd.DataFrame) -> dict:
        col_etage_zone = _find_col(bg_df.columns, "Étage - Zone")

        # Palettes par étage / zone
        try:
            if not col_etage_zone:
                raise KeyError("Étage - Zone")
            palette_cols = [col_etage_zone, "Production", "Terminaux"]
            if "Lot" in bg_df.columns:
                palette_cols = ["Lot"] + palette_cols
            palettes_zone = (
                bg_df[palette_cols]
                .dropna(subset=[col_etage_zone])
                .copy()
            )
            palettes_zone = palettes_zone.rename(columns={col_etage_zone: "Étage - Zone"})
            palettes_zone["Production"] = palettes_zone["Production"].fillna(0)
            palettes_zone["Terminaux"] = palettes_zone["Terminaux"].fillna(0)
            palettes_zone["Palettes"] = (
                palettes_zone["Production"] + palettes_zone["Terminaux"]
            )
        except Exception:
            palettes_zone = pd.DataFrame(columns=["Étage - Zone", "Palettes"])
        total_palettes = float(palettes_zone["Palettes"].sum()) if not palettes_zone.empty else 0.0

        # Flux mensuel palettes
        if {"Mois", "Volume (nombre de palettes équivalentes)"}.issubset(bg_df.columns):
            flux_palettes = (
                bg_df[[c for c in ["Lot", "Mois", "Volume (nombre de palettes équivalentes)"] if c in bg_df.columns]]
                .dropna(subset=["Mois"])
                .groupby(["Lot", "Mois"] if "Lot" in bg_df.columns else ["Mois"], as_index=False)
                .sum()
            )
            if not flux_palettes.empty:
                idx_pic_pal = flux_palettes["Volume (nombre de palettes équivalentes)"].idxmax()
                mois_pic_palettes = flux_palettes.loc[idx_pic_pal, "Mois"]
                pic_palettes = float(
                    flux_palettes.loc[idx_pic_pal, "Volume (nombre de palettes équivalentes)"]
                )
            else:
                mois_pic_palettes = ""
                pic_palettes = 0.0
        else:
            flux_palettes = pd.DataFrame(columns=["Mois", "Volume (nombre de palettes équivalentes)"])
            mois_pic_palettes = ""
            pic_palettes = 0.0

        # Camions par étage / zone (sans CCC)
        if col_etage_zone and {
            "Camions Production sans CCC",
            "Camions Terminaux sans CCC",
        }.issubset(bg_df.columns):
            camions_zone = (
                bg_df[[c for c in ["Lot", col_etage_zone, "Camions Production sans CCC", "Camions Terminaux sans CCC"] if c in bg_df.columns]]
                .dropna(subset=[col_etage_zone])
                .copy()
            )
            camions_zone = camions_zone.rename(columns={col_etage_zone: "Étage - Zone"})
            camions_zone["Camions Production sans CCC"] = camions_zone[
                "Camions Production sans CCC"
            ].fillna(0)
            camions_zone["Camions Terminaux sans CCC"] = camions_zone[
                "Camions Terminaux sans CCC"
            ].fillna(0)
            camions_zone["Camions totaux"] = (
                camions_zone["Camions Production sans CCC"]
                + camions_zone["Camions Terminaux sans CCC"]
            )
        else:
            camions_zone = pd.DataFrame(columns=["Étage - Zone", "Camions totaux"])

        # Total camions
        if "Nombre de Camions" in bg_df.columns:
            total_camions = float(bg_df["Nombre de Camions"].fillna(0).sum())
        else:
            total_camions = float(camions_zone["Camions totaux"].sum()) if not camions_zone.empty else 0.0

        # Remplissage par zone + remplissage moyen (sans CCC)
        if col_etage_zone and "Remplissage camions sans CCC" in bg_df.columns:
            rempl_zone = (
                bg_df[[c for c in ["Lot", col_etage_zone, "Remplissage camions sans CCC"] if c in bg_df.columns]]
                .dropna(subset=[col_etage_zone])
                .copy()
            )
            rempl_zone = rempl_zone.rename(columns={col_etage_zone: "Étage - Zone"})
            rempl_zone["Remplissage (%)"] = (
                rempl_zone["Remplissage camions sans CCC"].fillna(0)
            )
            rempl_brut = bg_df["Remplissage camions sans CCC"].dropna()
            rempl_moyen = float(rempl_brut.mean()) if not rempl_brut.empty else 0.0
        else:
            rempl_zone = pd.DataFrame(columns=["Étage - Zone", "Remplissage (%)"])
            rempl_moyen = 0.0

        # Flux mensuel camions
        if "Nombre de Camions" in bg_df.columns and "Mois" in bg_df.columns:
            flux_camions = (
                bg_df[[c for c in ["Lot", "Mois", "Nombre de Camions"] if c in bg_df.columns]]
                .dropna(subset=["Mois"])
                .groupby(["Lot", "Mois"] if "Lot" in bg_df.columns else ["Mois"], as_index=False)
                .sum()
            )
            if not flux_camions.empty:
                idx_pic_cam = flux_camions["Nombre de Camions"].idxmax()
                mois_pic_camions = flux_camions.loc[idx_pic_cam, "Mois"]
                pic_camions = float(flux_camions.loc[idx_pic_cam, "Nombre de Camions"])
            else:
                mois_pic_camions = ""
                pic_camions = 0.0
        else:
            flux_camions = pd.DataFrame(columns=["Mois", "Nombre de Camions"])
            mois_pic_camions = ""
            pic_camions = 0.0

        # Camions par type
        cols = list(bg_df.columns)
        col_etage_type = None
        for c in cols:
            if c.startswith("Étage") and c != "Étage - Zone":
                col_etage_type = c
                break
        if col_etage_type:
            idx_etage_type = cols.index(col_etage_type)
            col_zone_type = cols[idx_etage_type + 1] if idx_etage_type + 1 < len(cols) else None
            col_type_camion = cols[idx_etage_type + 2] if idx_etage_type + 2 < len(cols) else None
            col_nb_camions_type = cols[idx_etage_type + 3] if idx_etage_type + 3 < len(cols) else None
            if col_type_camion and col_nb_camions_type:
                camions_type = (
                    bg_df[[col_type_camion, col_nb_camions_type]]
                    .dropna(subset=[col_type_camion])
                    .copy()
                )
                camions_type = (
                    camions_type
                    .groupby(col_type_camion, as_index=False)[col_nb_camions_type]
                    .sum()
                )
                camions_type.rename(
                    columns={
                        col_type_camion: "Type de Camion",
                        col_nb_camions_type: "Nombre de Camions",
                    },
                    inplace=True,
                )
            else:
                camions_type = pd.DataFrame(columns=["Type de Camion", "Nombre de Camions"])
        else:
            camions_type = pd.DataFrame(columns=["Type de Camion", "Nombre de Camions"])

        return {
            "palettes_zone": palettes_zone,
            "total_palettes": total_palettes,
            "flux_palettes": flux_palettes,
            "mois_pic_palettes": mois_pic_palettes,
            "pic_palettes": pic_palettes,
            "camions_zone": camions_zone,
            "total_camions": total_camions,
            "flux_camions": flux_camions,
            "mois_pic_camions": mois_pic_camions,
            "pic_camions": pic_camions,
            "rempl_zone": rempl_zone,
            "rempl_moyen": rempl_moyen,
            "camions_type": camions_type,
        }

    def pipeline_avec_ccc(bg_df: pd.DataFrame) -> dict:
        # On réutilise la partie palettes du pipeline sans CCC
        base = pipeline_sans_ccc(bg_df)
        col_etage_zone = _find_col(bg_df.columns, "Étage - Zone")

        # Flux mensuel camions CCC
        if "Nombre de Camions CCC" in bg_df.columns and "Mois" in bg_df.columns:
            flux_camions_ccc = (
                bg_df[[c for c in ["Lot", "Mois", "Nombre de Camions CCC"] if c in bg_df.columns]]
                .dropna(subset=["Mois"])
                .groupby(["Lot", "Mois"] if "Lot" in bg_df.columns else ["Mois"], as_index=False)
                .sum()
            )
            if not flux_camions_ccc.empty:
                idx_pic_cam_v1 = flux_camions_ccc["Nombre de Camions CCC"].idxmax()
                mois_pic_camions_v1 = flux_camions_ccc.loc[idx_pic_cam_v1, "Mois"]
                pic_camions_v1 = float(
                    flux_camions_ccc.loc[idx_pic_cam_v1, "Nombre de Camions CCC"]
                )
            else:
                mois_pic_camions_v1 = ""
                pic_camions_v1 = 0.0
        else:
            flux_camions_ccc = pd.DataFrame(columns=["Mois", "Nombre de Camions CCC"])
            mois_pic_camions_v1 = ""
            pic_camions_v1 = 0.0

        # Indicateurs globaux CCC
        if "Nombre de Camions CCC" in bg_df.columns:
            total_camions_ccc = float(bg_df["Nombre de Camions CCC"].fillna(0).sum())
        else:
            total_camions_ccc = 0.0

        if col_etage_zone and "Remplissage camions avec CCC" in bg_df.columns:
            rempl_brut_ccc = bg_df["Remplissage camions avec CCC"].dropna()
            rempl_moyen_ccc = float(rempl_brut_ccc.mean()) if not rempl_brut_ccc.empty else 0.0
            rempl_zone_ccc = (
                bg_df[[c for c in ["Lot", col_etage_zone, "Remplissage camions avec CCC"] if c in bg_df.columns]]
                .dropna(subset=[col_etage_zone])
                .copy()
            )
            rempl_zone_ccc = rempl_zone_ccc.rename(columns={col_etage_zone: "Étage - Zone"})
            rempl_zone_ccc["Remplissage (%)"] = (
                rempl_zone_ccc["Remplissage camions avec CCC"]
            )
        else:
            rempl_moyen_ccc = 0.0
            rempl_zone_ccc = pd.DataFrame(columns=["Étage - Zone", "Remplissage (%)"])

        # Camions par étage / zone CCC
        if col_etage_zone and {
            "Camions Production avec CCC",
            "Camions Terminaux avec CCC",
        }.issubset(bg_df.columns):
            camions_zone_ccc = (
                bg_df[
                    [c for c in [
                        "Lot",
                        col_etage_zone,
                        "Camions Production avec CCC",
                        "Camions Terminaux avec CCC",
                    ] if c in bg_df.columns]
                ]
                .dropna(subset=[col_etage_zone])
                .copy()
            )
            camions_zone_ccc = camions_zone_ccc.rename(columns={col_etage_zone: "Étage - Zone"})
            camions_zone_ccc["Total CCC"] = (
                camions_zone_ccc["Camions Production avec CCC"].fillna(0)
                + camions_zone_ccc["Camions Terminaux avec CCC"].fillna(0)
            )
        else:
            camions_zone_ccc = pd.DataFrame(columns=["Étage - Zone", "Total CCC"])

        # On renvoie des noms neutres pour la comparaison
        return {
            "palettes_zone": base["palettes_zone"],
            "total_palettes": base["total_palettes"],
            "flux_palettes": base["flux_palettes"],
            "mois_pic_palettes": base["mois_pic_palettes"],
            "pic_palettes": base["pic_palettes"],
            "camions_zone": camions_zone_ccc.rename(columns={"Total CCC": "Camions"}) if not camions_zone_ccc.empty else pd.DataFrame(columns=["Étage - Zone", "Camions"]),
            "total_camions": total_camions_ccc,
            "flux_camions": flux_camions_ccc.rename(columns={"Nombre de Camions CCC": "Camions"}) if not flux_camions_ccc.empty else pd.DataFrame(columns=["Mois", "Camions"]),
            "mois_pic_camions": mois_pic_camions_v1,
            "pic_camions": pic_camions_v1,
            "rempl_zone": rempl_zone_ccc,
            "rempl_moyen": rempl_moyen_ccc,
            "camions_type": base["camions_type"],  # typologie identique
        }

    # =====================================================
    # 3 bis) Préparation des données de base via pipelines
    # =====================================================

    metrics_v0 = pipeline_sans_ccc(bg)
    palettes_zone = metrics_v0["palettes_zone"]
    total_palettes = metrics_v0["total_palettes"]
    flux_palettes = metrics_v0["flux_palettes"]
    mois_pic_palettes = metrics_v0["mois_pic_palettes"]
    pic_palettes = metrics_v0["pic_palettes"]
    camions_zone = metrics_v0["camions_zone"]
    total_camions = metrics_v0["total_camions"]
    flux_camions = metrics_v0["flux_camions"]
    mois_pic_camions = metrics_v0["mois_pic_camions"]
    pic_camions = metrics_v0["pic_camions"]
    rempl_zone = metrics_v0["rempl_zone"]
    rempl_moyen = metrics_v0["rempl_moyen"]
    camions_type = metrics_v0["camions_type"]

    metrics_v1 = pipeline_avec_ccc(bg)
    total_camions_ccc = metrics_v1["total_camions"]
    flux_camions_ccc = metrics_v1["flux_camions"].rename(columns={"Camions": "Nombre de Camions CCC"}) if not metrics_v1["flux_camions"].empty else pd.DataFrame(columns=["Mois", "Nombre de Camions CCC"])
    mois_pic_camions_v1 = metrics_v1["mois_pic_camions"]
    pic_camions_v1 = metrics_v1["pic_camions"]
    rempl_zone_ccc = metrics_v1["rempl_zone"]
    rempl_moyen_ccc = metrics_v1["rempl_moyen"]
    camions_zone_ccc = metrics_v1["camions_zone"].rename(columns={"Camions": "Total CCC"}) if not metrics_v1["camions_zone"].empty else pd.DataFrame(columns=["Étage - Zone", "Total CCC"])

    # ---------- Hypothèses générales ----------
    lot_col = param.columns[1]  # même logique que ta macro
    try:
        nb_etages = param.loc[param["Lot"] == "Nombre étage :", lot_col].iloc[0]
    except Exception:
        nb_etages = ""
    familles_identifiees = (
        materiel["Nom"].dropna().astype(str).sort_values().unique().tolist()
        if "Nom" in materiel.columns
        else []
    )

    # DPGF + indice planning depuis la session
    dpgf_date = st.session_state.get("dpgf_date", "")
    dpgf_date = st.text_input("Date du DPGF :", value=dpgf_date)
    st.session_state["dpgf_date"] = dpgf_date
    planning_indice = st.session_state.get("parametrage", {}).get("planning_indice", "")

    # =====================================================
    # 4) Gestion des variantes (initialisation)
    # =====================================================
    if "variants" not in st.session_state:
        st.session_state["variants"] = {}      # {"V2": {"with_ccc": bool, "bytes": ...}}
    if "variant_counter" not in st.session_state:
        st.session_state["variant_counter"] = 2

    # 4) Onglets V0 / V1 / Variantes / Comparatif
    tabs = []

    if use_ccc:
        tabs.append("V1")
    else:
        tabs.append("V0")

    tabs.extend(["Variantes", "Comparatif"])

    tab_objects = st.tabs(tabs)
    main_tab = tab_objects[0]
    tab_var = tab_objects[1]
    tab_comp = tab_objects[2]



    with main_tab:
        if use_ccc:


            try:
                src_v1 = src.copy()
            except Exception:
                src_v1 = pd.DataFrame()

            # ---------- 3 onglets internes : Hypothèses / Palettes / Camions ----------
            ong_hyp_v1, ong_pal_v1, ong_cam_v1 = st.tabs(
                ["📘 Hypothèses", "📦 Palettes", "🚚 Camions"]
            )

            # ======================================================
            # 📘 ONGLET HYPOTHÈSES V1
            # ======================================================
            with ong_hyp_v1:
                st.markdown("### 📘 Hypothèses")

                # ================= DPGF + PIC =================
                st.markdown("#### Document de source")

                if dpgf_date and planning_indice:
                    default_dpgf_v1 = f"DPGF du {dpgf_date} – Indice {planning_indice}"
                elif dpgf_date:
                    default_dpgf_v1 = f"DPGF du {dpgf_date}"
                elif planning_indice:
                    default_dpgf_v1 = f"Indice {planning_indice}"
                else:
                    default_dpgf_v1 = ""

                dpgf_txt_v1 = st.text_area(
                    "DPGF + Indice :",
                    value=default_dpgf_v1,
                    key="dpgf_v1",
                    placeholder="DPGF du … – Indice …",
                )

                pic_file_v1 = st.file_uploader(
                    "Veuillez joindre le fichier PIC", key="pic_v1"
                )

                st.markdown("---")

                # ================= Hypothèse planning ================
                st.markdown("#### Hypothèse planning")
                st.markdown(f"- Planning indice : **{planning_indice or '…'}**")

                planning_hyp_v1 = st.text_area(
                    "Hypothèses prises pour le planning :",
                    key="planning_hyp_v1",
                    placeholder="Décrire ici les hypothèses utilisées pour le planning…",
                )

                st.markdown("---")

                # ================= Hypothèses de l’étude ================
                st.markdown("#### Hypothèses de l’étude")
                st.markdown("- regroupement du matériel en grandes catégories")
                st.markdown("- conversion des conditionnements en équivalent palette")
                st.markdown("- 2 phases de travaux par étage")

                st.markdown("---")

                # ================= Paramètres CCC ================
                st.markdown("#### Paramètres CCC")

                def _val_param_ccc(libel):
                    try:
                        return param.loc[param["Lot"] == libel, lot_col].iloc[0]
                    except:
                        return ""

                duree_ccc = _val_param_ccc("Durée de stockage CCC (en mois)")
                tarif_mois = _val_param_ccc("Tarif mois de stockage (en €)")
                frais_sup = _val_param_ccc("Frais supplémentaires/palette (en €)")
                frais_liv = _val_param_ccc("Frais de livraison par camion")

                st.markdown(f"- Durée stockage : **{duree_ccc} mois**")
                st.markdown(f"- Tarif de stockage : **{tarif_mois} €/mois**")
                st.markdown(f"- Frais supplémentaires/palette : **{frais_sup} €**")
                st.markdown(f"- Frais de livraison : **{frais_liv} €**")

                st.markdown("---")

                # ================= Familles CCC (Tableau Source + BG) ================
                st.markdown("#### Hypothèse de base déportée par famille")

                if (
                    {"Nom de l'élément", "Utilisation d'une CCC"}.issubset(src_v1.columns)
                    and {"Matériel CCC", "Nombre de matériels CCC"}.issubset(bg.columns)
                ):
                    df_src = src_v1[["Nom de l'élément", "Utilisation d'une CCC"]].dropna()
                    df_src["use_ccc"] = df_src["Utilisation d'une CCC"].astype(str).str.lower().isin(
                        ["oui", "yes", "y", "1"]
                    )

                    df_yes = (
                        df_src.groupby("Nom de l'élément", as_index=False)["use_ccc"]
                        .any()
                        .rename(columns={"Nom de l'élément": "Famille"})
                    )

                    df_qty = (
                        bg[["Matériel CCC", "Nombre de matériels CCC"]]
                        .dropna(subset=["Matériel CCC"])
                        .groupby("Matériel CCC", as_index=False)["Nombre de matériels CCC"]
                        .sum()
                        .rename(columns={"Matériel CCC": "Famille", "Nombre de matériels CCC": "Quantité"})
                    )

                    df_merge = pd.merge(df_yes, df_qty, on="Famille", how="left")
                    df_merge["Quantité"] = df_merge["Quantité"].fillna(0)
                    df_merge["Stocké en CCC ?"] = df_merge["use_ccc"].apply(lambda x: "✔️" if x else "❌")

                    st.dataframe(df_merge[["Famille", "Stocké en CCC ?", "Quantité"]], use_container_width=True)

                else:
                    st.info("Colonnes nécessaires introuvables dans Tableau Source / BG")

            # ======================================================
            # 📦 ONGLET PALETTES (V1)
            # ======================================================
            with ong_pal_v1:

                st.markdown("### 📦 Palettes – V1 (CCC)")

                colA, colB = st.columns(2)
                with colA:
                    st.metric(
                        "Palettes équivalentes totales (identiques V0/V1)",
                        f"{total_palettes:,.0f}".replace(",", " "),
                    )
                with colB:
                    st.metric(
                        "Surface totale (m²)",
                        f"{(total_palettes * 0.96):,.0f}".replace(",", " "),
                    )

                c1, c2 = st.columns(2)

                # -------- Matériaux stockés en CCC --------
                with c1:
                    st.markdown("#### Matériaux stockés en CCC")

                    df_v1_mat = extras.get("mat_ccc", pd.DataFrame())
                    if not df_v1_mat.empty:
                        color_arg = _lot_color(df_v1_mat, None)
                        fig_mat_v1 = px.bar(
                            df_v1_mat,
                            x="Nombre de matériels CCC",
                            y="Matériel CCC",
                            color=color_arg,
                            orientation="h",
                            title="Répartition des matériaux stockés en CCC",
                        )
                        st.plotly_chart(fig_mat_v1, key="bar_mat_v1", use_container_width=True)
                    else:
                        st.info("Colonnes Matériel CCC absentes")

                # -------- Flux palettes (identique V0) --------
                with c2:
                    st.markdown("#### Flux mensuel de palettes (PIC identique)")

                    flux_palettes_plot = flux_palettes.copy()
                    if not flux_palettes_plot.empty:
                        flux_palettes_plot["Mois"] = flux_palettes_plot["Mois"].astype(str)

                    color_arg = _lot_color(flux_palettes_plot, None)
                    fig_flux_pal_v1 = px.area(
                        flux_palettes_plot,
                        x="Mois",
                        y="Volume (nombre de palettes équivalentes)",
                        color=color_arg,
                    )

                    # PIC palettes V1
                    if pic_palettes > 0:
                        mois_pic_palettes_str = str(mois_pic_palettes)

                        fig_flux_pal_v1.add_vline(
                            x=mois_pic_palettes_str,
                            line_dash="dot",
                            line_color="red",
                        )
                        fig_flux_pal_v1.add_hline(
                            y=pic_palettes,
                            line_dash="dot",
                            line_color="red",
                        )
                        fig_flux_pal_v1.add_scatter(
                            x=[mois_pic_palettes_str],
                            y=[pic_palettes],
                            mode="markers",
                            marker=dict(color="red", size=10),
                            name="Pic",
                        )
                        fig_flux_pal_v1.add_annotation(
                            x=mois_pic_palettes_str,
                            y=pic_palettes,
                            text=f"Pic : {pic_palettes:.0f} palettes ({mois_pic_palettes_str})",
                            showarrow=True,
                            arrowhead=2,
                            ax=0,
                            ay=-40,
                            font=dict(color="red"),
                        )

                    fig_flux_pal_v1.update_layout(margin=dict(l=10, r=10, t=30, b=40))

                    st.plotly_chart(
                        fig_flux_pal_v1,
                        use_container_width=True,
                        key="flux_palettes_v1",
                    )


                # -------- Palettes par étage --------
                st.markdown("#### Répartition des palettes par étage / zone")
                color_arg = _lot_color(palettes_zone, "Palettes")
                fig_pal_v1 = px.bar(
                    palettes_zone,
                    x="Étage - Zone",
                    y="Palettes",
                    color=color_arg,
                    barmode="group" if color_arg == "Lot" else None,
                )
                st.plotly_chart(fig_pal_v1, key="palettes_zone_v1", use_container_width=True)

            # ======================================================
            # 🚚 ONGLET CAMIONS (V1)
            # ======================================================
            with ong_cam_v1:

                st.markdown("### 🚚 Camions avec CCC")

                colA, colB = st.columns(2)
                colA.metric("Nombre total de camions (CCC)", f"{total_camions_ccc:,.0f}")
                colB.metric("Remplissage moyen (CCC)", f"{rempl_moyen_ccc:.1f} %")

                c1, c2 = st.columns(2)

                # -------- Camions par zone --------
                with c1:
                    st.markdown("#### Camions par étage (CCC)")
                    if not camions_zone_ccc.empty:
                        color_arg = _lot_color(camions_zone_ccc, "Total CCC")
                        fig_zone_ccc = px.bar(
                            camions_zone_ccc,
                            x="Étage - Zone",
                            y="Total CCC",
                            color=color_arg,
                            barmode="group" if color_arg == "Lot" else None,
                        )
                        st.plotly_chart(fig_zone_ccc, key="camions_zone_v1", use_container_width=True)
                    else:
                        st.info("Colonnes camions CCC manquantes")

                # -------- Flux camions CCC --------
                with c2:
                    st.markdown("#### Flux mensuel de camions CCC")

                    if not flux_camions_ccc.empty:

                        flux_camions_plot_v1 = flux_camions_ccc.copy()
                        flux_camions_plot_v1["Nombre de Camions CCC"] = pd.to_numeric(
                            flux_camions_plot_v1["Nombre de Camions CCC"], errors="coerce"
                        ).fillna(0)

                        color_arg = _lot_color(flux_camions_plot_v1, None)
                        fig_flux_ccc = px.area(
                            flux_camions_plot_v1,
                            x="Mois",
                            y="Nombre de Camions CCC",
                            color=color_arg,
                        )

                        # PIC camions V1
                        if mois_pic_camions_v1:
                            fig_flux_ccc.add_vline(
                                x=mois_pic_camions_v1,
                                line_dash="dot",
                                line_color="red",
                            )
                            fig_flux_ccc.add_hline(
                                y=pic_camions_v1,
                                line_dash="dot",
                                line_color="red",
                            )
                            fig_flux_ccc.add_scatter(
                                x=[mois_pic_camions_v1],
                                y=[pic_camions_v1],
                                mode="markers",
                                marker=dict(color="red", size=10),
                                name="Pic",
                            )
                            fig_flux_ccc.add_annotation(
                                x=mois_pic_camions_v1,
                                y=pic_camions_v1,
                                text=f"Pic : {pic_camions_v1:.0f} camions ({mois_pic_camions_v1})",
                                showarrow=True,
                                arrowhead=2,
                                ax=0,
                                ay=-40,
                                font=dict(color="red"),
                            )

                        fig_flux_ccc.update_layout(margin=dict(l=10, r=10, t=30, b=40))

                        st.plotly_chart(
                            fig_flux_ccc,
                            use_container_width=True,
                            key="flux_camions_ccc_v1",
                        )

                    else:
                        st.info("Aucun flux de camions CCC")

                    
                    
                    
                    
                    
                    
                    
                    

                # -------- Remplissage CCC --------
                st.markdown("#### Remplissage par étage (CCC)")
                if not rempl_zone_ccc.empty:
                    color_arg = _lot_color(rempl_zone_ccc, "Remplissage (%)")
                    fig_r_ccc = px.bar(
                        rempl_zone_ccc,
                        x="Étage - Zone",
                        y="Remplissage (%)",
                        color=color_arg,
                        barmode="group" if color_arg == "Lot" else None,
                    )
                    st.plotly_chart(fig_r_ccc, key="remplissage_ccc_v1", use_container_width=True)
                else:
                    st.info("Aucune donnée de remplissage CCC disponible")
                

                # Typologie des camions (V1 ? CCC, agr?gation Python)

                st.markdown("## ?? Typologie des camions (CCC)")

                df_camions_ccc = extras.get("cam_type_ccc", pd.DataFrame())

                if df_camions_ccc.empty:
                    st.info("Aucun camion CCC trouv?.")
                else:
                    has_multi_lot = "Lot" in df_camions_ccc.columns and df_camions_ccc["Lot"].nunique() > 1
                    for lot, df_lot in df_camions_ccc.groupby("Lot") if "Lot" in df_camions_ccc.columns else [("GLOBAL", df_camions_ccc)]:
                        if has_multi_lot:
                            st.markdown(f"#### Lot : {lot}")

                        for _, row in df_lot.iterrows():
                            nom_camion = str(row["Type de Camion"]).strip()
                            quantite = int(row["Nombre de Camions"])

                            # Filtrer uniquement les camions connus
                            if nom_camion not in daba.liste_camions:
                                continue

                            img_path = f"images/image_camions/{nom_camion}.png"
                            c_img, c_nom, c_nb = st.columns([1, 3, 1])

                            with c_img:
                                try:
                                    st.image(img_path, width=70)
                                except:
                                    st.write("??")

                            with c_nom:
                                st.write(f"**{nom_camion}**")

                            with c_nb:
                                st.write(f"**{quantite}**")



        else: 


            # ========= Sous-onglets internes =========
            ong_hyp, ong_pal, ong_cam = st.tabs(
                ["📘 Hypothèses", "📦 Palettes", "🚚 Camions"]
            )

            # ======================================================
            # 📘 ONGLET HYPOTHÈSES (V0)
            # ======================================================
            with ong_hyp:
                # ---------- Document de source ----------
                st.markdown("### 📘 Document de source : DPGF + Indice + PIC")

                # Pré-remplissage DPGF + Indice à partir de dpgf_date et planning_indice
                if dpgf_date and planning_indice:
                    default_dpgf = f"DPGF du {dpgf_date} – Indice {planning_indice}"
                elif dpgf_date:
                    default_dpgf = f"DPGF du {dpgf_date}"
                elif planning_indice:
                    default_dpgf = f"Indice {planning_indice}"
                else:
                    default_dpgf = ""

                dpgf_txt = st.text_area(
                    "DPGF + Indice :",
                    value=default_dpgf,
                    key="dpgf_v0",
                    placeholder="DPGF du … – Indice …"
                )

                st.file_uploader("Veuillez joindre le fichier PIC", key="pic_v0")

                st.markdown("---")

                # ---------- Hypothèse planning ----------
                st.markdown("### 🕒 Hypothèse planning")
                st.markdown(f"- Planning indice : **{planning_indice or '…'}**")

                planning_hyp = st.text_area(
                    "Hypothèses prises pour planning :",
                    key="planning_hyp_v0",
                    placeholder="Décrire ici les hypothèses utilisées pour le planning…"
                )

                if planning_hyp.strip() == "":
                    st.markdown(
                        "<p style='color:red;'>Hypothèses planning non complétées</p>",
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        "<p style='color:green;'>Hypothèses planning complétées</p>",
                        unsafe_allow_html=True,
                    )

                st.markdown("---")

                # ---------- Hypothèses de l’étude ----------
                st.markdown("### 📄 Hypothèses de l’étude")
                st.markdown("- regroupement du matériel en grandes catégories")
                st.markdown(
                    "- conversion des conditionnements en équivalent palette "
                    "(palette européenne 1,2 × 0,8)"
                )
                st.markdown("- 2 phases de travaux par étage : Production et Terminaux")

                st.markdown("---")

                # ---------- Hypothèse de base par famille ----------
                st.markdown("### 🧩 Hypothèse de base déportée par famille")

                # Familles depuis Tableau Source ou Matériel

                if "Nom de l'élément" in src.columns:
                    familles_src = (
                        src["Nom de l'élément"]
                        .dropna()
                        .astype(str)
                        .sort_values()
                        .unique()
                    )
                elif "Nom" in materiel.columns:
                    familles_src = (
                        materiel["Nom"]
                        .dropna()
                        .astype(str)
                        .sort_values()
                        .unique()
                    )
                else:
                    familles_src = []

                if len(familles_src) == 0:
                    st.info("Aucune famille trouvée.")
                else:
                    df_fam = pd.DataFrame({
                        "Famille": familles_src,
                        "Stocké en CCC ?": ["❌"] * len(familles_src),
                        "Quantité": [0] * len(familles_src),  # 👈 logique explicite
                    })

                    st.dataframe(df_fam, use_container_width=True)

            # ======================================================
            # 📦 ONGLET PALETTES (V0)
            # ======================================================
            with ong_pal:
                st.markdown("### 📦 Palettes")

                # Métriques : palettes + surface
                colA, colB = st.columns(2)

                with colA:
                    st.metric(
                        "Palettes équivalentes totales (V0)",
                        f"{total_palettes:,.0f}".replace(",", " "),
                    )

                surface_totale_v0 = total_palettes * 0.96
                with colB:
                    st.metric(
                        "Surface totale occupée (m²)",
                        f"{surface_totale_v0:,.0f}".replace(",", " "),
                    )

                # Deux graphiques côte à côte
                c1, c2 = st.columns(2)

                # --------------------------------------------------
                # Graphique horizontal : répartition des matériaux (V0)
                # --------------------------------------------------
                with c1:
                    st.markdown("#### Répartition des matériaux")

                    df_v0_mat = extras.get("mat_v0", pd.DataFrame())
                    if not df_v0_mat.empty:
                        color_arg = _lot_color(df_v0_mat, None)
                        fig_mat_v1 = px.bar(
                            df_v0_mat,
                            x="Nombre total de matériels",
                            y="Matériel complet",
                            color=color_arg,
                            orientation="h",
                            title="Répartition des matériaux",
                        )
                        st.plotly_chart(fig_mat_v1, key="bar_mat_v0", use_container_width=True)
                    else:
                        st.info("Colonnes absentes")

                # --------------------------------------------------
                # Flux mensuel de palettes + PIC (V0)
                # --------------------------------------------------
                with c2:
                    st.markdown("#### Flux mensuel de palettes")

                    flux_palettes_plot_v0 = flux_palettes.copy()
                    if not flux_palettes_plot_v0.empty:
                        flux_palettes_plot_v0["Mois"] = flux_palettes_plot_v0["Mois"].astype(str)
                    mois_pic_palettes_str_v0 = str(mois_pic_palettes)

                    color_arg = _lot_color(flux_palettes_plot_v0, None)
                    fig_flux_pal_v0 = px.area(
                        flux_palettes_plot_v0,
                        x="Mois",
                        y="Volume (nombre de palettes équivalentes)",
                        color=color_arg,
                    )

                    if pic_palettes > 0:
                        fig_flux_pal_v0.add_vline(
                            x=mois_pic_palettes_str_v0, line_dash="dot", line_color="red"
                        )
                        fig_flux_pal_v0.add_hline(
                            y=pic_palettes, line_dash="dot", line_color="red"
                        )
                        fig_flux_pal_v0.add_scatter(
                            x=[mois_pic_palettes_str_v0],
                            y=[pic_palettes],
                            mode="markers",
                            marker=dict(color="red", size=10),
                            name="Pic de livraison",
                        )
                        fig_flux_pal_v0.add_annotation(
                            x=mois_pic_palettes_str_v0,
                            y=pic_palettes,
                            text=(
                                f"Pic : {pic_palettes:.0f} palettes "
                                f"({mois_pic_palettes_str_v0})"
                            ),
                            showarrow=True,
                            arrowhead=2,
                            ax=0,
                            ay=-40,
                            font=dict(color="red"),
                        )

                    fig_flux_pal_v0.update_layout(
                        margin=dict(l=10, r=10, t=30, b=40)
                    )
                    st.plotly_chart(
                        fig_flux_pal_v0,
                        use_container_width=True,
                        height=320,
                        key="flux_palettes_v0",
                    )

                # --------------------------------------------------
                # Répartition par étage / zone (plein largeur)
                # --------------------------------------------------
                st.markdown("#### Répartition des palettes par étage / zone")
                color_arg = _lot_color(palettes_zone, "Palettes")
                fig_pal_zone_v0 = px.bar(
                    palettes_zone,
                    x="Étage - Zone",
                    y="Palettes",
                    color=color_arg,
                    barmode="group" if color_arg == "Lot" else None,
                )
                fig_pal_zone_v0.update_layout(margin=dict(l=10, r=10, t=30, b=40))
                st.plotly_chart(
                    fig_pal_zone_v0,
                    use_container_width=True,
                    height=320,
                    key="palettes_zone_v0",
                )

            # ======================================================
            # 🚚 ONGLET CAMIONS (V0)
            # ======================================================
            with ong_cam:
                st.markdown("### 🚚 Camions ")

                colA, colB = st.columns(2)
                with colA:
                    st.metric(
                        "Nombre total de camions ",
                        f"{total_camions:,.0f}".replace(",", " "),
                    )
                with colB:
                    st.metric(
                        "Remplissage moyen camions ",
                        f"{rempl_moyen:.1f} %",
                    )

                c2_, c3_ = st.columns(2)

                # -----------------------------------------------------------
                # CAMIONS PAR ÉTAGE / ZONE
                # -----------------------------------------------------------
                with c2_:
                    st.markdown("#### Camions par étage ")
                    color_arg = _lot_color(camions_zone, "Camions totaux")
                    fig_cam_zone = px.bar(
                        camions_zone,
                        x="Étage - Zone",
                        y="Camions totaux",
                        color=color_arg,
                        barmode="group" if color_arg == "Lot" else None,
                    )
                    fig_cam_zone.update_layout(
                        margin=dict(l=10, r=10, t=30, b=40)
                    )
                    st.plotly_chart(
                        fig_cam_zone,
                        use_container_width=True,
                        key="cam_v0_zone",
                    )

                # -----------------------------------------------------------
                # FLUX MENSUEL CAMIONS (CORRIGÉ AVEC LIGNE HORIZONTALE)
                # -----------------------------------------------------------
                with c3_:
                    st.markdown("#### Flux mensuel de camions ")

                    # Sécurisation des données (important)
                    flux_camions_plot = flux_camions.copy()
                    flux_camions_plot["Nombre de Camions"] = pd.to_numeric(
                        flux_camions_plot["Nombre de Camions"], errors="coerce"
                    ).fillna(0)

                    color_arg = _lot_color(flux_camions_plot, None)
                    fig_flux_cam = px.area(
                        flux_camions_plot,
                        x="Mois",
                        y="Nombre de Camions",
                        color=color_arg,
                    )

                    # PIC – Affichage même si pic_camions == 0
                    if not flux_camions_plot.empty and mois_pic_camions:

                        # Ligne verticale
                        fig_flux_cam.add_vline(
                            x=mois_pic_camions,
                            line_dash="dot",
                            line_color="red"
                        )

                        # Ligne horizontale (toujours affichée)
                        fig_flux_cam.add_hline(
                            y=pic_camions,
                            line_dash="dot",
                            line_color="red"
                        )

                        # Point rouge
                        fig_flux_cam.add_scatter(
                            x=[mois_pic_camions],
                            y=[pic_camions],
                            mode="markers",
                            marker=dict(color="red", size=10),
                            name="Pic de livraison",
                        )

                        # Annotation
                        fig_flux_cam.add_annotation(
                            x=mois_pic_camions,
                            y=pic_camions,
                            text=f"Pic : {pic_camions:.0f} camions ({mois_pic_camions})",
                            showarrow=True,
                            arrowhead=2,
                            ax=0,
                            ay=-40,
                            font=dict(color="red"),
                        )

                    fig_flux_cam.update_layout(
                        margin=dict(l=10, r=10, t=30, b=40)
                    )

                    st.plotly_chart(
                        fig_flux_cam,
                        use_container_width=True,
                        key="cam_v0_flux_pic",
                    )

                # -----------------------------------------------------------
                # REMPLISSAGE PAR ÉTAGE / ZONE
                # -----------------------------------------------------------
                st.markdown("#### Remplissage des camions par étage ")
                if not rempl_zone.empty:
                    color_arg = _lot_color(rempl_zone, "Remplissage (%)")
                    fig_rempl = px.bar(
                        rempl_zone,
                        x="Étage - Zone",
                        y="Remplissage (%)",
                        color=color_arg,
                        color_continuous_scale="Purples" if color_arg != "Lot" else None,
                        barmode="group" if color_arg == "Lot" else None,
                    )
                    fig_rempl.update_layout(
                        margin=dict(l=10, r=10, t=30, b=40)
                    )
                    st.plotly_chart(
                        fig_rempl,
                        use_container_width=True,
                        key="cam_v0_rempl",
                    )
                else:
                    st.info("Aucune donnée de remplissage disponible.")

    
                # Typologie des camions (V0 ? agr?gation Python)

                st.markdown("## ?? Typologie des camions (V0)")

                df_camions_v0 = extras.get("cam_type_v0", pd.DataFrame())

                if df_camions_v0.empty:
                    st.info("Aucun camion trouv?.")
                else:
                    has_multi_lot = "Lot" in df_camions_v0.columns and df_camions_v0["Lot"].nunique() > 1
                    for lot, df_lot in df_camions_v0.groupby("Lot") if "Lot" in df_camions_v0.columns else [("GLOBAL", df_camions_v0)]:
                        if has_multi_lot:
                            st.markdown(f"#### Lot : {lot}")

                        for _, row in df_lot.iterrows():
                            nom_camion = str(row["Type de Camion"]).strip()
                            quantite = int(row["Nombre de Camions"])

                            # Filtrer uniquement les camions connus
                            if nom_camion not in daba.liste_camions:
                                continue

                            img_path = f"images/image_camions/{nom_camion}.png"
                            c_img, c_nom, c_nb = st.columns([1, 3, 1])

                            with c_img:
                                try:
                                    st.image(img_path, width=70)
                                except:
                                    st.write("??")

                            with c_nom:
                                st.write(f"**{nom_camion}**")

                            with c_nb:
                                st.write(f"**{quantite}**")



    # ======================= VARIANTES =======================
    with tab_var:
        st.subheader("Variantes personnalisées")

        # Choix type de la nouvelle variante (même logique V0/V1)
        choix_type = st.radio(
            "Type de la nouvelle variante :",
            ["Sans CCC", "Avec CCC"],
            horizontal=True,
            key="type_variante_crea"
        )

        # --- CRÉATION D'UNE NOUVELLE VARIANTE ---
        if st.button("Créer une variante"):
            if not file_bytes:
                st.warning("Création de variantes indisponible sans fichier Excel de base.")
                st.stop()
            vid = f"V{st.session_state['variant_counter']}"

            # Détection du format XLSX / XLSM à partir du fichier d'origine
            import zipfile
            try:
                zip_test = zipfile.ZipFile(io.BytesIO(file_bytes))
                if any("vbaProject.bin" in f.filename for f in zip_test.filelist):
                    ext = ".xlsm"
                else:
                    ext = ".xlsx"
            except:
                st.error("Le fichier Excel d'origine est invalide.")
                st.stop()

            file_name = f"{vid}{ext}"

            # Écriture du fichier sur disque
            try:
                with open(file_name, "wb") as f:
                    f.write(file_bytes)
            except Exception as e:
                st.error(f"Erreur lors de la création de la variante : {e}")
                st.stop()

            # Ajout dans la session
            with open(file_name, "rb") as f:
                st.session_state["variants"][vid] = {
                    "with_ccc": (choix_type == "Avec CCC"),
                    "bytes": f.read(),
                    "ext": ext,
                }

            # Bouton de téléchargement immédiat
            mime = (
                "application/vnd.ms-excel" if ext == ".xlsm"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            with open(file_name, "rb") as f:
                st.download_button(
                    f"Télécharger {file_name}",
                    data=f.read(),
                    file_name=file_name,
                    mime=mime,
                    key=f"download_{vid}_creation",
                )

            st.session_state["variant_counter"] += 1
            st.success(f"Variante {vid} créée.")

        # --- LISTE DES VARIANTES EXISTANTES ---
        variants = st.session_state.get("variants", {})

        if not variants:
            st.info("Aucune variante pour le moment.")
        else:
            for vid, meta in variants.items():

                st.markdown(
                    f"### {vid} – {'avec CCC' if meta['with_ccc'] else 'sans CCC'}"
                )

                # Bouton téléchargement du fichier de la variante
                if meta.get("bytes"):
                    st.download_button(
                        f"Télécharger {vid}{meta.get('ext', '.xlsx')}",
                        data=meta["bytes"],
                        file_name=f"{vid}{meta.get('ext', '.xlsx')}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{vid}_existing",
                    )

                # Upload d'un fichier Excel modifié pour cette variante
                uploaded_var = st.file_uploader(
                    f"Uploader le fichier modifié pour {vid}",
                    type=["xlsx", "xlsm"],
                    key=f"upload_{vid}",
                )

                if uploaded_var is not None:
                    meta["bytes"] = uploaded_var.read()
                    st.success(f"Fichier de {vid} mis à jour.")

                if meta.get("bytes") is None:
                    st.info("Aucun fichier chargé pour cette variante.")
                    continue

                # ---------- Lecture des feuilles du fichier VARIANTE ----------
                try:
                    excel_io_var = io.BytesIO(meta["bytes"])
                    xls_var = pd.ExcelFile(excel_io_var)
                    bg_var = xls_var.parse("Bilan Graphique")
                    param_var = xls_var.parse("Paramétrage")
                    materiel_var = xls_var.parse("Matériel")
                    src_var = xls_var.parse("Tableau Source")
                except Exception as e:
                    st.error(f"Erreur lecture Excel pour {vid}: {e}")
                    continue

                # ---------- Application du pipeline adapté ----------
                if meta["with_ccc"]:
                    metrics_var = pipeline_avec_ccc(bg_var)
                else:
                    metrics_var = pipeline_sans_ccc(bg_var)

                palettes_zone_var = metrics_var["palettes_zone"]
                total_palettes_var = metrics_var["total_palettes"]
                flux_palettes_var = metrics_var["flux_palettes"]
                mois_pic_palettes_var = metrics_var["mois_pic_palettes"]
                pic_palettes_var = metrics_var["pic_palettes"]
                camions_zone_var = metrics_var["camions_zone"]
                total_camions_var = metrics_var["total_camions"]
                flux_camions_var = metrics_var["flux_camions"]
                mois_pic_camions_var = metrics_var["mois_pic_camions"]
                pic_camions_var = metrics_var["pic_camions"]
                rempl_zone_var = metrics_var["rempl_zone"]
                rempl_moyen_var = metrics_var["rempl_moyen"]
                camions_type_var = metrics_var["camions_type"]

                # ======================= DASHBOARD VARIANTE =======================
                with st.expander(f"Afficher le dashboard de {vid}", expanded=False):

                    # EXACTEMENT la même structure : Hypothèses / Palettes / Camions
                    ong_hyp_v, ong_pal_v, ong_cam_v = st.tabs(
                        ["📘 Hypothèses", "📦 Palettes", "🚚 Camions"]
                    )

                    # ============================================================
                    #  VARIANTE SANS CCC  → miroir du DASHBOARD V0
                    # ============================================================
                    if not meta["with_ccc"]:

                        # ---------------- HYPO V0----------------
                        with ong_hyp_v:
                            st.markdown("### 📘 Hypothèses – Variante Sans CCC (V0)")
                            st.markdown(
                                "- Même logique que V0, mais toutes les valeurs proviennent "
                                "du fichier Excel de la variante."
                            )

                            # Hypothèses de l’étude (idem V0)
                            st.markdown("### 📄 Hypothèses de l’étude")
                            st.markdown("- regroupement du matériel en grandes catégories")
                            st.markdown(
                                "- conversion des conditionnements en équivalent palette "
                                "(palette européenne 1,2 × 0,8)"
                            )
                            st.markdown(
                                "- 2 phases de travaux par étage : Production et Terminaux"
                            )

                            st.markdown("---")

                            # Hypothèse de base déportée par famille (variante)
                            st.markdown("### 🧩 Hypothèse de base déportée par famille")

                            if "Nom de l'élément" in src_var.columns:
                                familles_src_v = (
                                    src_var["Nom de l'élément"]
                                    .dropna()
                                    .astype(str)
                                    .sort_values()
                                    .unique()
                                )
                            elif "Nom" in materiel_var.columns:
                                familles_src_v = (
                                    materiel_var["Nom"]
                                    .dropna()
                                    .astype(str)
                                    .sort_values()
                                    .unique()
                                )
                            else:
                                familles_src_v = []

                            if len(familles_src_v) == 0:
                                st.info(
                                    "Aucune famille trouvée dans Tableau Source / Matériel de la variante."
                                )
                            else:
                                df_fam_v = pd.DataFrame(
                                    {
                                        "Famille": familles_src_v,
                                        "Stocké en CCC ?": ["❌" for _ in familles_src_v],
                                    }
                                )
                                st.dataframe(df_fam_v, use_container_width=True)

                        # ---------------- PALETTES V0 (variante) ----------------
                        with ong_pal_v:
                            st.markdown("### 📦 Palettes – Variante Sans CCC (V0)")

                            cA, cB = st.columns(2)
                            with cA:
                                st.metric(
                                    "Palettes équivalentes totales (Variante)",
                                    f"{total_palettes_var:,.0f}".replace(",", " "),
                                )
                            with cB:
                                st.metric(
                                    "Surface totale occupée (m²)",
                                    f"{(total_palettes_var * 0.96):,.0f}".replace(",", " "),
                                )

                            c1, c2 = st.columns(2)

                            # Répartition matériaux (on refait comme en V0 mais sur bg_var)
                            with c1:
                                st.markdown("#### Répartition des matériaux ")

                                design_col_v = None
                                qty_col_v = None
                                for c in bg_var.columns:
                                    cname = str(c).strip().lower()
                                    if cname.startswith("désignation"):
                                        design_col_v = c
                                    if cname.startswith("quantité"):
                                        qty_col_v = c

                                if design_col_v is not None and qty_col_v is not None:
                                    df_mat_v0_var = (
                                        bg_var[[design_col_v, qty_col_v]]
                                        .dropna(subset=[design_col_v])
                                        .groupby(design_col_v, as_index=False)[qty_col_v]
                                        .sum()
                                    )

                                    N = 15
                                    df_mat_v0_var = df_mat_v0_var.sort_values(
                                        qty_col_v, ascending=True
                                    ).tail(N)

                                    fig_bar_v0_var = px.bar(
                                        df_mat_v0_var,
                                        x=qty_col_v,
                                        y=design_col_v,
                                        orientation="h",
                                        text=qty_col_v,
                                    )
                                    fig_bar_v0_var.update_traces(textposition="outside")
                                    fig_bar_v0_var.update_layout(
                                        title="Répartition des matériaux (variante V0)",
                                        margin=dict(l=10, r=10, t=40, b=10),
                                        height=380,
                                        showlegend=False,
                                    )
                                    st.plotly_chart(
                                        fig_bar_v0_var,
                                        use_container_width=True,
                                    )
                                else:
                                    st.info(
                                        "Impossible de trouver les colonnes 'Désignation' et 'Quantité' dans le BG de la variante."
                                    )

                            # Flux mensuel palettes (variante)
                            with c2:
                                st.markdown("#### Flux mensuel de palettes ")

                                flux_palettes_plot_v = flux_palettes_var.copy()
                                if not flux_palettes_plot_v.empty:
                                    flux_palettes_plot_v["Mois"] = flux_palettes_plot_v["Mois"].astype(str)

                                # 2e colonne = valeur
                                if not flux_palettes_plot_v.empty:
                                    y_col_pal = [
                                        c for c in flux_palettes_plot_v.columns if c != "Mois"
                                    ][0]
                                else:
                                    y_col_pal = "Volume (nombre de palettes équivalentes)"

                                fig_flux_pal_v = px.area(
                                    flux_palettes_plot_v,
                                    x="Mois",
                                    y=y_col_pal,
                                )

                                if pic_palettes_var > 0 and mois_pic_palettes_var:
                                    mois_pic_palettes_str_v = str(mois_pic_palettes_var)
                                    fig_flux_pal_v.add_vline(
                                        x=mois_pic_palettes_str_v,
                                        line_dash="dot",
                                        line_color="red",
                                    )
                                    fig_flux_pal_v.add_hline(
                                        y=pic_palettes_var,
                                        line_dash="dot",
                                        line_color="red",
                                    )
                                    fig_flux_pal_v.add_scatter(
                                        x=[mois_pic_palettes_str_v],
                                        y=[pic_palettes_var],
                                        mode="markers",
                                        marker=dict(color="red", size=10),
                                        name="Pic de livraison",
                                    )

                                fig_flux_pal_v.update_layout(
                                    margin=dict(l=10, r=10, t=30, b=40)
                                )
                                st.plotly_chart(
                                    fig_flux_pal_v,
                                    use_container_width=True,
                                )

                            # Palettes par étage / zone
                            st.markdown("#### Répartition des palettes par étage / zone ")
                            fig_pal_zone_v = px.bar(
                                palettes_zone_var,
                                x="Étage - Zone",
                                y="Palettes",
                                color="Palettes",
                            )
                            fig_pal_zone_v.update_layout(
                                margin=dict(l=10, r=10, t=30, b=40)
                            )
                            st.plotly_chart(
                                fig_pal_zone_v,
                                use_container_width=True,
                            )

                        # ---------------- CAMIONS V0 (variante) ----------------
                        with ong_cam_v:
                            st.markdown("### 🚚 Camions – Variante Sans CCC (V0)")

                            colA, colB = st.columns(2)
                            with colA:
                                st.metric(
                                    "Nombre total de camions",
                                    f"{total_camions_var:,.0f}".replace(",", " "),
                                )
                            with colB:
                                st.metric(
                                    "Remplissage moyen camions ",
                                    f"{rempl_moyen_var:.1f} %",
                                )

                            c2_, c3_ = st.columns(2)

                            # Camions par étage / zone
                            with c2_:
                                st.markdown("#### Camions par étage ")
                                fig_cam_zone_v = px.bar(
                                    camions_zone_var,
                                    x="Étage - Zone",
                                    y=camions_zone_var.columns[-1],
                                    color=camions_zone_var.columns[-1],
                                )
                                fig_cam_zone_v.update_layout(
                                    margin=dict(l=10, r=10, t=30, b=40)
                                )
                                st.plotly_chart(
                                    fig_cam_zone_v,
                                    use_container_width=True,
                                )

                            # Flux mensuel camions (variante)
                            with c3_:
                                st.markdown("#### Flux mensuel de camions ")

                                fc_v = flux_camions_var.copy()
                                if not fc_v.empty:
                                    fc_v["Mois"] = fc_v["Mois"].astype(str)
                                    # Nom de la colonne Y (peut être 'Nombre de Camions' ou 'Camions')
                                    y_col_cam = [c for c in fc_v.columns if c != "Mois"][0]
                                    fc_v[y_col_cam] = pd.to_numeric(
                                        fc_v[y_col_cam],
                                        errors="coerce",
                                    ).fillna(0)
                                else:
                                    y_col_cam = "Camions"

                                fig_flux_cam_v = px.area(
                                    fc_v,
                                    x="Mois",
                                    y=y_col_cam,
                                )

                                if mois_pic_camions_var:
                                    fig_flux_cam_v.add_vline(
                                        x=mois_pic_camions_var,
                                        line_dash="dot",
                                        line_color="red",
                                    )
                                    fig_flux_cam_v.add_hline(
                                        y=pic_camions_var,
                                        line_dash="dot",
                                        line_color="red",
                                    )

                                fig_flux_cam_v.update_layout(
                                    margin=dict(l=10, r=10, t=30, b=40)
                                )
                                st.plotly_chart(
                                    fig_flux_cam_v,
                                    use_container_width=True,
                                )

                            # Remplissage par étage / zone
                            st.markdown("#### Remplissage des camions par étage")
                            if not rempl_zone_var.empty:
                                fig_rempl_v = px.bar(
                                    rempl_zone_var,
                                    x="Étage - Zone",
                                    y="Remplissage (%)",
                                    color="Remplissage (%)",
                                )
                                fig_rempl_v.update_layout(
                                    margin=dict(l=10, r=10, t=30, b=40)
                                )
                                st.plotly_chart(
                                    fig_rempl_v,
                                    use_container_width=True,
                                )
                            else:
                                st.info("Aucune donnée de remplissage disponible pour cette variante.")

                            # Typologie camions (variante) – même info que V0
                            # -------------------------------------------------------
                            # 🚚 Typologie des camions – Variante Sans CCC (structure V0)
                            # -------------------------------------------------------

                            st.markdown("## 🚚 Typologie des camions ")

                            try:
                                # Identification stricte des colonnes T et U (comme V0)
                                col_type = bg_var.columns[19]   # colonne T
                                col_nb   = bg_var.columns[20]   # colonne U

                                df_camions_var = (
                                    bg_var[[col_type, col_nb]]
                                    .dropna(subset=[col_type])
                                    .groupby(col_type, as_index=False)[col_nb]
                                    .sum()
                                    .rename(columns={
                                        col_type: "Type de Camion",
                                        col_nb:   "Nombre de Camions"
                                    })
                                )

                                if df_camions_var.empty:
                                    st.info("Aucun camion trouvé dans les colonnes T et U pour cette variante.")
                                else:
                                    for _, row in df_camions_var.iterrows():
                                        nom_camion = str(row["Type de Camion"]).strip()
                                        quantite = int(row["Nombre de Camions"])

                                        # Filtrer uniquement les camions connus
                                        if nom_camion not in daba.liste_camions:
                                            continue

                                        img_path = f"images/image_camions/{nom_camion}.png"
                                        c_img, c_nom, c_nb = st.columns([1, 3, 1])

                                        with c_img:
                                            try:
                                                st.image(img_path, width=70)
                                            except:
                                                st.write("🚚")

                                        with c_nom:
                                            st.write(f"**{nom_camion}**")

                                        with c_nb:
                                            st.write(f"**{quantite}**")

                            except Exception as e:
                                st.error(f"Erreur lecture typologie variante Sans CCC : {e}")

                    # ============================================================
                    #  VARIANTE AVEC CCC  → miroir du DASHBOARD V1
                    # ============================================================
                    else:
                        # ---------------- HYPO V1 (variante) ----------------
                        with ong_hyp_v:
                            st.markdown("### 📘 Hypothèses ")


                            # Paramètres CCC depuis la feuille Paramétrage de la variante
                            st.markdown("#### Paramètres CCC ")
                            def _val_param_ccc_var(libel):
                                try:
                                    lot_col_var = param_var.columns[1]
                                    return param_var.loc[
                                        param_var["Lot"] == libel,
                                        lot_col_var,
                                    ].iloc[0]
                                except Exception:
                                    return ""

                            duree_ccc_v = _val_param_ccc_var("Durée de stockage CCC (en mois)")
                            tarif_mois_v = _val_param_ccc_var("Tarif mois de stockage (en €)")
                            frais_sup_v = _val_param_ccc_var("Frais supplémentaires/palette (en €)")
                            frais_liv_v = _val_param_ccc_var("Frais de livraison par camion")

                            st.markdown(f"- Durée stockage : **{duree_ccc_v} mois**")
                            st.markdown(f"- Tarif de stockage : **{tarif_mois_v} €/mois**")
                            st.markdown(f"- Frais supplémentaires/palette : **{frais_sup_v} €**")
                            st.markdown(f"- Frais de livraison : **{frais_liv_v} €**")

                            st.markdown("---")
                            st.markdown("#### Hypothèse de base déportée par famille ")

                            if (
                                {"Nom de l'élément", "Utilisation d'une CCC"}.issubset(src_var.columns)
                                and {"Matériel CCC", "Nombre de matériels CCC"}.issubset(bg_var.columns)
                            ):
                                df_src_loc = src_var[
                                    ["Nom de l'élément", "Utilisation d'une CCC"]
                                ].dropna()
                                df_src_loc["use_ccc"] = (
                                    df_src_loc["Utilisation d'une CCC"]
                                    .astype(str)
                                    .str.lower()
                                    .isin(["oui", "yes", "y", "1"])
                                )

                                df_yes_v = (
                                    df_src_loc.groupby("Nom de l'élément", as_index=False)["use_ccc"]
                                    .any()
                                    .rename(columns={"Nom de l'élément": "Famille"})
                                )

                                df_qty_v = (
                                    bg_var[["Matériel CCC", "Nombre de matériels CCC"]]
                                    .dropna(subset=["Matériel CCC"])
                                    .groupby("Matériel CCC", as_index=False)["Nombre de matériels CCC"]
                                    .sum()
                                    .rename(
                                        columns={
                                            "Matériel CCC": "Famille",
                                            "Nombre de matériels CCC": "Quantité",
                                        }
                                    )
                                )

                                df_merge_v = pd.merge(df_yes_v, df_qty_v, on="Famille", how="left")
                                df_merge_v["Quantité"] = df_merge_v["Quantité"].fillna(0)
                                df_merge_v["Stocké en CCC ?"] = df_merge_v["use_ccc"].apply(
                                    lambda x: "✔️" if x else "❌"
                                )

                                st.dataframe(
                                    df_merge_v[["Famille", "Stocké en CCC ?", "Quantité"]],
                                    use_container_width=True,
                                )
                            else:
                                st.info(
                                    "Colonnes nécessaires introuvables dans Tableau Source / Bilan Graphique de la variante."
                                )

                        # ---------------- PALETTES V1 (variante) ----------------
                        with ong_pal_v:
                            st.markdown("### 📦 Palettes ")

                            colA, colB = st.columns(2)
                            with colA:
                                st.metric(
                                    "Palettes équivalentes totales (Variante)",
                                    f"{total_palettes_var:,.0f}".replace(",", " "),
                                )
                            with colB:
                                st.metric(
                                    "Surface totale (m²)",
                                    f"{(total_palettes_var * 0.96):,.0f}".replace(",", " "),
                                )

                            c1, c2 = st.columns(2)

                            # Matériaux stockés en CCC (variante)
                            with c1:
                                st.markdown("#### Matériaux stockés en CCC ")

                                mat_col_v = None
                                qty_col_v = None
                                for c in bg_var.columns:
                                    cname = str(c).lower()
                                    if cname.startswith("matériel ccc"):
                                        mat_col_v = c
                                    if cname.startswith("nombre de matériels ccc"):
                                        qty_col_v = c

                                if mat_col_v and qty_col_v:
                                    df_v1_mat_var = (
                                        bg_var[[mat_col_v, qty_col_v]]
                                        .dropna(subset=[mat_col_v])
                                        .groupby(mat_col_v, as_index=False)[qty_col_v]
                                        .sum()
                                    )

                                    fig_mat_v1_var = px.bar(
                                        df_v1_mat_var,
                                        x=qty_col_v,
                                        y=mat_col_v,
                                        orientation="h",
                                        title="Répartition des matériaux stockés en CCC ",
                                    )
                                    st.plotly_chart(
                                        fig_mat_v1_var,
                                        use_container_width=True,
                                    )
                                else:
                                    st.info("Colonnes Matériel CCC absentes dans le BG de la variante.")

                            # Flux palettes (identique logique V1)
                            with c2:
                                st.markdown("#### Flux mensuel de palettes")

                                flux_palettes_plot_v1 = flux_palettes_var.copy()
                                if not flux_palettes_plot_v1.empty:
                                    flux_palettes_plot_v1["Mois"] = flux_palettes_plot_v1[
                                        "Mois"
                                    ].astype(str)
                                    y_col_pal_v1 = [
                                        c for c in flux_palettes_plot_v1.columns if c != "Mois"
                                    ][0]
                                else:
                                    y_col_pal_v1 = "Volume (nombre de palettes équivalentes)"

                                fig_flux_pal_v1_var = px.area(
                                    flux_palettes_plot_v1,
                                    x="Mois",
                                    y=y_col_pal_v1,
                                )

                                if pic_palettes_var > 0 and mois_pic_palettes_var:
                                    mois_pic_palettes_str_v1 = str(mois_pic_palettes_var)
                                    fig_flux_pal_v1_var.add_vline(
                                        x=mois_pic_palettes_str_v1,
                                        line_dash="dot",
                                        line_color="red",
                                    )
                                    fig_flux_pal_v1_var.add_hline(
                                        y=pic_palettes_var,
                                        line_dash="dot",
                                        line_color="red",
                                    )

                                fig_flux_pal_v1_var.update_layout(
                                    margin=dict(l=10, r=10, t=30, b=40)
                                )
                                st.plotly_chart(
                                    fig_flux_pal_v1_var,
                                    use_container_width=True,
                                )

                            # Répartition palettes par étage / zone
                            st.markdown("#### Répartition des palettes par étage / zone ")
                            fig_pal_v1_var = px.bar(
                                palettes_zone_var,
                                x="Étage - Zone",
                                y="Palettes",
                                color="Palettes",
                            )
                            st.plotly_chart(
                                fig_pal_v1_var,
                                use_container_width=True,
                            )

                        # ---------------- CAMIONS V1 (variante) ----------------
                        with ong_cam_v:
                            st.markdown("### 🚚 Camions – Variante Avec CCC (V1)")

                            colA, colB = st.columns(2)
                            with colA:
                                st.metric(
                                    "Nombre total de camions (CCC – Variante)",
                                    f"{total_camions_var:,.0f}".replace(",", " "),
                                )
                            with colB:
                                st.metric(
                                    "Remplissage moyen (CCC – Variante)",
                                    f"{rempl_moyen_var:.1f} %",
                                )

                            c1, c2 = st.columns(2)

                            # Camions par étage / zone CCC
                            with c1:
                                st.markdown("#### Camions par étage (CCC – Variante)")
                                if not camions_zone_var.empty:
                                    # colonne Y = 'Camions' (pipeline_avec_ccc)
                                    y_col_zone = [
                                        c for c in camions_zone_var.columns if c != "Étage - Zone"
                                    ][0]
                                    fig_zone_ccc_var = px.bar(
                                        camions_zone_var,
                                        x="Étage - Zone",
                                        y=y_col_zone,
                                        color=y_col_zone,
                                    )
                                    st.plotly_chart(
                                        fig_zone_ccc_var,
                                        use_container_width=True,
                                    )
                                else:
                                    st.info("Aucune donnée de camions CCC par zone pour cette variante.")

                            # Flux camions CCC
                            with c2:
                                st.markdown("#### Flux mensuel de camions CCC ")

                                if not flux_camions_var.empty:
                                    fc_v1 = flux_camions_var.copy()
                                    fc_v1["Mois"] = fc_v1["Mois"].astype(str)
                                    y_col_cam_v1 = [
                                        c for c in fc_v1.columns if c != "Mois"
                                    ][0]
                                    fc_v1[y_col_cam_v1] = pd.to_numeric(
                                        fc_v1[y_col_cam_v1],
                                        errors="coerce",
                                    ).fillna(0)

                                    fig_flux_ccc_v = px.area(
                                        fc_v1,
                                        x="Mois",
                                        y=y_col_cam_v1,
                                    )

                                    if mois_pic_camions_var:
                                        fig_flux_ccc_v.add_vline(
                                            x=mois_pic_camions_var,
                                            line_dash="dot",
                                            line_color="red",
                                        )
                                        fig_flux_ccc_v.add_hline(
                                            y=pic_camions_var,
                                            line_dash="dot",
                                            line_color="red",
                                        )

                                    fig_flux_ccc_v.update_layout(
                                        margin=dict(l=10, r=10, t=30, b=40)
                                    )
                                    st.plotly_chart(
                                        fig_flux_ccc_v,
                                        use_container_width=True,
                                    )
                                else:
                                    st.info("Aucun flux de camions CCC pour cette variante.")

                            # Remplissage CCC
                            st.markdown("#### Remplissage par étage (CCC – Variante)")
                            if not rempl_zone_var.empty:
                                fig_r_ccc_v = px.bar(
                                    rempl_zone_var,
                                    x="Étage - Zone",
                                    y="Remplissage (%)",
                                    color="Remplissage (%)",
                                )
                                st.plotly_chart(
                                    fig_r_ccc_v,
                                    use_container_width=True,
                                )
                            else:
                                st.info("Aucune donnée de remplissage CCC pour cette variante.")

                            # Typologie camions CCC (variante)
                            # -------------------------------------------------------
                            # 🚚 Typologie des camions – Variante Avec CCC (structure V1)
                            # -------------------------------------------------------

                            st.markdown("## 🚚 Typologie des camions – Variante Avec CCC")

                            try:
                                # Colonnes X et Y = index 23 et 24 (comme V1)
                                col_type = bg_var.columns[23]   # colonne X
                                col_nb   = bg_var.columns[24]   # colonne Y

                                df_camions_ccc_var = (
                                    bg_var[[col_type, col_nb]]
                                    .dropna(subset=[col_type])
                                    .groupby(col_type, as_index=False)[col_nb]
                                    .sum()
                                    .rename(columns={
                                        col_type: "Type de Camion",
                                        col_nb:   "Nombre de Camions"
                                    })
                                )

                                if df_camions_ccc_var.empty:
                                    st.info("Aucun camion (CCC) trouvé dans les colonnes X et Y pour cette variante.")
                                else:
                                    for _, row in df_camions_ccc_var.iterrows():
                                        nom_camion = str(row["Type de Camion"]).strip()
                                        quantite = int(row["Nombre de Camions"])

                                        # Filtrer uniquement les camions connus
                                        if nom_camion not in daba.liste_camions:
                                            continue

                                        img_path = f"images/image_camions/{nom_camion}.png"
                                        c_img, c_nom, c_nb = st.columns([1, 3, 1])

                                        with c_img:
                                            try:
                                                st.image(img_path, width=70)
                                            except:
                                                st.write("🚚")

                                        with c_nom:
                                            st.write(f"**{nom_camion}**")

                                        with c_nb:
                                            st.write(f"**{quantite}**")

                            except Exception as e:
                                st.error(f"Erreur lecture typologie variante Avec CCC : {e}")





        # ======================= COMPARATIF — ONGLET COMPLET =======================
        with tab_comp:
            st.subheader("Comparatif multi-versions")

            # --------- 1) Construire la liste de toutes les versions ---------
            all_versions = {
                "V0": {"with_ccc": False, "source": "base"},
                "V1": {"with_ccc": True,  "source": "base"},
            }

            for vid, meta in st.session_state.get("variants", {}).items():
                all_versions[vid] = {
                    "with_ccc": meta["with_ccc"],
                    "source": "variant",
                    "bytes": meta["bytes"],
                }

            version_names = list(all_versions.keys())

            if len(version_names) < 1:
                st.info("Aucune version disponible.")
                st.stop()

            # --------- 2) Choix des versions à comparer (multi-sélection) ---------
            selected_versions = st.multiselect(
                "Choisir les versions à comparer",
                version_names,
                default=[v for v in ["V0", "V1"] if v in version_names],
                key="comp_versions_multiselect",
            )

            if len(selected_versions) == 0:
                st.info("Sélectionner au moins une version.")
                st.stop()



            def compute_metrics_for_version(vname, info):
                col_etage_zone = None
                # Charger les bonnes feuilles
                if info["source"] == "base":
                    bg_loc = bg
                    param_loc = param
                    src_loc = src
                else:
                    try:
                        excel_io_loc = io.BytesIO(info["bytes"])
                        xls_loc = pd.ExcelFile(excel_io_loc)
                        bg_loc = xls_loc.parse("Bilan Graphique")
                        param_loc = xls_loc.parse("Paramétrage")
                        src_loc = xls_loc.parse("Tableau Source")
                    except Exception:
                        return {"ok": False}

                col_etage_zone = _find_col(bg_loc.columns, "Étage - Zone")

                # Palettes
                try:
                    if not col_etage_zone:
                        raise KeyError("Étage - Zone")
                    pal_cols = [col_etage_zone, "Production", "Terminaux"]
                    if "Lot" in bg_loc.columns:
                        pal_cols = ["Lot"] + pal_cols
                    pz = (
                        bg_loc[pal_cols]
                        .dropna(subset=[col_etage_zone])
                        .copy()
                    )
                    pz = pz.rename(columns={col_etage_zone: "Étage - Zone"})
                    pz["Production"] = pz["Production"].fillna(0)
                    pz["Terminaux"] = pz["Terminaux"].fillna(0)
                    pz["Palettes"] = pz["Production"] + pz["Terminaux"]
                    total_pal = float(pz["Palettes"].sum())
                except Exception:
                    pz = pd.DataFrame()
                    total_pal = 0.0

                # Flux palettes
                if {"Mois", "Volume (nombre de palettes équivalentes)"}.issubset(bg_loc.columns):
                    flux_pal = (
                        bg_loc[[c for c in ["Lot", "Mois", "Volume (nombre de palettes équivalentes)"] if c in bg_loc.columns]]
                        .dropna(subset=["Mois"])
                        .groupby(["Lot", "Mois"] if "Lot" in bg_loc.columns else ["Mois"], as_index=False)
                        .sum()
                    )
                else:
                    flux_pal = pd.DataFrame(columns=["Mois", "Volume (nombre de palettes équivalentes)"])

                # Camions
                if info["with_ccc"]:
                    col_cam_total = "Nombre de Camions CCC"
                    col_cam_prod = "Camions Production avec CCC"
                    col_cam_term = "Camions Terminaux avec CCC"
                    col_rempl = "Remplissage camions avec CCC"
                else:
                    col_cam_total = "Nombre de Camions"
                    col_cam_prod = "Camions Production sans CCC"
                    col_cam_term = "Camions Terminaux sans CCC"
                    col_rempl = "Remplissage camions sans CCC"

                # Total camions
                if col_cam_total in bg_loc.columns:
                    total_cam = float(bg_loc[col_cam_total].fillna(0).sum())
                    flux_cam = (
                        bg_loc[[c for c in ["Lot", "Mois", col_cam_total] if c in bg_loc.columns]]
                        .dropna(subset=["Mois"])
                        .groupby(["Lot", "Mois"] if "Lot" in bg_loc.columns else ["Mois"], as_index=False)
                        .sum()
                        .rename(columns={col_cam_total: "Camions"})
                    )
                else:
                    total_cam = 0.0
                    flux_cam = pd.DataFrame(columns=["Mois", "Camions"])

                # Camions zone
                if col_etage_zone and {col_cam_prod, col_cam_term}.issubset(bg_loc.columns):
                    cz = (
                        bg_loc[[c for c in ["Lot", col_etage_zone, col_cam_prod, col_cam_term] if c in bg_loc.columns]]
                        .dropna(subset=[col_etage_zone])
                        .copy()
                    )
                    cz = cz.rename(columns={col_etage_zone: "Étage - Zone"})
                    cz["Camions"] = cz[col_cam_prod].fillna(0) + cz[col_cam_term].fillna(0)
                else:
                    cz = pd.DataFrame(columns=["Étage - Zone", "Camions"])

                # Remplissage
                if col_etage_zone and col_rempl in bg_loc.columns:
                    rz = (
                        bg_loc[[c for c in ["Lot", col_etage_zone, col_rempl] if c in bg_loc.columns]]
                        .dropna(subset=[col_etage_zone])
                        .copy()
                    )
                    rz = rz.rename(columns={col_etage_zone: "Étage - Zone"})
                    rz["Remplissage (%)"] = rz[col_rempl].fillna(0)
                    rbrut = bg_loc[col_rempl].dropna()
                    rmoy = float(rbrut.mean()) if not rbrut.empty else 0.0
                else:
                    rz = pd.DataFrame(columns=["Étage - Zone", "Remplissage (%)"])
                    rmoy = 0.0

                # Coût total logistique
                if "Coût total logistique" in bg_loc.columns:
                    cout_total = float(bg_loc["Coût total logistique"].fillna(0).sum())
                else:
                    cout_total = None

                return {
                    "ok": True,
                    "with_ccc": info["with_ccc"],
                    "bg": bg_loc,
                    "param": param_loc,
                    "src": src_loc,
                    "palettes_zone": pz,
                    "flux_palettes": flux_pal,
                    "total_palettes": total_pal,
                    "camions_zone": cz,
                    "total_camions": total_cam,
                    "flux_camions": flux_cam,
                    "rempl_zone": rz,
                    "rempl_moyen": rmoy,
                    "cout_total": cout_total,
                }

            # --------- 4) Calcul des données pour toutes les versions sélectionnées ---------

            data_versions = {}
            for v in selected_versions:
                data_versions[v] = compute_metrics_for_version(v, all_versions[v])

            if any(not dv["ok"] for dv in data_versions.values()):
                st.error("Impossible de lire les données pour au moins une version.")
                st.stop()

            st.markdown("### 🧩 Performance & Coûts CCC")  # titre plus petit


            def read_ccc_from_bytes(excel_bytes: bytes) -> dict:
                wb = load_workbook(filename=io.BytesIO(excel_bytes), data_only=True, keep_vba=True)
                ws = wb["Bilan Graphique"]

                labels = [ws.cell(row=1, column=c).value for c in range(31, 37)]
                values = [ws.cell(row=2, column=c).value for c in range(31, 37)]

                out = {}
                for k, v in zip(labels, values):
                    if k is None:
                        continue
                    out[str(k).strip()] = v
                return out


            def fmt_percent(x):
                try:
                    return f"{x * 100:.0f} %"
                except Exception:
                    return "—"


            def fmt_euro(x):
                try:
                    return f"{int(round(x)):,}".replace(",", " ") + " €"
                except Exception:
                    return "—"



            for v in selected_versions:
                info = all_versions[v]

                if not info["with_ccc"]:
                    continue

                excel_bytes_v = file_bytes if info["source"] == "base" else info["bytes"]
                if not excel_bytes_v:
                    st.info("Mesures CCC indisponibles sans fichier Excel source.")
                    continue
                ccc = read_ccc_from_bytes(excel_bytes_v)

                st.markdown(f"#### {v}")

                col1, col2, col3 = st.columns(3)

                with col1:
                    with st.container(border=True):
                        st.markdown("**Stock**")
                        st.markdown(f"<h3>{fmt_percent(ccc.get('% Stock CCC'))}</h3>", unsafe_allow_html=True)
                        st.markdown("<span style='color:gray'>Réduction camions</span>", unsafe_allow_html=True)
                        st.markdown(f"<h3>{fmt_percent(ccc.get('% réduction Camions'))}</h3>", unsafe_allow_html=True)

                with col2:
                    with st.container(border=True):
                        st.markdown("**Flux**")
                        st.markdown("<span style='color:gray'>Remplissage moyen</span>", unsafe_allow_html=True)
                        st.markdown(
                            f"<h3>{fmt_percent(ccc.get('% remplissage moyen des camions'))}</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown("<span style='color:gray'>Stockage</span>", unsafe_allow_html=True)
                        st.markdown(
                            f"<h3>{fmt_euro(ccc.get('Coût CCC stockage'))}</h3>",
                            unsafe_allow_html=True
                        )

                with col3:
                    with st.container(border=True):
                        st.markdown("**Coûts**")
                        st.markdown("<span style='color:gray'>Livraison</span>", unsafe_allow_html=True)
                        st.markdown(
                            f"<h3>{fmt_euro(ccc.get('Coût CCC livraison'))}</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown("<span style='color:gray'>Total</span>", unsafe_allow_html=True)
                        st.markdown(
                            f"<h2>{fmt_euro(ccc.get('Coût CCC Total'))}</h2>",
                            unsafe_allow_html=True
                        )

                st.markdown("---")

            # ---- ensuite viennent les onglets ----
            ong_hyp_comp, ong_pal_comp, ong_cam_comp = st.tabs(["📘 Hypothèses", "📦 Palettes", "🚚 Camions"])


            # ======================= ONGLET HYPOTHÈSES =======================
            with ong_hyp_comp:
                st.markdown("### 📘 Hypothèses de l’étude")
                st.markdown("- regroupement du matériel en grandes catégories")
                st.markdown("- conversion en équivalent palette (1,2 × 0,8 m)")
                st.markdown("- 2 phases par étage : Production & Terminaux")

                # ---------- Paramètres CCC par version ----------
                st.markdown("### ⚙️ Paramètres CCC par version")
                for v, dv in data_versions.items():
                    if not dv["with_ccc"]:
                        continue

                    st.markdown(f"#### {v} – Avec CCC")

                    def _get_param_ccc(dv_local, libel):
                        try:
                            col = dv_local["param"].columns[1]
                            return dv_local["param"].loc[dv_local["param"]["Lot"] == libel, col].iloc[0]
                        except Exception:
                            return ""

                    st.markdown(
                        f"- Durée de stockage CCC : **{_get_param_ccc(dv, 'Durée de stockage CCC (en mois)')} mois**"
                    )
                    st.markdown(
                        f"- Tarif mois de stockage : **{_get_param_ccc(dv, 'Tarif mois de stockage (en €)')} €**"
                    )
                    st.markdown(
                        f"- Frais supplémentaires/palette : **{_get_param_ccc(dv, 'Frais supplémentaires/palette (en €)')} €**"
                    )
                    st.markdown(
                        f"- Frais de livraison par camion : **{_get_param_ccc(dv, 'Frais de livraison par camion')} €**"
                    )

                st.markdown("---")


                st.markdown("### 📦 Hypothèse de base déportée par famille ")

                # df_final existe déjà avec la colonne "Famille"
                # On le reconstruit proprement à partir de ce qui est déjà calculé

                familles = sorted(
                    set(
                        bg["Matériel CCC"].dropna().astype(str)
                        if "Matériel CCC" in bg.columns
                        else []
                    )
                    | set(
                        src["Nom de l'élément"].dropna().astype(str)
                        if "Nom de l'élément" in src.columns
                        else []
                    )
                )

                df_final = pd.DataFrame({"Famille": familles})

                for v, dv in data_versions.items():

                    qty_col = f"Quantité_{v}"
                    flag_col = f"Stocké en CCC ? ({v})"

                    if v == "V0" or not dv["with_ccc"]:
                        if {"Désignation", "Production", "Terminaux"}.issubset(dv["bg"].columns):
                            df_qty = (
                                dv["bg"][["Désignation", "Production", "Terminaux"]]
                                .dropna(subset=["Désignation"])
                                .assign(
                                    Quantité=lambda x: x["Production"].fillna(0)
                                    + x["Terminaux"].fillna(0)
                                )
                                .groupby("Désignation", as_index=False)["Quantité"]
                                .sum()
                                .rename(columns={"Désignation": "Famille"})
                            )
                        else:
                            df_qty = pd.DataFrame(columns=["Famille", "Quantité"])
                    else:
                        if {"Matériel CCC", "Nombre de matériels CCC"}.issubset(dv["bg"].columns):
                            df_qty = (
                                dv["bg"][["Matériel CCC", "Nombre de matériels CCC"]]
                                .dropna(subset=["Matériel CCC"])
                                .groupby("Matériel CCC", as_index=False)["Nombre de matériels CCC"]
                                .sum()
                                .rename(columns={
                                    "Matériel CCC": "Famille",
                                    "Nombre de matériels CCC": "Quantité"
                                })
                            )
                        else:
                            df_qty = pd.DataFrame(columns=["Famille", "Quantité"])

                    df_final = df_final.merge(df_qty, on="Famille", how="left")
                    df_final.rename(columns={"Quantité": qty_col}, inplace=True)
                    df_final[qty_col] = df_final[qty_col].fillna(0).astype(int)
                    df_final[flag_col] = df_final[qty_col].apply(lambda x: "✔️" if x > 0 else "❌")

                # 🔥 CETTE LIGNE EST OBLIGATOIRE
                st.dataframe(df_final, use_container_width=True)


                                                



            # ======================= ONGLET PALETTES =======================
            with ong_pal_comp:
                st.markdown("### 📦 Comparaison des palettes")

                # Totaux par version
                cols_tot = st.columns(len(selected_versions))
                for i, v in enumerate(selected_versions):
                    dv = data_versions[v]
                    with cols_tot[i]:
                        st.metric(
                            f"Palettes totales – {v}",
                            f"{dv['total_palettes']:,.0f}".replace(",", " "),
                        )

                # Palettes par étage / zone
                pal_list = [
                    dv["palettes_zone"].assign(Version=v)
                    for v, dv in data_versions.items()
                    if not dv["palettes_zone"].empty
                ]
                df_pal = pd.concat(pal_list, ignore_index=True) if pal_list else pd.DataFrame()

                if not df_pal.empty:
                    if "Lot" in df_pal.columns:
                        df_pal = df_pal.copy()
                        df_pal["Lot - Étage - Zone"] = df_pal["Lot"].astype(str) + " | " + df_pal["Étage - Zone"].astype(str)
                        x_pal_col = "Lot - Étage - Zone"
                    else:
                        x_pal_col = "Étage - Zone"
                    fig_pal = px.bar(
                        df_pal,
                        x=x_pal_col,
                        y="Palettes",
                        color="Version",
                        barmode="group",
                        title="Palettes par étage / zone",
                    )
                    st.plotly_chart(fig_pal, use_container_width=True)
                else:
                    st.info("Aucune donnée palettes pour ces versions.")

                # Flux palettes
                flux_pal_list = [
                    dv["flux_palettes"].assign(Version=v)
                    for v, dv in data_versions.items()
                    if not dv["flux_palettes"].empty
                ]
                df_flux_pal = pd.concat(flux_pal_list, ignore_index=True) if flux_pal_list else pd.DataFrame()

                if not df_flux_pal.empty:
                    if "Lot" in df_flux_pal.columns:
                        df_flux_pal = df_flux_pal.copy()
                        df_flux_pal["Version_Lot"] = df_flux_pal["Version"].astype(str) + " | " + df_flux_pal["Lot"].astype(str)
                        color_col = "Version_Lot"
                    else:
                        color_col = "Version"
                    fig_flux_pal = px.line(
                        df_flux_pal,
                        x="Mois",
                        y="Volume (nombre de palettes équivalentes)",
                        color=color_col,
                        title="Flux mensuel de palettes",
                    )
                    st.plotly_chart(fig_flux_pal, use_container_width=True)
                else:
                    st.info("Aucun flux palettes pour ces versions.")

            # ======================= ONGLET CAMIONS =======================
            with ong_cam_comp:
                st.markdown("### 🚚 Comparaison des camions")

                # Totaux + remplissage
                cols_cam = st.columns(len(selected_versions))
                for i, v in enumerate(selected_versions):
                    dv = data_versions[v]
                    with cols_cam[i]:
                        st.metric(
                            f"Camions totaux – {v}",
                            f"{dv['total_camions']:,.0f}".replace(",", " "),
                        )
                        st.metric(
                            f"Remplissage moyen – {v}",
                            f"{dv['rempl_moyen']:.1f} %",
                        )

                st.markdown("---")
                st.markdown("### 🚚 Camions par étage / zone")

                cam_list = [
                    dv["camions_zone"].assign(Version=v)
                    for v, dv in data_versions.items()
                    if not dv["camions_zone"].empty
                ]
                df_cam = pd.concat(cam_list, ignore_index=True) if cam_list else pd.DataFrame()

                if not df_cam.empty:
                    if "Lot" in df_cam.columns:
                        df_cam = df_cam.copy()
                        df_cam["Lot - Étage - Zone"] = df_cam["Lot"].astype(str) + " | " + df_cam["Étage - Zone"].astype(str)
                        x_cam_col = "Lot - Étage - Zone"
                    else:
                        x_cam_col = "Étage - Zone"
                    fig_cam = px.bar(
                        df_cam,
                        x=x_cam_col,
                        y="Camions",
                        color="Version",
                        barmode="group",
                        title="Camions par étage / zone",
                    )
                    st.plotly_chart(fig_cam, use_container_width=True)
                else:
                    st.info("Aucune donnée camions par étage pour ces versions.")

                st.markdown("---")
                st.markdown("### 📈 Flux mensuel de camions")

                flux_cam_list = [
                    dv["flux_camions"].assign(Version=v)
                    for v, dv in data_versions.items()
                    if not dv["flux_camions"].empty
                ]
                df_flux_cam = pd.concat(flux_cam_list, ignore_index=True) if flux_cam_list else pd.DataFrame()

                if not df_flux_cam.empty:
                    if "Lot" in df_flux_cam.columns:
                        df_flux_cam = df_flux_cam.copy()
                        df_flux_cam["Version_Lot"] = df_flux_cam["Version"].astype(str) + " | " + df_flux_cam["Lot"].astype(str)
                        color_cam = "Version_Lot"
                    else:
                        color_cam = "Version"
                    fig_flux_cam = px.line(
                        df_flux_cam,
                        x="Mois",
                        y="Camions",
                        color=color_cam,
                        title="Flux mensuel de camions",
                    )
                    st.plotly_chart(fig_flux_cam, use_container_width=True)
                else:
                    st.info("Aucun flux camions pour ces versions.")

                st.markdown("---")
                st.markdown("### 📦 Remplissage des camions par étage / zone")

                rempl_list = [
                    dv["rempl_zone"].assign(Version=v)
                    for v, dv in data_versions.items()
                    if not dv["rempl_zone"].empty
                ]
                df_rempl = pd.concat(rempl_list, ignore_index=True) if rempl_list else pd.DataFrame()

                if not df_rempl.empty:
                    if "Lot" in df_rempl.columns:
                        df_rempl = df_rempl.copy()
                        df_rempl["Lot - Étage - Zone"] = df_rempl["Lot"].astype(str) + " | " + df_rempl["Étage - Zone"].astype(str)
                        x_rempl_col = "Lot - Étage - Zone"
                    else:
                        x_rempl_col = "Étage - Zone"
                    fig_rempl = px.bar(
                        df_rempl,
                        x=x_rempl_col,
                        y="Remplissage (%)",
                        color="Version",
                        barmode="group",
                        title="Remplissage des camions par étage / zone",
                    )
                    st.plotly_chart(fig_rempl, use_container_width=True)
                else:
                    st.info("Aucune donnée de remplissage disponible pour ces versions.")













# Onglet 4 : Entraînement modèles :
elif menu == "Entraînement modèles":
    st.header("Entraînement des modèles")
    st.markdown("""
    ## Entraînement modèles
    
    Cette section permet d'entraîner les modèles utilisés pour classer les éléments dans les bordereaux.

    **Étape 1** : Préparer le fichier d'entraînement  
    - Téléchargez le fichier `template_entraînement.xlsx` ci-dessous.  
    - Remplissez les colonnes :
        - **Désignation** : Liste des matériels utilisés dans les bordereaux.  
        - **Catégorie Prédite** : Catégorie correspondante à chaque matériel.  

    **Étape 2** : Choisir le lot d'entraînement  
    **Étape 3** : Charger et lancer l'entraînement  
    """)

    # Liste des modèles disponibles
    models = [m.replace(".pkl", "") for m in os.listdir("models") if m.endswith(".pkl")]
    if "GLOBAL" not in models:
        models.insert(0, "GLOBAL")

    model_name = st.selectbox("Étape 2 : Choisir le lot d'entraînement", models)

    train_file = st.file_uploader("Déposez un fichier d’entraînement (xlsx)")

    if st.button("Lancer entraînement") and train_file:
        msg = entmod.train_model(train_file, model_name)
        st.success(msg)

# Onglet 5 : Base de données : 

elif menu == "Base de données":
    st.header("🗄️ Base de données SQLite")

    # Choix table + lot
    table_choice = st.selectbox("Choisir une table", ["Matériel", "Conditionnement", "Camion"])
    models = [m.replace(".pkl", "") for m in os.listdir("models") if m.endswith(".pkl")]
    if "GLOBAL" not in models:
        models.insert(0, "GLOBAL")
    lot_choice = st.selectbox("Modèle", models)

    # Lire les données du lot sélectionné
    df = daba.afficher_donnees(table_choice, lot_choice)

    st.subheader("📝 Modifier la table")

    df_modifie = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="crud_editor"
    )

    if st.button("💾 Enregistrer les modifications"):
        msg = daba.enregistrer_modifications(table_choice, lot_choice, df_modifie)
        st.success(msg)
        st.rerun()
