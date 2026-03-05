import sqlite3
import pandas as pd
import os

# Connexion SQLite
conn = sqlite3.connect("logistique.db", check_same_thread=False) 
cursor = conn.cursor()

QTY_COLUMNS_MATERIEL = [
    "quantite_par_conditionnement",
    "quantite_par_conditionnement_2",
    "quantite_par_conditionnement_3",
]


def _coerce_nullable_nonneg_float_columns(df: pd.DataFrame, columns) -> pd.DataFrame:
    for col in columns:
        if col in df.columns:
            df[col] = _to_numeric_loose(df[col]).clip(lower=0)
    return df


def _to_numeric_loose(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", ".", regex=False), errors="coerce")


def nom_table (table):
    if table == "Matériel":
        table = "materiel"
    elif table == "Conditionnement":
        table = "conditionnement"
    else:
        table = "camion"
    return (table)

# Fonction pour afficher les données
def afficher_donnees(table, model_choice):
    table = nom_table (table)
    if table == "materiel" :
        if str(model_choice).strip().upper() in {"GLOBAL", "TCE"}:
            query = f"SELECT * FROM {table}"
        else:
            query = f"SELECT * FROM {table} WHERE lot = '{model_choice}' "
    else :
        query = f"SELECT * FROM {table}"
    df = pd.read_sql(query, conn)
    if table == "materiel":
        df = _coerce_nullable_nonneg_float_columns(df, QTY_COLUMNS_MATERIEL)
    if table == "conditionnement":
        if "nombre_equiv_palettes" in df.columns:
            df["nombre_equiv_palettes"] = _to_numeric_loose(df["nombre_equiv_palettes"]).clip(lower=0)
        if "masse_max" in df.columns:
            df["masse_max"] = (
                _to_numeric_loose(df["masse_max"])
                .round()
                .clip(lower=0)
                .astype("Int64")
            )
    return df

# Fonction pour enregistrer les modifications
def enregistrer_modifications(table, df_modifie, lot_choice=None):
    table = nom_table (table)
    conn = sqlite3.connect("logistique.db", check_same_thread=False)
    
    # Charger l'état courant de la table
    df_base = pd.read_sql(f"SELECT * FROM {table}", conn)
    if table == "materiel":
        df_modifie = _coerce_nullable_nonneg_float_columns(df_modifie, QTY_COLUMNS_MATERIEL)
    if table == "conditionnement":
        df_modifie["nombre_equiv_palettes"] = (
            _to_numeric_loose(df_modifie["nombre_equiv_palettes"])
            .clip(lower=0)
        )
        df_modifie["masse_max"] = (
            _to_numeric_loose(df_modifie["masse_max"])
            .round()
            .clip(lower=0)
            .astype("Int64")
        )
    if table == "camion":
        df_modifie['capacite_palette'] = pd.to_numeric(df_modifie['capacite_palette'], errors='coerce')
        df_modifie['capacite_m3'] = pd.to_numeric(df_modifie['capacite_m3'], errors='coerce')
        df_modifie['capacite_kg'] = pd.to_numeric(df_modifie['capacite_kg'], errors='coerce')
        df_modifie['cout'] = pd.to_numeric(df_modifie['cout'], errors='coerce')
    # Nettoyage de base pour éviter l'insertion de lignes vides
    if "nom" in df_modifie.columns:
        df_modifie = df_modifie[df_modifie["nom"].astype(str).str.strip() != ""]

    # Générer un id pour les nouvelles lignes si besoin
    if "id" in df_base.columns and "id" in df_modifie.columns:
        df_base["id"] = pd.to_numeric(df_base["id"], errors="coerce")
        df_modifie["id"] = pd.to_numeric(df_modifie["id"], errors="coerce")
        max_id = int(df_base["id"].max()) if df_base["id"].notna().any() else 0
        missing_id = df_modifie["id"].isna()
        if missing_id.any():
            new_ids = list(range(max_id + 1, max_id + 1 + int(missing_id.sum())))
            df_modifie.loc[missing_id, "id"] = new_ids

    # Construire l'état final :
    # - materiel par lot : on remplace uniquement le lot chargé
    # - autres cas : on remplace tout ce qui était affiché (suppression persistante)
    if table == "materiel" and lot_choice and str(lot_choice).strip().upper() not in {"GLOBAL", "TCE"}:
        if "lot" in df_modifie.columns:
            df_modifie["lot"] = lot_choice
        else:
            df_modifie["lot"] = lot_choice
        df_final = pd.concat([df_base[df_base["lot"] != lot_choice], df_modifie], ignore_index=True)
    else:
        df_final = df_modifie.copy()

    # Garder l'ordre des colonnes SQL
    if not df_base.empty:
        for col in df_base.columns:
            if col not in df_final.columns:
                df_final[col] = None
        df_final = df_final[df_base.columns]

    # Conserver le schéma SQL (PK/autoincrement) en vidant puis en réinsérant
    conn.execute(f"DELETE FROM {table}")
    df_final.to_sql(table, conn, if_exists="append", index=False)

    conn.commit()
    conn.close()

    return f"La table {table} a été mise à jour avec succès."

cursor.execute("SELECT nom FROM camion ORDER BY id ASC")
liste_camions = [camion[0] for camion in cursor.fetchall()]  # Extraire uniquement les noms

cursor.execute("SELECT nom FROM conditionnement ORDER BY id ASC")
liste_conditionnement = [conditionnement[0] for conditionnement in cursor.fetchall()]  # Extraire uniquement les noms

def ajouter_supportage(materiels_df, model_choice):
    """
    Ajoute une ligne 'Supportage' au DataFrame pour chaque matériel nécessitant un supportage.

    :param materiels_df: DataFrame contenant les catégories et quantités de matériels.
    :param conn: Connexion SQLite à la base de données.
    :param model_choice: Valeur de filtre pour la colonne 'lot' dans la base SQLite.
    :return: DataFrame mis à jour avec les entrées "Supportage".
    """
    
    def _norm(s):
        return "" if s is None else str(s).strip().lower()

    df = materiels_df.copy()
    has_lot = "Lot" in df.columns

    supportage_df = pd.read_sql_query(
        "SELECT nom, lot FROM materiel WHERE supportage = 'Oui'",
        conn
    )
    supportage_df["nom_norm"] = supportage_df["nom"].map(_norm)
    supportage_df["lot_norm"] = supportage_df["lot"].map(_norm)

    cat_col = "Catégorie Prédite"
    qty_col = "Quantité"
    if cat_col not in df.columns or qty_col not in df.columns:
        return df

    df["_cat_norm"] = df[cat_col].map(_norm)
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)

    rows_to_add = []

    if has_lot:
        df["_lot_norm"] = df["Lot"].map(_norm)
        for lot_norm in sorted(df["_lot_norm"].dropna().unique().tolist()):
            noms_lot = set(
                supportage_df.loc[supportage_df["lot_norm"] == lot_norm, "nom_norm"].dropna().tolist()
            )
            if not noms_lot:
                continue
            q = df.loc[(df["_lot_norm"] == lot_norm) & (df["_cat_norm"].isin(noms_lot)), qty_col].sum() * 0.03
            if q > 0:
                lot_value = df.loc[df["_lot_norm"] == lot_norm, "Lot"].iloc[0]
                rows_to_add.append({"Lot": lot_value, cat_col: "Supportage", qty_col: q})
        if rows_to_add:
            df = pd.concat([df, pd.DataFrame(rows_to_add)], ignore_index=True)
    else:
        model_norm = _norm(model_choice)
        if model_norm in {"global", "tce"}:
            noms = set(supportage_df["nom_norm"].dropna().tolist())
        else:
            noms = set(
                supportage_df.loc[supportage_df["lot_norm"] == model_norm, "nom_norm"].dropna().tolist()
            )
        q = df.loc[df["_cat_norm"].isin(noms), qty_col].sum() * 0.03
        if q > 0:
            df = pd.concat([df, pd.DataFrame([{cat_col: "Supportage", qty_col: q}])], ignore_index=True)

    return df.drop(columns=[c for c in ["_cat_norm", "_lot_norm"] if c in df.columns])
