import xlwings as xw
import pandas as pd
import unicodedata

def get_workbook(path="sortie/fichier_pilotage.xlsm"):
    # Excel reste visible pour que tu voies les modifs
    return xw.Book(path)

def lire_feuille(sheet_name, path="sortie/fichier_pilotage.xlsm"):
    wb = get_workbook(path)
    try:
        sht = wb.sheets[sheet_name]
    except:
        return pd.DataFrame()


    used = sht.used_range
    values = used.value

    if not values:
        return pd.DataFrame()

    df = pd.DataFrame(values)

    if sheet_name == "Données":
        # Forcer colonnes en 1, 2, 3 ...
        df.columns = [str(i+1) for i in range(df.shape[1])]
        #  Remplacer None par "" et tout convertir en string
        df = df.fillna("").astype(str)
    else:
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)
        df = df.fillna("")

    return df


def ecrire_feuille(sheet_name, df, path="sortie/fichier_pilotage.xlsm"):
    wb = get_workbook(path)
    sht = wb.sheets[sheet_name]
    sht.clear_contents()

    if sheet_name == "Données":
        # On écrit sans entêtes
        sht.range("A1").value = df.fillna("").values.tolist()
    else:
        # On écrit avec entêtes
        sht.range("A1").value = [df.columns.tolist()] + df.fillna("").values.tolist()

    wb.save()

def creer_tableau_source(path="sortie/fichier_pilotage.xlsm"):
    wb = get_workbook(path)
    wb.macro("CreerTableauSource")()
    wb.save()

def creer_bilan(path="sortie/fichier_pilotage.xlsm"):
    wb = get_workbook(path)
    wb.macro("CreerBilan")()
    wb.save()

def creer_bilan_zones(path="sortie/fichier_pilotage.xlsm"):
    try:
        wb = get_workbook(path)
        macro = wb.macro("CreerBilanZones")  # ⚠️ nom exact de la macro
        macro()
        wb.save()
        return "Macro 'CreerBilanZones' exécutée avec succès ✅"
    except Exception as e:
        return f"Erreur exécution macro : {e}"
    
def lancer_macro_bilan(path="sortie/fichier_pilotage.xlsm"):
    """
    Lance la macro VBA 'CreerBilanZones' qui génère Bilan + Graphiques + Livrable.
    """
    try:
        wb = get_workbook(path)
        macro = wb.macro("CreerBilanZones")  # ⚠️ nom exact de ta macro
        macro()
        wb.save()
        return "Macro 'CreerBilanZones' exécutée avec succès ✅"
    except Exception as e:
        return f"Erreur exécution macro : {e}"




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
    return None


def build_repartition_df(
    bordereau_df: pd.DataFrame,
    planning_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Construit le tableau de repartition Etage x Zone
    a partir du bordereau classe et du planning detaille.
    """
    cat_col = _find_col(bordereau_df.columns, "Categorie Predite")
    qty_col = _find_col(bordereau_df.columns, "Quantite")
    if not cat_col or not qty_col:
        raise ValueError(f"Colonnes manquantes dans bordereau: {list(bordereau_df.columns)}")

    etage_col = _find_col(planning_df.columns, "Numero etage (pas de lettres)")
    zone_col = _find_col(planning_df.columns, "Nom Zone")
    if not etage_col or not zone_col:
        raise ValueError(f"Colonnes manquantes dans planning: {list(planning_df.columns)}")

    grouped = (
        bordereau_df
        .groupby(cat_col, as_index=False)[qty_col]
        .sum()
    )
    grouped[qty_col] = pd.to_numeric(
        grouped[qty_col], errors="coerce"
    ).fillna(0).astype(int)

    repart_cols = []
    for _, row in planning_df.iterrows():
        etage = int(row[etage_col])
        zone = str(row[zone_col])
        repart_cols.append(f"E{etage}_Z{zone}")

    df = grouped.copy()
    for c in repart_cols:
        df[c] = 0

    df["Exclus"] = 0
    df["Commentaires"] = ""

    n = len(repart_cols)
    for i, r in df.iterrows():
        q = int(r[qty_col])
        cat = str(r[cat_col]).lower()

        if cat == "exclus":
            df.at[i, "Exclus"] = q
            continue

        if n == 0:
            continue

        base = q // n
        reste = q - base * n

        for c in repart_cols:
            df.at[i, c] = base

        if reste > 0:
            df.at[i, repart_cols[0]] += reste

    df["Reste"] = df[qty_col] - df[repart_cols].sum(axis=1)

    df = df.rename(columns={cat_col: "Categorie Predite", qty_col: "Quantite"})

    return df


def build_donnees_grid(
    bordereau_df: pd.DataFrame,
    planning_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Construit une version "feuille Donnees" pour l'affichage Streamlit,
    avec en-tetes sur plusieurs lignes et colonnes numerotees.
    """
    repart_df = build_repartition_df(bordereau_df, planning_df)

    etage_col = _find_col(planning_df.columns, "Numero etage (pas de lettres)")
    zone_col = _find_col(planning_df.columns, "Nom Zone")
    if not etage_col or not zone_col:
        raise ValueError(f"Colonnes manquantes dans planning: {list(planning_df.columns)}")

    etages = planning_df[etage_col].tolist()
    zones = planning_df[zone_col].tolist()
    nbzones = len(etages)

    # Colonnes numerotees (meme rendu que l'import Excel)
    n_cols = 4 + nbzones + 2  # col4 = Reste, + zones, + Exclus + Commentaires
    columns = [str(i) for i in range(1, n_cols + 1)]
    rows = []

    # Ligne 1 : titre + Etage + en-tetes zones
    row0 = [""] * n_cols
    row0[1] = "Repartition faite de maniere proportionnelle"
    row0[3] = "Etage"
    for i, etage in enumerate(etages):
        row0[4 + i] = etage
    row0[4 + nbzones] = "Exclus"
    row0[5 + nbzones] = "Commentaires"
    rows.append(row0)

    # Ligne 2 : Categorie / Quantite / Zone + zones
    row1 = [""] * n_cols
    row1[1] = "Categorie"
    row1[2] = "Quantite"
    row1[3] = "Zone"
    for i, zone in enumerate(zones):
        row1[4 + i] = zone
    rows.append(row1)

    # Ligne 3 : libelle "Repartition" (colonne reste)
    row2 = [""] * n_cols
    row2[3] = "Repartition"
    rows.append(row2)

    # Lignes donnees
    zone_cols = []
    for etage, zone in zip(etages, zones):
        zone_cols.append(f"E{int(etage)}_Z{zone}")

    for _, r in repart_df.iterrows():
        row = [""] * n_cols
        row[1] = r.get("Categorie Predite", "")
        row[2] = r.get("Quantite", "")
        row[3] = r.get("Reste", "")
        for i, col in enumerate(zone_cols):
            if col in repart_df.columns:
                row[4 + i] = r.get(col, "")
        row[4 + nbzones] = r.get("Exclus", "")
        row[5 + nbzones] = r.get("Commentaires", "")
        rows.append(row)

    return pd.DataFrame(rows, columns=columns)
