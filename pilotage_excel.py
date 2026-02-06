
import pandas as pd
import unicodedata
import database as daba





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
    planning_df: pd.DataFrame,
    lot: str = ""
) -> pd.DataFrame:
    """
    Construit le tableau de repartition Etage x Zone
    a partir du bordereau classe et du planning detaille.
    """
    cat_col = _find_col(bordereau_df.columns, "Catégorie Prédite")
    qty_col = _find_col(bordereau_df.columns, "Quantité")
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
    grouped = grouped.rename(
        columns={cat_col: "Catégorie Prédite", qty_col: "Quantité"}
    )
    grouped["Quantité"] = pd.to_numeric(
        grouped["Quantité"], errors="coerce"
    ).fillna(0).astype(int)
    grouped = daba.ajouter_supportage(grouped, lot)

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
        q = int(r["Quantité"])
        cat = str(r["Catégorie Prédite"]).lower()

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

    df["Reste"] = df["Quantité"] - df[repart_cols].sum(axis=1)

    df = df.rename(columns={"Catégorie Prédite": "Categorie Predite", "Quantité": "Quantite"})

    return df


def build_donnees_grid(
    bordereau_df: pd.DataFrame,
    planning_df: pd.DataFrame,
    lot: str = ""
) -> pd.DataFrame:
    """
    Construit une version "feuille Donnees" pour l'affichage Streamlit,
    avec en-tetes sur plusieurs lignes et colonnes numerotees.
    """
    repart_df = build_repartition_df(bordereau_df, planning_df, lot)

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
    row1[1] = "Catégorie"
    row1[2] = "Quantité"
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


def build_tableau_source(
    bordereau_df: pd.DataFrame,
    donnees_grid: pd.DataFrame,
    lot: str = ""
) -> pd.DataFrame:
    """
    Construit le Tableau Source en suivant EXACTEMENT la logique du Module1.bas VBA.
    
    Logique:
    - Parcourir chaque colonne de zone (col 5+) du donnees_grid
    - Pour chaque ligne de matériel (row 4+), récupérer quantité à l'intersection
    - Si quantité > 0, créer une ligne Tableau Source avec les infos matériel
    """
    import sqlite3
    import math

    colonnes = [
        "Etage",
        "Zone",
        "Lot",
        "Phase de traveaux",
        "Nom de l'élément",
        "Unité",
        "Quantité",
        "Conditionnement",
        "Quantité par UM",
        "Nombre d'UM nécessaires",
        "Nombre palettes equivalent total",
        "Type de camion requis",
        "Nombre de camions nécessaires",
        "Dont camions pleins",
        "Remplissage camion non plein",
        "Utilisation d'une CCC"
    ]
    
    rows = []
    
    try:
        # Charger la base de données
        conn = sqlite3.connect('logistique.db')
        cur = conn.cursor()

        # Récupérer tables utiles
        cur.execute("SELECT * FROM materiel WHERE lot = ?", (lot,))
        materiel_rows = cur.fetchall()
        # fallback: si aucun materiel pour ce lot, charger tout pour diagnostic
        if not materiel_rows:
            print(f"build_tableau_source: aucun matériel trouvé pour lot='{lot}', fallback sur tous les matériels")
            cur.execute("SELECT * FROM materiel")
            materiel_rows = cur.fetchall()

        cur.execute("SELECT * FROM conditionnement")
        cond_rows = cur.fetchall()
        cur.execute("SELECT * FROM camion")
        cam_rows = cur.fetchall()

        # Construire dictionnaire conditionnement: nom -> (nombre_equiv_palettes, type_camion)
        cond_dict = {}
        # conditionnement table columns: id, nom, type_camion, nombre_equiv_palettes, ...
        for r in cond_rows:
            try:
                nb = float(r[3]) if r[3] not in (None, '') else 1.0
            except:
                nb = 1.0
            cond_dict[str(r[1]).strip().lower()] = {
                'nom': r[1],
                'type_camion': r[2],
                'nb_pal_eq': nb
            }

        # Construire liste camions (dict) par type
        cam_by_type = {}
        # camion columns: id, nom, type_camion, capacite_palette, ...
        for r in cam_rows:
            nom = r[1]
            typ = r[2]
            cap = r[3] if r[3] is not None else 0
            cam_by_type.setdefault(typ, []).append({'nom': nom, 'capacite': cap})

        # Construire dict materiel
        materiel_dict = {}
        # helper to normalize strings (remove accents, lower)
        def _norm(s):
            import unicodedata
            if s is None:
                return ""
            v = str(s)
            v = unicodedata.normalize('NFKD', v)
            v = ''.join(c for c in v if not unicodedata.combining(c))
            return v.lower().strip()

        for r in materiel_rows:
            # schema: id, nom, lot, unite, phase_travaux, utilisation_ccc, supportage,
            # conditionnement_defaut, quantite_par_conditionnement,
            # autre_conditionnement_2, quantite_par_conditionnement_2,
            # autre_conditionnement_3, quantite_par_conditionnement_3, commentaire
            key = _norm(r[1])
            materiel_dict[key] = {
                'id': r[0], 'nom': r[1], 'lot': r[2], 'unite': r[3],
                'phase': r[4], 'util_ccc': r[5], 'supportage': r[6],
                'cond1': r[7], 'cond1_qty': r[8],
                'cond2': r[9], 'cond2_qty': r[10],
                'cond3': r[11], 'cond3_qty': r[12],
            }

        # Helper: picker camion (portage of OptimiserRemplissage)
        def pick_camion(nb_palettes, type_cam):
            candidates = cam_by_type.get(type_cam, [])
            if not candidates:
                return None
            best = None
            meilleurNbCamions = 999999
            meilleurTaux = -1
            for c in candidates:
                cap = c['capacite'] if c['capacite'] else 1
                nbCamions = max(1, math.ceil(nb_palettes / cap))
                tauxRemplissage = (nb_palettes / (nbCamions * cap)) * 100 if nbCamions * cap > 0 else 0
                palettesDernier = nb_palettes % cap
                if palettesDernier == 0:
                    tauxDernier = 100
                else:
                    tauxDernier = (palettesDernier / cap) * 100

                if nb_palettes == 1:
                    if tauxRemplissage > meilleurTaux:
                        best = c
                        meilleurTaux = tauxRemplissage
                else:
                    if (nbCamions < meilleurNbCamions) or (nbCamions == meilleurNbCamions and tauxRemplissage > meilleurTaux and tauxDernier > 50):
                        best = c
                        meilleurNbCamions = nbCamions
                        meilleurTaux = tauxRemplissage

            return best

        # Parcourir comme en VBA : colonnes j 5..lastColumn-2 (ici index 4..)
        # quick diagnostics
        print(f"build_tableau_source: materiel_rows={len(materiel_rows)}, cond_rows={len(cond_rows)}, cam_rows={len(cam_rows)}")
        print(f"donnees_grid shape={donnees_grid.shape}")

        # VBA: j 5 To lastColumnDonnees - 2 (skip "Exclus" + "Commentaires")
        last_col = max(0, len(donnees_grid.columns) - 2)
        for col_idx in range(4, last_col):
            etage = donnees_grid.iloc[0, col_idx] if col_idx < len(donnees_grid.columns) else ""
            zone = donnees_grid.iloc[1, col_idx] if col_idx < len(donnees_grid.columns) else ""

            for row_idx in range(3, len(donnees_grid)):
                categorie = donnees_grid.iloc[row_idx, 1]
                if pd.isna(categorie) or str(categorie).strip() == "":
                    continue
                # convert quantity robustly
                raw = donnees_grid.iloc[row_idx, col_idx]
                try:
                    if isinstance(raw, str):
                        raw = raw.replace(" ", "").replace(",", ".")
                    quantite = pd.to_numeric(raw, errors='coerce')
                    quantite = float(quantite) if not pd.isna(quantite) else 0
                except Exception:
                    quantite = 0
                if quantite <= 0:
                    continue

                key = _norm(categorie)
                mat = materiel_dict.get(key)
                if not mat:
                    continue

                # choisir conditionnement et sa quantite par UM (meme paire que VBA)
                cond_name = 'Palette'
                cond_qty = None
                for name, qty in [
                    (mat.get('cond1'), mat.get('cond1_qty')),
                    (mat.get('cond2'), mat.get('cond2_qty')),
                    (mat.get('cond3'), mat.get('cond3_qty')),
                ]:
                    if name is not None and str(name).strip() != "":
                        cond_name = name
                        cond_qty = qty
                        break
                # try convert cond_qty to number
                try:
                    cond_qty = float(cond_qty) if cond_qty not in (None, '', 'None') else None
                except:
                    cond_qty = None

                # Lookup conditionnement table for equivalence
                cond_lookup = cond_dict.get(_norm(cond_name), None)
                nb_pal_eq = cond_lookup['nb_pal_eq'] if cond_lookup else 1.0
                type_cam_req = cond_lookup['type_camion'] if cond_lookup else None

                # Calculs
                quantite_par_um = cond_qty if cond_qty and cond_qty > 0 else None
                nb_um = math.ceil(quantite / quantite_par_um) if quantite_par_um else None
                nb_palettes = (nb_um * nb_pal_eq) if nb_um else None

                # Choisir camion et calculs camions
                chosen = pick_camion(nb_palettes, type_cam_req)
                camion_nom = chosen['nom'] if chosen else None
                camion_cap = chosen['capacite'] if chosen else None
                if camion_cap and camion_cap > 0:
                    nb_camions = int(math.ceil(nb_palettes / camion_cap))
                    full_trucks = int(nb_palettes // camion_cap)
                    fill_last = round((nb_palettes / camion_cap) - full_trucks, 2)
                else:
                    nb_camions = None
                    full_trucks = None
                    fill_last = None

                new_row = [
                    etage,
                    zone,
                    lot,
                    mat.get('phase'),
                    mat.get('nom'),
                    mat.get('unite'),
                    int(quantite) if quantite == int(quantite) else quantite,
                    cond_name,
                    quantite_par_um,
                    nb_um,
                    nb_palettes,
                    camion_nom,
                    nb_camions,
                    full_trucks,
                    fill_last,
                    mat.get('util_ccc')
                ]
                rows.append(new_row)

        # Après avoir ajouté tous les matériaux, ajouter les lignes Stock CCC (Production / Terminaux) par zone
        # Construire DataFrame temporaire pour sommation
        df_rows = pd.DataFrame(rows, columns=colonnes) if rows else pd.DataFrame(columns=colonnes)
        print(f"build_tableau_source: created {len(rows)} material rows before CCC stocks")
        # Parcourir toutes les combinaisons d'etage/zone présentes
        if not df_rows.empty:
            zones = df_rows[['Etage', 'Zone']].drop_duplicates()
            for _, z in zones.iterrows():
                et, zo = z['Etage'], z['Zone']
                for phase in ['Production', 'Terminaux']:
                    mask = (
                        (df_rows['Etage'] == et)
                        & (df_rows['Zone'] == zo)
                        & (df_rows["Utilisation d'une CCC"] == 'Oui')
                        & (df_rows['Phase de traveaux'] == phase)
                    )
                    total_pal = df_rows.loc[mask, 'Nombre palettes equivalent total'].sum()

                    # Créer ligne stock (même logique camions que VBA)
                    cond = cond_dict.get(_norm("Palette"))
                    camion_nom = None
                    nb_camions = None
                    full_trucks = None
                    fill_last = None
                    if cond:
                        type_cam_req = cond.get("type_camion")
                        chosen = pick_camion(total_pal, type_cam_req)
                        camion_nom = chosen['nom'] if chosen else None
                        camion_cap = chosen['capacite'] if chosen else None
                        if camion_cap and camion_cap > 0:
                            nb_camions = int(math.ceil(total_pal / camion_cap))
                            full_trucks = int(total_pal // camion_cap)
                            fill_last = round((total_pal / camion_cap) - full_trucks, 2)

                    stock_row = [
                        et,
                        zo,
                        lot,
                        phase,
                        f"Stock CCC {phase}",
                        '',
                        '',
                        'Palette',
                        None,
                        None,
                        total_pal,
                        camion_nom,
                        nb_camions,
                        full_trucks,
                        fill_last,
                        ''
                    ]
                    rows.append(stock_row)

        return pd.DataFrame(rows, columns=colonnes)

    except Exception as e:
        print(f"Erreur dans build_tableau_source: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=colonnes)
