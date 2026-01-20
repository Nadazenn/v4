import sqlite3
import pandas as pd

conn = sqlite3.connect("logistique.db", check_same_thread=False)
cursor = conn.cursor()


def nom_table(table):
    if table == "Matériel":
        return "materiel"
    elif table == "Conditionnement":
        return "conditionnement"
    else:
        return "camion"


def afficher_donnees(table, lot_choice):
    table_sql = nom_table(table)

    if table_sql == "materiel":
        if lot_choice == "GLOBAL":
            query = f"SELECT * FROM {table_sql}"
        else:
            query = f"SELECT * FROM {table_sql} WHERE lot = '{lot_choice}'"
    else:
        query = f"SELECT * FROM {table_sql}"

    return pd.read_sql(query, conn)


def enregistrer_modifications(table, lot_choice, df_modifie):
    table_sql = nom_table(table)
    conn_local = sqlite3.connect("logistique.db", check_same_thread=False)

    # Lire toute la base SQL
    df_base = pd.read_sql(f"SELECT * FROM {table_sql}", conn_local)

    # Nettoyer les colonnes internes Streamlit
    df_modifie = df_modifie.drop(columns=[c for c in df_modifie.columns if c.startswith("_")], errors="ignore")
    df_modifie = df_modifie.dropna(how="all")

    # Forcer la colonne lot pour materiel
    if table_sql == "materiel":
        if "lot" in df_modifie.columns:
            df_modifie["lot"] = df_modifie["lot"].fillna(lot_choice)
            df_modifie.loc[df_modifie["lot"] == "", "lot"] = lot_choice

        df_other_lots = df_base[df_base["lot"] != lot_choice].copy()
    else:
        df_other_lots = pd.DataFrame()

    # Gestion IDs
    if "id" in df_modifie.columns:
        df_modifie["id"] = pd.to_numeric(df_modifie["id"], errors="coerce")

        # Lignes à ajouter
        df_add = df_modifie[df_modifie["id"].isna()].copy()

        if not df_add.empty:
            # Sécuriser max_id (float -> int)
            max_id = int(df_base["id"].max()) if not df_base.empty else 0

            # Générer de nouveaux IDs entiers
            next_ids = list(range(max_id + 1, max_id + 1 + len(df_add)))
            df_add["id"] = next_ids

        # Lignes modifiées existantes
        df_update = df_modifie[df_modifie["id"].notna()].copy()

    # Combine
    if table_sql == "materiel":
        df_final = pd.concat([df_other_lots, df_add, df_update], ignore_index=True)
    else:
        df_final = pd.concat([df_add, df_update], ignore_index=True)

    df_final = df_final.drop_duplicates(subset=["id"]).reset_index(drop=True)

    # Écriture SQL
    df_final.to_sql(table_sql, conn_local, if_exists="replace", index=False)

    conn_local.commit()
    conn_local.close()

    return "✔️ Base mise à jour avec succès."


cursor.execute("SELECT nom FROM camion ORDER BY id ASC")
liste_camions = [x[0] for x in cursor.fetchall()]

cursor.execute("SELECT nom FROM conditionnement ORDER BY id ASC")
liste_conditionnement = [x[0] for x in cursor.fetchall()]
def ajouter_supportage(materiels_df, model_choice):
    """
    Ajoute une ligne 'Supportage' au DataFrame pour chaque matériel nécessitant un supportage.

    :param materiels_df: DataFrame contenant les catégories et quantités de matériels.
    :param conn: Connexion SQLite à la base de données.
    :param model_choice: Valeur de filtre pour la colonne 'lot' dans la base SQLite.
    :return: DataFrame mis à jour avec les entrées "Supportage".
    """
    
    # Récupérer les matériels ayant 'Oui' dans la colonne 'supportage' et correspondant au lot choisi
    query = f"SELECT nom FROM materiel WHERE supportage = 'Oui' AND lot = '{model_choice}'"
    materiels_supportage = pd.read_sql_query(query, conn)['nom'].tolist()

    # Calculer la quantité totale de supportage (somme de 50% des quantités des matériels concernés)
    quantite_supportage_totale = materiels_df[materiels_df['Catégorie Prédite'].isin(materiels_supportage)]['Quantité'].sum() * 0.03

    # Ajouter UNE SEULE ligne "Supportage" si la quantité est positive
    if quantite_supportage_totale > 0:
        supportage_row = pd.DataFrame([{'Catégorie Prédite': 'Supportage', 'Quantité': quantite_supportage_totale}])
        materiels_df = pd.concat([materiels_df, supportage_row], ignore_index=True)

    return materiels_df