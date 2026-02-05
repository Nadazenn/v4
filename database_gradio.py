import sqlite3
import pandas as pd
import os

# Connexion SQLite
conn = sqlite3.connect("logistique.db", check_same_thread=False) 
cursor = conn.cursor()

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
        query = f"SELECT * FROM {table} WHERE lot = '{model_choice}' "
    else :
        query = f"SELECT * FROM {table}"
    df = pd.read_sql(query, conn)
    return df

# Fonction pour enregistrer les modifications
def enregistrer_modifications(table, df_modifie):
    table = nom_table (table)
    conn = sqlite3.connect("logistique.db", check_same_thread=False)
    
    # Vérifier la clé unique (ex: "id" ou "nom")
    df_base = pd.read_sql(f"SELECT * FROM {table}", conn)
    if "id" in df_base.columns:
        clé_unique = "id"
    elif "nom" in df_base.columns:
        clé_unique = "nom"
    else:
        return "Erreur : Impossible d'identifier une clé unique."
    if table == "materiel":
        df_modifie['quantite_par_conditionnement'] = pd.to_numeric(df_modifie['quantite_par_conditionnement'], errors='coerce')
        df_modifie['quantite_par_conditionnement_2'] = pd.to_numeric(df_modifie['quantite_par_conditionnement_2'], errors='coerce')
        df_modifie['quantite_par_conditionnement_3'] = pd.to_numeric(df_modifie['quantite_par_conditionnement_3'], errors='coerce')
    if table == "conditionnement":
        df_modifie['nombre_equiv_palettes'] = pd.to_numeric(df_modifie['nombre_equiv_palettes'], errors='coerce')
        df_modifie['masse_max'] = pd.to_numeric(df_modifie['masse_max'], errors='coerce')
    if table == "camion":
        df_modifie['capacite_palette'] = pd.to_numeric(df_modifie['capacite_palette'], errors='coerce')
        df_modifie['capacite_m3'] = pd.to_numeric(df_modifie['capacite_m3'], errors='coerce')
        df_modifie['capacite_kg'] = pd.to_numeric(df_modifie['capacite_kg'], errors='coerce')
        df_modifie['cout'] = pd.to_numeric(df_modifie['cout'], errors='coerce')
    # Supprimer les anciennes lignes qui existent déjà dans la base
    df_base = df_base[~df_base[clé_unique].isin(df_modifie[clé_unique])]

    # Ajouter les nouvelles données mises à jour
    df_final = pd.concat([df_base, df_modifie]).reset_index(drop=True)

    # Remplacer la table dans la base de données
    df_final.to_sql(table, conn, if_exists="replace", index=False)

    conn.commit()

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