import pandas as pd
import string


# ==========================================================
#  Générer le premier tableau (étages / zones)
# ==========================================================
def generate_table(nombre_etages, zones_par_etage_defaut, numero_etage_inf):
    """Crée le tableau des étages et du nombre de zones par étage."""
    num_etages = [numero_etage_inf + i for i in range(nombre_etages)]
    zones = [zones_par_etage_defaut] * nombre_etages

    df = pd.DataFrame({
        "Étage": [f"Étage {i + 1}" for i in range(nombre_etages)],
        "Numéro étage (pas de lettres)": num_etages,
        "Nombre de zones": zones
    })
    return df


# ==========================================================
# Générer le tableau détaillé (planning complet)
# ==========================================================
def generate_details_table(
    etages_zones,
    zones_per_etage,
    delai_livraison,
    date_debut_prod,
    date_debut_term,
    délai,
    duree_prodmoyen_paretage,
    duree_termmoyen_paretage
):
    """Crée le tableau détaillé du planning (production et terminaux)."""
    rows = []
    current_date_prod = pd.to_datetime(date_debut_prod, dayfirst=True)
    current_date_term = pd.to_datetime(date_debut_term, dayfirst=True)

    for etage, zones in zip(etages_zones, zones_per_etage):
        for i in range(zones):
            zone_letter = string.ascii_uppercase[i % 26]
            row = [
                etage,
                f"Zone {i + 1}",
                zone_letter,
                current_date_prod.strftime("%d/%m/%Y"),
                current_date_term.strftime("%d/%m/%Y"),
                delai_livraison,
                duree_prodmoyen_paretage,
                duree_termmoyen_paretage
            ]
            rows.append(row)
            current_date_prod += pd.Timedelta(days=délai)
            current_date_term += pd.Timedelta(days=délai)

    columns = [
        "Étage",
        "Zone",
        "Nom Zone",
        "Date début phase production",
        "Date début phase terminaux",
        "Délai de livraison avant travaux (jours)",
        "Durée travaux production",
        "Durée travaux terminaux"
    ]

    df = pd.DataFrame(rows, columns=columns)
    df = df.fillna("")
    return df


# ==========================================================
#  Validation du paramétrage
# ==========================================================
def validate_parametrage():
    """Retourne un message de validation pour l'étape suivante."""
    return True, "Le paramétrage est terminé, vous pouvez passer à l'onglet Données."
