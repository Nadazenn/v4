import pandas as pd
import io
import joblib
import os
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import LabelEncoder
from sklearn.linear_model import SGDClassifier
import numpy as np


def train_model(file_path, model_choice):
    """
    file_path : fichier uploadé Streamlit (BytesIO)
    model_choice : nom du modèle à entraîner
    """

    #
    # 1. Chargement du fichier Excel
    #
    try:
        # Lecture compatible Streamlit
        if hasattr(file_path, "read"):
            df = pd.read_excel(io.BytesIO(file_path.read()), engine="openpyxl")
        else:
            df = pd.read_excel(file_path, engine="openpyxl")

    except Exception as e:
        return f"Erreur lors du chargement des données : {e}"

    # Vérification colonnes obligatoires
    required_cols = ["Désignation", "Catégorie Prédite"]
    if not all(col in df.columns for col in required_cols):
        return f"Erreur : le fichier doit contenir les colonnes {required_cols}"

    # Extraction des données
    X = df["Désignation"].astype(str)
    y = df["Catégorie Prédite"].astype(str)

    #
    # 2. Chargement ou création du modèle
    #
    model_name = f"{model_choice}.pkl"
    model_path = os.path.join("models", model_name)

    try:
        model, vectorizer, label_encoder = joblib.load(model_path)
        print("Modèle existant chargé.")
    except FileNotFoundError:
        model = SGDClassifier(loss="log_loss", random_state=42)
        vectorizer = TfidfVectorizer()
        label_encoder = LabelEncoder()
        print("Nouveau modèle créé.")

    #
    # 3. Encodage des catégories
    #
    y_encoded = label_encoder.fit_transform(y)

    #
    # 4. Transformation du texte
    #
    if getattr(vectorizer, "vocabulary_", None) is not None:
        # vectorizer déjà entraîné → transformation simple
        X_transformed = vectorizer.transform(X)
    else:
        # première utilisation → fit_transform
        X_transformed = vectorizer.fit_transform(X)

    #
    # 5. Gestion correcte des classes
    #
    existing_classes = model.classes_ if hasattr(model, "classes_") else np.array([])
    all_classes = np.union1d(existing_classes, np.unique(y_encoded))

    #
    # 6. Entraînement incrémental
    #
    model.partial_fit(X_transformed, y_encoded, classes=all_classes)

    #
    # 7. Sauvegarde propre
    #
    os.makedirs("models", exist_ok=True)
    joblib.dump((model, vectorizer, label_encoder), model_path)

    return f"Modèle {model_choice} mis à jour et sauvegardé."
