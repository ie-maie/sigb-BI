import ast
import pandas as pd
import numpy as np

def export_for_sql(df: pd.DataFrame):
    # 1. FIXED: Populate mandatory 'libelle' column to prevent MySQL rejection of constraints
    langue = df[["language"]].dropna().drop_duplicates().copy()
    langue.columns = ["code_langue"]
    langue["libelle"] = langue["code_langue"].map({"fre": "Francais", "ara": "Arabe"}).fillna("Inconnu")

    editeur = df[["Edition"]].dropna().drop_duplicates()
    editeur = editeur[editeur["Edition"].astype(str).str.strip() != ""].copy()
    editeur.columns = ["nom_editeur"]

    classification = df[["Cote"]].dropna().drop_duplicates()
    classification = classification[classification["Cote"].astype(str).str.strip() != ""].copy()
    classification.columns = ["cote"]

    # 2. BRIDGE DATA CONTAINERS
    exploded_authors = []
    notice_author_bridge = []
    exploded_matieres = []
    notice_matiere_bridge = []

    # 3. SINGLE PASS EXTRACTION (Robust Parsing)
    for _, row in df.iterrows():
        n_key = str(row.get("notice_key", "")).strip()

        # --- Author Parsing ---
        raw_authors = row.get("Auteur_Pairs")
        author_pairs_list = []

        if isinstance(raw_authors, list):
            author_pairs_list = raw_authors
        elif isinstance(raw_authors, np.ndarray):
            author_pairs_list = raw_authors.tolist()
        elif isinstance(raw_authors, str):
            try:
                evaluated = ast.literal_eval(raw_authors)
                author_pairs_list = evaluated if isinstance(evaluated, list) else []
            except: pass

        if not author_pairs_list:
            fallback = str(row.get("Auteur_Names", "")).strip()
            if fallback and fallback.lower() not in ["nan", "unknown"]:
                author_pairs_list = [(fallback, "Auteur")]

        for name, role in author_pairs_list:
            clean_name = str(name).strip()
            if clean_name and clean_name.lower() not in ["nan", "unknown"]:
                exploded_authors.append(clean_name)
                notice_author_bridge.append({"notice_key": n_key, "nom_complet": clean_name})

        # --- Matiere Parsing ---
        raw_mat = str(row.get("Matiere", ""))
        if raw_mat and raw_mat.lower() != "nan":
            subjects = [s.strip() for s in raw_mat.split(',') if s.strip()]
            for sub in subjects:
                exploded_matieres.append(sub)
                notice_matiere_bridge.append({"notice_key": n_key, "libelle": sub})

    auteur = pd.DataFrame(exploded_authors, columns=["nom_complet"]).drop_duplicates().reset_index(drop=True)
    matiere = pd.DataFrame(exploded_matieres, columns=["libelle"]).drop_duplicates().reset_index(drop=True)
    notice_author = pd.DataFrame(notice_author_bridge).drop_duplicates().reset_index(drop=True)
    notice_matiere = pd.DataFrame(notice_matiere_bridge).drop_duplicates().reset_index(drop=True)

    # 4. FIXED: Drop structural notice duplicates. Notice entities MUST be unique across key index.
    notice = df[["notice_key", "Titre", "Annee", "Nb_pages", "Lieu", "Edition", "language", "Cote"]].copy()
    notice.rename(columns={
        "Annee": "annee_pub", "Nb_pages": "nb_pages", "Lieu": "lieu_edition",
        "Edition": "nom_editeur", "language": "code_langue", "Cote": "cote"
    }, inplace=True)
    notice = notice.drop_duplicates(subset=["notice_key"]).reset_index(drop=True)

    # 5. EXEMPLAIRE (Keeps multiple real physical inventories tied to single notice keys)
    exemplaire = df[["Inventaire", "notice_key"]].copy()
    exemplaire.rename(columns={"Inventaire": "num_inventaire"}, inplace=True)
    exemplaire = exemplaire.dropna(subset=["num_inventaire"])

    return {
        "langue": langue,
        "editeur": editeur,
        "auteur": auteur,
        "classification": classification,
        "matiere": matiere,
        "notice": notice,
        "notice_author": notice_author,
        "notice_matiere": notice_matiere,
        "exemplaire": exemplaire
    }