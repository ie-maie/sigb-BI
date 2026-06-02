# Projet SIGB - guide PyCharm

## 1. Ouvrir le projet

Dans PyCharm, ouvrir ce dossier :

```text
C:\Users\hp\Downloads\sigb-BI-main\sigb-BI-main
```

## 2. Installer les dependances

Dans le terminal de PyCharm :

```bash
python -m pip install -r requirements.txt
```

## 3. Verifier le fichier `.env`

Le fichier `.env` doit etre a la racine du projet :

```env
DB_HOST=84.8.219.151
DB_PORT=3306
DB_NAME=sigb
DB_USER=sigb_user
DB_PASSWORD=mot_de_passe_a_demander
```

Ne pas publier ce fichier sur GitHub.

## 4. Tester sans modifier la base

Lancer d'abord :

```bash
python python/test_connection.py
```

Puis :

```bash
python python/04_validate_counts.py
```

## 5. Attention aux scripts ETL

Ne relance pas `02_etl_csv.py` ou `03_etl_excel.py` sur la base du serveur si elle contient deja les donnees.
Ces scripts ajoutent des notices dans la base.

Pour tester l'ETL sans risque, utilise une base de test vide, puis execute `sql/01_create_schema.sql`.
