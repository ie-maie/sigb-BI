-- 04_validation.sql
-- Requetes de controle qualite pour le projet SIGB
-- A executer apres le chargement des donnees ETL.

-- 1. Nombre total de notices
SELECT COUNT(*) AS total_notices
FROM notice;

-- 2. Nombre total d'auteurs
SELECT COUNT(*) AS total_auteurs
FROM auteur;

-- 3. Nombre total d'editeurs
SELECT COUNT(*) AS total_editeurs
FROM editeur;

-- 4. Nombre total de matieres
SELECT COUNT(*) AS total_matieres
FROM matiere;

-- 5. Nombre total de classifications
SELECT COUNT(*) AS total_classifications
FROM classification;

-- 6. Nombre total d'exemplaires
SELECT COUNT(*) AS total_exemplaires
FROM exemplaire;

-- 7. Repartition des notices par langue
SELECT
    l.code_langue,
    l.libelle,
    COUNT(n.id_notice) AS total_notices
FROM langue l
LEFT JOIN notice n ON n.id_langue = l.id_langue
GROUP BY l.id_langue, l.code_langue, l.libelle
ORDER BY total_notices DESC;

-- 8. Notices sans auteur associe
SELECT COUNT(*) AS notices_sans_auteur
FROM notice n
LEFT JOIN notice_auteur na ON na.id_notice = n.id_notice
WHERE na.id_auteur IS NULL;

-- 9. Notices sans matiere associee
SELECT COUNT(*) AS notices_sans_matiere
FROM notice n
LEFT JOIN notice_matiere nm ON nm.id_notice = n.id_notice
WHERE nm.id_matiere IS NULL;

-- 10. Notices sans exemplaire associe
SELECT COUNT(*) AS notices_sans_exemplaire
FROM notice n
LEFT JOIN exemplaire e ON e.id_notice = n.id_notice
WHERE e.id_exemplaire IS NULL;

-- 11. Annees de publication suspectes
SELECT
    id_notice,
    titre,
    annee_pub
FROM notice
WHERE annee_pub IS NOT NULL
  AND (annee_pub < 1400 OR annee_pub > YEAR(CURRENT_DATE))
ORDER BY annee_pub;

-- 12. Doublons possibles de notices
SELECT
    titre,
    annee_pub,
    id_langue,
    COUNT(*) AS nombre_occurrences
FROM notice
GROUP BY titre, annee_pub, id_langue
HAVING COUNT(*) > 1
ORDER BY nombre_occurrences DESC, titre
LIMIT 50;
