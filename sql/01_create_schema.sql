CREATE TABLE langue (
    id_langue INT AUTO_INCREMENT PRIMARY KEY,
    code_langue VARCHAR(10) UNIQUE NOT NULL,
    libelle VARCHAR(100) NOT NULL
);

CREATE TABLE editeur (
    id_editeur INT AUTO_INCREMENT PRIMARY KEY,
    nom_editeur VARCHAR(255) NOT NULL,
    UNIQUE(nom_editeur)
);

CREATE TABLE auteur (
    id_auteur INT AUTO_INCREMENT PRIMARY KEY,
    nom_complet VARCHAR(500) NOT NULL UNIQUE
);

CREATE TABLE classification (
    id_classification INT AUTO_INCREMENT PRIMARY KEY,
    cote VARCHAR(100) NOT NULL UNIQUE
);

CREATE TABLE matiere (
    id_matiere INT AUTO_INCREMENT PRIMARY KEY,
    libelle VARCHAR(500) NOT NULL UNIQUE
);

CREATE TABLE notice (
    id_notice INT AUTO_INCREMENT PRIMARY KEY,
    titre TEXT NOT NULL,
    annee_pub SMALLINT,
    nb_pages SMALLINT,
    lieu_edition VARCHAR(255),
    id_editeur INT,
    id_langue INT,
    id_classification INT,
    date_catalogage DATE DEFAULT CURRENT_DATE,
    note TEXT,
    FOREIGN KEY (id_editeur) REFERENCES editeur(id_editeur),
    FOREIGN KEY (id_langue) REFERENCES langue(id_langue),
    FOREIGN KEY (id_classification) REFERENCES classification(id_classification)
);

CREATE TABLE notice_auteur (
    id_notice INT,
    id_auteur INT,
    PRIMARY KEY(id_notice,id_auteur),
    FOREIGN KEY(id_notice) REFERENCES notice(id_notice) ON DELETE CASCADE,
    FOREIGN KEY(id_auteur) REFERENCES auteur(id_auteur)
);

CREATE TABLE notice_matiere (
    id_notice INT,
    id_matiere INT,
    PRIMARY KEY(id_notice,id_matiere),
    FOREIGN KEY(id_notice) REFERENCES notice(id_notice) ON DELETE CASCADE,
    FOREIGN KEY(id_matiere) REFERENCES matiere(id_matiere)
);

CREATE TABLE exemplaire (
    id_exemplaire INT AUTO_INCREMENT PRIMARY KEY,
    num_inventaire VARCHAR(50) UNIQUE,
    id_notice INT NOT NULL,
    etat VARCHAR(50) DEFAULT 'Disponible',
    FOREIGN KEY(id_notice) REFERENCES notice(id_notice)
);

-- ============================================================================
-- Add indexes for performance
-- ============================================================================
CREATE INDEX idx_notice_titre ON notice (titre(100));
CREATE INDEX idx_auteur_nom ON auteur (nom_complet(100));
CREATE INDEX idx_exemplaire_inv ON exemplaire (num_inventaire);

-- ============================================================================
-- MIGRATION COMPLETE
-- ============================================================================
