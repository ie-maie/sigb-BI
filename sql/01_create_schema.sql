CREATE TABLE IF NOT EXISTS langue (
    id_langue     INT AUTO_INCREMENT PRIMARY KEY,
    code_langue   VARCHAR(10)  NOT NULL UNIQUE,
    libelle       VARCHAR(100) NOT NULL
);

CREATE TABLE IF NOT EXISTS editeur (
    id_editeur    INT AUTO_INCREMENT PRIMARY KEY,
    nom_editeur   VARCHAR(255) NOT NULL,
    ville         VARCHAR(150),
    UNIQUE (nom_editeur, ville)
);

CREATE TABLE IF NOT EXISTS auteur (
    id_auteur     INT AUTO_INCREMENT PRIMARY KEY,
    nom           VARCHAR(150) NOT NULL,
    prenom        VARCHAR(150),
    UNIQUE (nom, prenom)
);

CREATE TABLE IF NOT EXISTS classification (
    id_classification INT AUTO_INCREMENT PRIMARY KEY,
    cote              VARCHAR(100) NOT NULL UNIQUE,
    libelle           VARCHAR(500)
);

CREATE TABLE IF NOT EXISTS matiere (
    id_matiere    INT AUTO_INCREMENT PRIMARY KEY,
    libelle       VARCHAR(500) NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS notice (
    id_notice         INT AUTO_INCREMENT PRIMARY KEY,
    titre             TEXT         NOT NULL,
    titre_parallele   TEXT,
    annee_pub         SMALLINT,
    nb_pages          SMALLINT,
    id_editeur        INT,
    id_langue         INT,
    id_classification INT,
    date_catalogage   DATE DEFAULT (CURRENT_DATE),
    note              TEXT,
    FOREIGN KEY (id_editeur)        REFERENCES editeur(id_editeur),
    FOREIGN KEY (id_langue)         REFERENCES langue(id_langue),
    FOREIGN KEY (id_classification) REFERENCES classification(id_classification)
);

CREATE TABLE IF NOT EXISTS notice_auteur (
    id_notice   INT NOT NULL,
    id_auteur   INT NOT NULL,
    role        VARCHAR(80) DEFAULT 'auteur',
    PRIMARY KEY (id_notice, id_auteur, role),
    FOREIGN KEY (id_notice) REFERENCES notice(id_notice) ON DELETE CASCADE,
    FOREIGN KEY (id_auteur) REFERENCES auteur(id_auteur)
);

CREATE TABLE IF NOT EXISTS notice_matiere (
    id_notice   INT NOT NULL,
    id_matiere  INT NOT NULL,
    PRIMARY KEY (id_notice, id_matiere),
    FOREIGN KEY (id_notice) REFERENCES notice(id_notice) ON DELETE CASCADE,
    FOREIGN KEY (id_matiere) REFERENCES matiere(id_matiere)
);

CREATE TABLE IF NOT EXISTS exemplaire (
    id_exemplaire   INT AUTO_INCREMENT PRIMARY KEY,
    num_inventaire  VARCHAR(50) UNIQUE,
    id_notice       INT NOT NULL,
    cote_exemplaire VARCHAR(100),
    etat            VARCHAR(50) DEFAULT 'bon',
    date_acquisition DATE,
    note_ex         TEXT,
    FOREIGN KEY (id_notice) REFERENCES notice(id_notice)
);

INSERT IGNORE INTO langue (code_langue, libelle) VALUES
    ('fre', 'Français'),
    ('ara', 'Arabe'),
    ('eng', 'Anglais');

CREATE INDEX idx_notice_titre   ON notice (titre(100));
CREATE INDEX idx_auteur_nom     ON auteur (nom);
CREATE INDEX idx_exemplaire_inv ON exemplaire (num_inventaire);
