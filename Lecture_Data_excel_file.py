
import openpyxl
import sqlite3

#creation de la base de donnees sqlite3 ANalyseTransac.db

# Fonction pour créer une connexion à la base de données
def create_connection():
    conn = sqlite3.connect('AnalyseTransac.db')
    return conn

# Fonction pour créer la table si elle n'existe pas
def create_table(conn):
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS Egogramme
                      (id INTEGER PRIMARY KEY AUTOINCREMENT,
                       c8 TEXT, c9 TEXT, c10 TEXT, ..., c67 TEXT)''')
    conn.commit()

# Fonction pour insérer les données dans la base de données
def insert_data(conn, data):
    cursor = conn.cursor()
    cursor.execute('''INSERT INTO Egogramme (c8, c9, c10, ..., c67)
                      VALUES (?, ?, ?, ..., ?)''', data)
    conn.commit()

# Chemin vers le fichier Excel
chemin_fichier = '

# Ouvrir le fichier Excel
wb = openpyxl.load_workbook(chemin_fichier)

# Sélectionner l'onglet "Egogramme"
feuille = wb['Egogramme']

# Lire les données de la colonne C8 à C67
donnees = []
for cellule in feuille['C8:C67']:
    donnees.append(cellule[0].value)

# Fermer le fichier Excel
wb.close()

# Créer une connexion à la base de données
conn = create_connection()

# Créer la table si elle n'existe pas
create_table(conn)

# Insérer les données dans la base de données
insert_data(conn, tuple(donnees))

# Fermer la connexion à la base de données
conn.close()

