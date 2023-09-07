#Initialisation de la structure de base de donnees AnalyseTransac.db a partir du fichier Excel INITTABLESTESTAT
import sqlite3
import pandas as pd

#charger les donnees Excel dans un dataframe Pandas

df_excel = pd.read_excel('/Users/evelyne/Documents/INITTABLESTESTAT.xlsx', sheet_name='initquestions', usecols='B', skiprows=0, nrows=142)
print(df_excel)
print(df_excel)

conn = sqlite3.connect('AnalyseTransac.db')
cursor = conn.cursor()
table_name = 'QuestionsAT'
df_excel['Question'].to_sql(table_name,conn,if_exists='replace',index=False)

#Ajout d'une colonne Driver a la table QuestionsAT et remplissage valeurs

column_name = 'Driver'
value1 ='E'
value2 = 'D'
value3 = 'PRESP'
value4 ='PPB'
value5 = 'PREG'
value6 = 'PCON'
value7 = 'PCOL'
value8 = 'PHIER'
value9 = 'PHUM'
value10 = 'PAUT'

data_type = 'CHAR'
query = "ALTER TABLE {table} ADD COLUMN {column} {type};".format(table=table_name, column=column_name, type=data_type)
conn.commit()
cursor.execute(query)
query=f"UPDATE {table_name} SET {column_name} = ? LIMIT 61;"
cursor.execute(query, (value1,))
query2=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 62 and 110;"
cursor.execute(query2, (value2,))
query3=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 111 and 114;"
cursor.execute(query3, (value3,))
query4=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 115 and 118;"
cursor.execute(query4, (value4,))
query5=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 119 and 122;"
cursor.execute(query5, (value5,))
query6=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 123 and 126;"
cursor.execute(query6, (value6,))
query7=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 127 and 130;"
cursor.execute(query7, (value7,))
query8=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 131 and 134;"
cursor.execute(query8, (value8,))
query9=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 135 and 138;"
cursor.execute(query9, (value9,))
query10=f"UPDATE {table_name} SET {column_name} = ? WHERE rowid BETWEEN 139 and 142;"
cursor.execute(query10, (value10,))


#Ajout colonne question_id a la table QuestionsAT
data_type2 = 'INT'
column_name2 = 'question_id'
query11 = "ALTER TABLE {table} ADD COLUMN {column} {type};".format(table=table_name, column=column_name2, type=data_type2)
cursor.execute(query11)

#numerotation de question id de 1 a 142
cursor = conn.cursor()
# Exécution de la requête SQL pour numéroter les lignes
cursor.execute("SELECT rowid FROM QuestionsAT")
rows = cursor.fetchall()

for index, row in enumerate(rows, start=0):
   rowid = row[0]
   query = f"UPDATE QuestionsAT SET question_id = {index} WHERE rowid = {rowid}"
   cursor.execute(query)
   print(rowid)

#Affichage table QuestionsAT

# Remplacez "nom_de_votre_table" par le nom de la table que vous souhaitez afficher
table_name = 'QuestionsAT'

# Créez un curseur pour exécuter les commandes SQL
cursor = conn.cursor()

# Exécutez une requête SELECT pour obtenir tout le contenu de la table
cursor.execute(f"SELECT * FROM {table_name}")

# Affichez le contenu de la table
for row in cursor.fetchall():
    print(row)

# Fermez le curseur et la connexion à la base de données
cursor.close()


#creation de la table users (id_user, nom, mail,nom_fichier)
# Créez un curseur pour exécuter les commandes SQL
cursor = conn.cursor()
query = '''
    CREATE TABLE users (id_user INT, nom TEXT, mail TEXT, nom_fichier TEXT);
'''
cursor.execute(query)

#creation de la table reponses (id_user,question,reponse)
query2 = '''
CREATE TABLE reponses (id_user INT, question INT, reponse INT);
'''
cursor.execute(query2)
conn.commit()

#creation de la table result_egogramme (id_user,egogramme,valeur)
query3 = '''
CREATE TABLE result_egogramme (id_user INT, egogramme TEXT, valeur INT);
'''
cursor.execute(query3)


conn.commit()
#fermeture du la base de donnees
conn.close()