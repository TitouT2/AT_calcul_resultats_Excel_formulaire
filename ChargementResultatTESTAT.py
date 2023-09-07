#Pour demarrer je charge le resultat du fichier EVE qui sera le user = 1 ensuite il faudra passer en parametre le nom du fichier et le nom de la personne
import sqlite3
import pandas as pd

#charger les donnees feuille  dans un dataframe Pandas
#chargement !er onglet = Egogramme

df_excel = pd.read_excel('/Users/evelyne/Documents/EVE.xlsx', sheet_name='Egogramme', usecols='C', skiprows=6, nrows=61)
#afficher le contenu du dataframe
print(df_excel)

conn = sqlite3.connect('AnalyseTransac.db')
cursor = conn.cursor()
#chargement de la table reponses (id_user,question,reponse)
table_name = 'reponses'
id_user = 'EVE'

#conversion du DataFrame excel en liste de tuple

reponse_data = [(reponse,) for reponse in df_excel['Note']]

#affichage de la liste de tuple
for tuple_data in reponse_data:
    for item in tuple_data:
        print(item)

#Ecriture dans la table
#insertion des lignes de tuples reponse_data qui correspondent au reponse du user
n = 0
query = "INSERT INTO reponses (id_user, question, reponse) VALUES (?, ?, ?)"

for tuple_data in reponse_data:
    for item in tuple_data:
        cursor.execute(f"SELECT Question FROM QuestionsAT LIMIT 1 OFFSET {n}")
        row = cursor.fetchone()
        print(row)
        cursor.execute(query, [id_user, n, item])
        n+=1
conn.commit()

#Affichage table reponse avec jointure sur le numero de reponse

conn = sqlite3.connect('AnalyseTransac.db')
cursor = conn.cursor()

#Affichage table reponses
cursor.execute("SELECT * FROM reponses")
results = cursor.fetchall()
for row in results:
    print("ligne:",row)
    Question = row[0]
    reponse = row[1]
    question_id = row[2]
    #reponse = row[2]
    print("id user:", id_user)
    print("Question :", Question)
    print("Reponse:", reponse)

#Affichage table QuestionAT
print("Affichage Question ID devrait demarrer a 0?")
cursor.execute("SELECT * FROM QuestionSAT")
results = cursor.fetchall()
for row in results:
    print("ligne:",row)
    Question = row[0]
    Driver = row[1]
    question_id = row[2]
    print("Question :", Question)
    print("Driver", Driver)
    print("Question_id:", question_id)

#Affichage jointure reponse et Question AT sur question_id
cursor.execute("SELECT * FROM reponses JOIN QuestionsAT ON reponses.question = QuestionsAT.question_id")
results = cursor.fetchall()

print("Affichage jointure reponse et QuestionsAT sur les questions")
for row in results:
    print("ligne de la jointure :",row)
    id_user = row[0]
    question = row[1]
    reponse = row[2]
    print("id user :", id_user)
    print("Question:", question)
    print("Reponse:", reponse)

#chargement des resultats dans la tablede result_egogramme (id_user,egogramme,valeur)

#Parent nourissier
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (3, 7, 13, 21,27,32,35,49,56,59) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Parent Nourissier"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

#Parent normatif
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (5, 11, 15, 23,25,31,36,47,51,55) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Parent Normatif"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

#Adulte
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (0, 10, 16, 18,26,28,37,41,43,53) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Adulte"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

#Enfant libre
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (4, 6, 14, 22,24,40,42,46,52,58) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Enfant libre"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

#Enfant Adapté soumis
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (2, 8, 12, 20,30,33, 39,45,50,57) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Enfant Adapte soumis"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

#Enfant Adapté rebelle
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (1, 9, 17, 19,29,34, 38,44,48,54) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
    #ID user en dur !!!!
id_user_EVE = 1
egogramme = "Enfant Adapte Rebelle"
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, result))

conn.commit()

#affichage resultat
cursor.execute("SELECT * FROM result_egogramme")
results = cursor.fetchall()
print("________________________________________________________________________________________")
for row in results:
    id_user = row[0]
    Posture = row[1]
    valeur = row[2]
    print("id_user :", id_user)
    print("Egogramme:", Posture)
    print("Valeur:", valeur)
print("________________________________________________________________________________________")
#------------------------------------------------------------------------------------------
#chargement des resultats de l'onglet Drivers d'un fichier Excel
#attention verifier
df_excel = pd.read_excel('.xlsx', sheet_name='Drivers', usecols='C', skiprows=7, nrows=50)
#afficher le contenu du dataframe
print(df_excel)

conn = sqlite3.connect('AnalyseTransac.db')
cursor = conn.cursor()
#chargement de la table reponses (id_user,question,reponse)
table_name = 'reponses'
id_user = 'EVE'

#conversion du DataFrame excel en liste de tuple

reponse_data = [(reponse,) for reponse in df_excel['Note']]

#affichage de la liste de tuple
for tuple_data in reponse_data:
    for item in tuple_data:
        print(item)

#Ecriture dans la table
#chargement en memoire des questions de la table QuestionAT
cursor.execute("SELECT Question FROM QuestionsAT")
questions = cursor.fetchall()

#insertion des lignes de tuples reponse_data qui correspondent aux reponse du user
n = 60
query = "INSERT INTO reponses (id_user, question, reponse) VALUES (?, ?, ?)"

for tuple_data in reponse_data:
    for item in tuple_data:
        cursor.execute(f"SELECT Question FROM QuestionsAT LIMIT 1 OFFSET {n}")
        row = cursor.fetchone()
        print(row)
        cursor.execute(query, [id_user, n, item])
        n+=1

conn.commit()
#Affichage table reponse avec jointure sur le numero de reponse

conn = sqlite3.connect('AnalyseTransac.db')
cursor = conn.cursor()

#Affichage table QuestionsAT
cursor.execute("SELECT * FROM QuestionsAT")
results = cursor.fetchall()
for row in results:
    print("ligne:",row)
    Question = row[0]
    Driver = row[1]
    question_id = row[2]
    #reponse = row[2]
    print("Question_id:", question_id)
    print("Question :", Question)
    print("Driver :", Question)

#Affichage table reponses
cursor.execute("SELECT * FROM reponses")
results = cursor.fetchall()
for row in results:
    print("ligne reponse, ordre colonnes :",row)
    id_user = row[0]
    question_id = row[1]
    reponse = row[2]
    print("id user:", id_user)
    print("Question :", question_id)
    print("Reponse:", reponse)

#Affichage jointure reponse et Question AT sur question_id
cursor.execute("SELECT * FROM reponses JOIN QuestionsAT ON reponses.question = QuestionsAT.question_id")
results = cursor.fetchall()

print("Affichage jointure des resultats EVE")
for row in results:
   print("ligne de la jointure :",row)
   id_user = row[0]
   question = row[1]
   reponse = row[2]
   print("id user :", id_user)
   print("Question:", question)
   print("Reponse:", reponse)

# chargement des resultats dans la table de result_egogramme (id_user,egogramme,valeur)

# Fais vite
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (60, 65, 70, 75, 80,81,85,90,95,100,105) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
print("somme pour Fais vite :")
print(query)

# affichage reponses 60,65,70,75,80,85,90,95,100,105,
cursor = conn.cursor()
query = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 60"
cursor.execute(query)
result = cursor.fetchone()[0]
print("reponse 60 :", result)

cursor = conn.cursor()
query2 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 65"
cursor.execute(query2)
result2 = cursor.fetchone()[0]
print("reponse 65 :", result2)

cursor = conn.cursor()
query3 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 70"
cursor.execute(query3)
result3 = cursor.fetchone()[0]
print("reponse 70 :", result3)

cursor = conn.cursor()
query4 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 75"
cursor.execute(query4)
result4 = cursor.fetchone()[0]
print("reponse 75 :", result4)

cursor = conn.cursor()
query5 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 80"
cursor.execute(query5)
result5 = cursor.fetchone()[0]
print("reponse 80 :", result5)

cursor = conn.cursor()
query6 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 85"
cursor.execute(query6)
result6 = cursor.fetchone()[0]
print("reponse 85 :", result6)

cursor = conn.cursor()
query7 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 90"
cursor.execute(query7)
result7 = cursor.fetchone()[0]
print("reponse 90 :", result7)

cursor = conn.cursor()
query8 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 95"
cursor.execute(query8)
result8 = cursor.fetchone()[0]
print("reponse 95 :", result8)

cursor = conn.cursor()
query9 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 100"
cursor.execute(query9)
result9 = cursor.fetchone()[0]
print("reponse 100 :", result9)

cursor = conn.cursor()
query10 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 105"
cursor.execute(query10)
result10 = cursor.fetchone()[0]
print("reponse 105 :", result10)

total = result + result2 + result3 + result4 + result5 + result6 + result7 + result8 + result9 + result10
print("total= ", total)
# ID user en dur !!!!
id_user_EVE = 1
egogramme = "Fais vite"
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, total))
print("EVE", egogramme, "=", total)

# Fais un effort
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (61, 66,71,76,81,86,91,96,101,106) AND id_user = 'EVE'"
cursor.execute(query)
result = cursor.fetchone()[0]
print("somme pour Fais un effort :")
print(query)
# affichage reponses 61,66,71,76,81,86,91,96,101,106
cursor = conn.cursor()
query = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 61"
cursor.execute(query)
result = cursor.fetchone()[0]
print("reponse 61 :", result)

cursor = conn.cursor()
query2 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 66"
cursor.execute(query2)
result2 = cursor.fetchone()[0]
print("reponse 66 :", result2)

cursor = conn.cursor()
query3 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 71"
cursor.execute(query3)
result3 = cursor.fetchone()[0]
print("reponse 71 :", result3)

cursor = conn.cursor()
query4 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 76"
cursor.execute(query4)
result4 = cursor.fetchone()[0]
print("reponse 76 :", result4)

cursor = conn.cursor()
query5 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 81"
cursor.execute(query5)
result5 = cursor.fetchone()[0]
print("reponse 81 :", result5)

cursor = conn.cursor()
query6 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 86"
cursor.execute(query6)
result6 = cursor.fetchone()[0]
print("reponse 86 :", result6)

cursor = conn.cursor()
query7 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 91"
cursor.execute(query7)
result7 = cursor.fetchone()[0]
print("reponse 91 :", result7)

cursor = conn.cursor()
query8 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 96"
cursor.execute(query8)
result8 = cursor.fetchone()[0]
print("reponse 96 :", result8)

cursor = conn.cursor()
query9 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 101"
cursor.execute(query9)
result9 = cursor.fetchone()[0]
print("reponse 101 :", result9)

cursor = conn.cursor()
query10 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 106"
cursor.execute(query10)
result10 = cursor.fetchone()[0]
print("reponse 106 :", result10)

total = result + result2 + result3 + result4 + result5 + result6 + result7 + result8 + result9 + result10
print("total= ", total)
# ID user en dur !!!!
id_user_EVE = 1
egogramme = "Fais un effort"
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, total))
print("EVE", egogramme, "=", total)

# Sois fort
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (62, 67, 72, 77,82,87,92,97,102,107) AND id_user = 'EVE'"
cursor = conn.cursor()
cursor.execute(query)
result = cursor.fetchone()[0]
print("somme pour Sois fort :")
print(query)
# affichage reponses 62,67,72,77,82,87,92,97,102,107
cursor = conn.cursor()
query = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 62"
cursor.execute(query)
result = cursor.fetchone()[0]
print("reponse 62 :", result)

cursor = conn.cursor()
query2 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 67"
cursor.execute(query2)
result2 = cursor.fetchone()[0]
print("reponse 67 :", result2)

cursor = conn.cursor()
query3 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 72"
cursor.execute(query3)
result3 = cursor.fetchone()[0]
print("reponse 72 :", result3)

cursor = conn.cursor()
query4 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 77"
cursor.execute(query4)
result4 = cursor.fetchone()[0]
print("reponse 77 :", result4)

cursor = conn.cursor()
query5 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 82"
cursor.execute(query5)
result5 = cursor.fetchone()[0]
print("reponse 82 :", result5)

cursor = conn.cursor()
query6 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 87"
cursor.execute(query6)
result6 = cursor.fetchone()[0]
print("reponse 87 :", result6)

cursor = conn.cursor()
query7 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 92"
cursor.execute(query7)
result7 = cursor.fetchone()[0]
print("reponse 92 :", result7)

cursor = conn.cursor()
query8 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 97"
cursor.execute(query8)
result8 = cursor.fetchone()[0]
print("reponse 97 :", result8)

cursor = conn.cursor()
query9 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 102"
cursor.execute(query9)
result9 = cursor.fetchone()[0]
print("reponse 102 :", result9)

cursor = conn.cursor()
query10 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 107"
cursor.execute(query10)
result10 = cursor.fetchone()[0]
print("reponse 107 :", result10)

total = result + result2 + result3 + result4 + result5 + result6 + result7 + result8 + result9 + result10
print("total= ", total)
# ID user en dur !!!!
id_user_EVE = 1
egogramme = "Sois fort"
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, total))
print("EVE", egogramme, "=", total)


# Sois parfait
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (63,68,73,78,83,88,93,98,103,108) AND id_user = 'EVE'"
cursor = conn.cursor()
cursor.execute(query)
result = cursor.fetchone()[0]
print("somme pour Sois parfait :")
print(query)
# affichage reponses 63,68,73,78,83,88,93,98,103,108
cursor = conn.cursor()
query = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 63"
cursor.execute(query)
result = cursor.fetchone()[0]
print("reponse 63 :", result)

cursor = conn.cursor()
query2 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 68"
cursor.execute(query2)
result2 = cursor.fetchone()[0]
print("reponse 68 :", result2)

cursor = conn.cursor()
query3 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 73"
cursor.execute(query3)
result3 = cursor.fetchone()[0]
print("reponse 73 :", result3)

cursor = conn.cursor()
query4 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 78"
cursor.execute(query4)
result4 = cursor.fetchone()[0]
print("reponse 78 :", result4)

cursor = conn.cursor()
query5 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 83"
cursor.execute(query5)
result5 = cursor.fetchone()[0]
print("reponse 83 :", result5)

cursor = conn.cursor()
query6 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 88"
cursor.execute(query6)
result6 = cursor.fetchone()[0]
print("reponse 88 :", result6)

cursor = conn.cursor()
query7 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 93"
cursor.execute(query7)
result7 = cursor.fetchone()[0]
print("reponse 93 :", result7)

cursor = conn.cursor()
query8 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 98"
cursor.execute(query8)
result8 = cursor.fetchone()[0]
print("reponse 98 :", result8)

cursor = conn.cursor()
query9 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 103"
cursor.execute(query9)
result9 = cursor.fetchone()[0]
print("reponse 103 :", result9)

cursor = conn.cursor()
query10 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 108"
cursor.execute(query10)
result10 = cursor.fetchone()[0]
print("reponse 108 :", result10)

total = result + result2 + result3 + result4 + result5 + result6 + result7 + result8 + result9 + result10
print("total= ", total)
# ID user en dur !!!!
id_user_EVE = 1
egogramme = "Sois Parfait"
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, total))
print("EVE", egogramme, "=", total)



# fais plaisir
cursor = conn.cursor()
query = "SELECT SUM(reponse) FROM reponses WHERE question IN (64,69,74,79,84,89,94,104,109) AND id_user = 'EVE'"
cursor = conn.cursor()
cursor.execute(query)
result = cursor.fetchone()[0]
print("somme pour Fais plaisir:")
print(query)
# affichage reponses 64,69,74,79,84,89,94,99,104,109
cursor = conn.cursor()
query = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 64"
cursor.execute(query)
result = cursor.fetchone()[0]
print("reponse 64 :", result)

cursor = conn.cursor()
query2 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 69"
cursor.execute(query2)
result2 = cursor.fetchone()[0]
print("reponse 69 :", result2)

cursor = conn.cursor()
query3 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 74"
cursor.execute(query3)
result3 = cursor.fetchone()[0]
print("reponse 74 :", result3)

cursor = conn.cursor()
query4 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 79"
cursor.execute(query4)
result4 = cursor.fetchone()[0]
print("reponse 79 :", result4)

cursor = conn.cursor()
query5 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 84"
cursor.execute(query5)
result5 = cursor.fetchone()[0]
print("reponse 84 :", result5)

cursor = conn.cursor()
query6 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 89"
cursor.execute(query6)
result6 = cursor.fetchone()[0]
print("reponse 89 :", result6)

cursor = conn.cursor()
query7 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 94"
cursor.execute(query7)
result7 = cursor.fetchone()[0]
print("reponse 94 :", result7)

cursor = conn.cursor()
query8 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 99"
cursor.execute(query8)
result8 = cursor.fetchone()[0]
print("reponse 99 :", result8)

cursor = conn.cursor()
query9 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 104"
cursor.execute(query9)
result9 = cursor.fetchone()[0]
print("reponse 104 :", result9)

cursor = conn.cursor()
query10 = "SELECT reponse FROM reponses WHERE id_user = 'EVE' and question = 109"
cursor.execute(query10)
result10 = cursor.fetchone()[0]
print("reponse 109 :", result10)

total = result + result2 + result3 + result4 + result5 + result6 + result7 + result8 + result9 + result10
print("total= ", total)
# ID user en dur !!!!
id_user_EVE = 1
egogramme = "Fais plaisir"
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, total))
print("EVE", egogramme, "=", total)
#Ecriture du graphique

#-------------------------------------------------------------------
#chargement des resultats de l'onglet Posture d'un fichier Excel
from openpyxl import load_workbook
fichier_excel = '/Users/evelyne/Documents/EVE.xlsx'
classeur = load_workbook(filename=fichier_excel)
onglet_postures = classeur["Postures"]

#Theme 1 : Lorsque je suis responsable
print("Lorsque je suis responsable")
valeur_cellule1 = onglet_postures['C11'].value
print("Je me justifie, je me défends ; parfois je critique, parfois je me protège est :",valeur_cellule1)
valeur_cellule2 = onglet_postures['C12'].value
print("J'utilise le contrôle et la persuasion, je n'hésite pas à faire pression est :",valeur_cellule2)
valeur_cellule3 = onglet_postures['C13'].value
print("J'aide ; ma sympathie m'aide à me faire accepter est :",valeur_cellule3)
valeur_cellule4 = onglet_postures['C14'].value
print("J'informe, je propose des occasions de développement, nous analysons ensemble les problèmes et les opportunités est :",valeur_cellule4)

#Thème 2 : Approche des Problèmes
print("Approche des Problèmes")
valeur_cellule5 = onglet_postures['C17'].value
print("J'essaie d'élucider, je m'arrange est :",valeur_cellule5)
valeur_cellule6 = onglet_postures['C18'].value
print("Je tiens aux objectifs et aussi à la qualité de la vie de chacun est :",valeur_cellule6)
valeur_cellule7 = onglet_postures['C19'].value
print("Je me soucie surtout de tenir les objectifs est :",valeur_cellule7)
valeur_cellule8 = onglet_postures['C20'].value
print("Je fais en sorte que chacun soit satisfait est :",valeur_cellule8)

#Thème 3 : Face aux règles
print("Face aux règles")
valeur_cellule9 = onglet_postures['C23'].value
print("Pour moi les règles sont les règles, c'est tout est :",valeur_cellule9)
valeur_cellule10 = onglet_postures['C24'].value
print("Les règles sont de bonnes choses. J'insiste pour qu'on les suive est :",valeur_cellule10)
valeur_cellule11 = onglet_postures['C25'].value
print("Ce sont des règles de conduite. Elles sont utiles mais n'en soyons pas prisonniers est :",valeur_cellule11)
valeur_cellule12 = onglet_postures['C26'].value
print("Je pense qu'on doit s'efforcer de les suivre est :",valeur_cellule12)

#Thème 4 : Vision des Conflits
print("Vision des Conflits")
valeur_cellule13 = onglet_postures['C29'].value
print("Pour moi les règles sont les règles, c'est tout est :",valeur_cellule13)
valeur_cellule14 = onglet_postures['C30'].value
print("Les règles sont de bonnes choses. J'insiste pour qu'on les suive est :",valeur_cellule14)
valeur_cellule15 = onglet_postures['C31'].value
print("Ce sont des règles de conduite. Elles sont utiles mais n'en soyons pas prisonniers est :",valeur_cellule15)
valeur_cellule16 = onglet_postures['C32'].value
print("Je pense qu'on doit s'efforcer de les suivre est :",valeur_cellule16)

#Thème 5 : Réaction à la Colère
print("Réaction à la Colère")
valeur_cellule17 = onglet_postures['C35'].value
print("Je n'aime pas avoir à gérer la colère, ca m'est pénible est :",valeur_cellule17)
valeur_cellule18 = onglet_postures['C36'].value
print("La colère me rend très désagréable  est :",valeur_cellule18)
valeur_cellule19 = onglet_postures['C37'].value
print("Je provoque une bonne confrontation est :",valeur_cellule19)
valeur_cellule20 = onglet_postures['C38'].value
print("J'en veux à ceux qui se permettent d'exprimer de la colère est :",valeur_cellule20)

#Thème 6 : Attitude envers un supérieur hiérarchique
print("Attitude envers un supérieur hiérarchique")
valeur_cellule21 = onglet_postures['C41'].value
print("Je vois bien ses points faibles, je le critique ou je le manipule est :",valeur_cellule21)
valeur_cellule22 = onglet_postures['C42'].value
print("Je fais de mon mieux et j'espère être apprécié est :",valeur_cellule22)
valeur_cellule23 = onglet_postures['C43'].value
print("Chacun son travail est :",valeur_cellule23)
valeur_cellule24 = onglet_postures['C44'].value
print("On discute, on échange, on négocie est :",valeur_cellule24)

#Thème 7 : Humour
print("Humour")
valeur_cellule25 = onglet_postures['C47'].value
print("Je fais rire à mes dépens est :",valeur_cellule25)
valeur_cellule26 = onglet_postures['C48'].value
print("Je pratique l'ironie désabusée est :",valeur_cellule26)
valeur_cellule27 = onglet_postures['C49'].value
print("Je sais trouver les mots qui libérent et détendent est :",valeur_cellule27)
valeur_cellule28 = onglet_postures['C50'].value
print("OMon humour est caustique et mordant est :",valeur_cellule28)

#Thème 8 : Face à l'Autre
print("Face à l'Autre")
valeur_cellule29 = onglet_postures['C53'].value
print("Je t'y ferai aller est :",valeur_cellule29)
valeur_cellule30 = onglet_postures['C54'].value
print("J'irai de l'avant avec toi est :",valeur_cellule30)
valeur_cellule31 = onglet_postures['C55'].value
print("Puisqu'il faut y aller est :",valeur_cellule31)
valeur_cellule32 = onglet_postures['C56'].value
print("Aller là ou ailleurs … est :",valeur_cellule32)

#calcul Décodage Positions de Vie
#Themes	Valeurs attribuées aux comportements
#1. Lorsque je suis responsable	4	5	2	1	3	2	1	2
#2. Approche des problèmes	    2	2	3	4	4	2	1	2
#3. Face aux règles         	3	4	2	1	4	3	1	2
#4. Vision des conflits	        1	5	3	2	2	3	4	0
#5. Face à la colère	        3	4	4	0	1	3	2	3
#6. Attitude envers un supéri	4	5	1	0	2	3	3	2
#7. Humour	                    3	4	4	4	1	2	2	0
#8. Face à l'autre             	2	6	1	4	3	0	4	0
#Total		                        ++      +-      -+      --
#CAlcul ++

totalplusplus = valeur_cellule4 + valeur_cellule6 + valeur_cellule11 + valeur_cellule13 + valeur_cellule19 + valeur_cellule24 + valeur_cellule27 + valeur_cellule30
print('total ++ =',totalplusplus)
print(valeur_cellule2,valeur_cellule7,valeur_cellule10,valeur_cellule15,valeur_cellule20,valeur_cellule21,valeur_cellule28,valeur_cellule29)
totalplusmoins = valeur_cellule2 + valeur_cellule7 + valeur_cellule10 + valeur_cellule15 + valeur_cellule20 + valeur_cellule21 + valeur_cellule28 + valeur_cellule29
print('total +- =',totalplusmoins)
totalmoinsplus = valeur_cellule3 + valeur_cellule8 + valeur_cellule12 + valeur_cellule14 + valeur_cellule17 + valeur_cellule22 + valeur_cellule25 + valeur_cellule31
print('total -+ =',totalmoinsplus)
totalmoinsmoins = valeur_cellule1 + valeur_cellule5 + valeur_cellule9 + valeur_cellule16 + valeur_cellule18 + valeur_cellule23 + valeur_cellule26 + valeur_cellule32
print('total -- =',totalmoinsmoins)
#insertions decodage Positions de vie dans egogramme

# ID user en dur !!!!
id_user_EVE = 1
egogramme = 'pluplus'
egogramme2 = 'plusmoins'
egogramme3 = 'moinsplus'
egogramme4 = 'moinsmoins'
cursor = conn.cursor()
query = "INSERT INTO result_egogramme (id_user,egogramme, valeur) VALUES (?, ?, ?)"
cursor.execute(query, (id_user_EVE, egogramme, totalplusplus))
cursor.execute(query, (id_user_EVE, egogramme2, totalplusmoins))
cursor.execute(query, (id_user_EVE, egogramme3, totalmoinsplus))
cursor.execute(query, (id_user_EVE, egogramme4, totalmoinsmoins))
#Attention verifier que les 4 insert ont bien ete faits
#tracage du graphique dans la feuille excel de resultat

#commit et fermeture base de donnees
conn.commit()

#afficahge de la table result_egogramme
cursor = conn.cursor()
cursor.execute("SELECT * FROM  result_egogramme")
results = cursor.fetchall()

print("Affichage table result_egogramme")
for row in results:
    id_user = row[0]
    egogramme = row[1]
    valeur = row[2]
    print("id user :", id_user)
    print("egogramme", egogramme)
    print("Reponse:", valeur)



# fin
conn.close()
