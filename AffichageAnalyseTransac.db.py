#Affichage de la structure de la base de donnees AnalyseTRansac.db
import sqlite3
import pandas as pd

conn = sqlite3.connect('AnalyseTransac.db')
cursor=conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()

#Parcours des tables
for table in tables:
    cursor.execute("PRAGMA table_info({})".format(table[0]))
    columns = cursor.fetchall()
    print("Table:", table[0])
    for column in columns:
        print("- Column:", column[1])
    print()

#affichage des donnees dans les tables de la base de donnees AnalyseTransac

cursor.execute("SELECT question_id, Question, Driver FROM QuestionsAT")

results = cursor.fetchall()

for row in results:

    question_id = row[0]
    question = row[1]
    Driver = row[2]
    print("Question_id:", question_id)
    print("Question:", question)
    print("Driver:", Driver)
    print()

#affichage des donnees dans les tables de la base de donnees reponses

#fermeture connexion base sqlite3
conn.close()