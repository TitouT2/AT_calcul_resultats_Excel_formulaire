import sqlite3

#reinitialisation de la base de donnees AnalyseTransac

conn = sqlite3.connect('AnalyseTransac.db')

cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
for table in tables:
    cursor.execute("DROP TABLE IF EXISTS " + table[0])

conn.commit()
conn.close()