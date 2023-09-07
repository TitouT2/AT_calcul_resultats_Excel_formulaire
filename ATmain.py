import pandas as pd

#Saisie de donnes dans formulaire HTML
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Récupérer les réponses du formulaire
        reponse_1 = request.form['reponse_1']
        reponse_2 = request.form['reponse_2']
        reponse_3 = request.form['reponse_3']

        # Vérifier si les réponses sont uniquement 0 ou 1
        if not (reponse_1 in ['0', '1']):
            return 'La réponse à la question 1 doit être 0 ou 1.'
        if not (reponse_2 in ['0', '1']):
            return 'La réponse à la question 2 doit être 0 ou 1.'
        if not (reponse_3 in ['0', '1']):
            return 'La réponse à la question 3 doit être 0 ou 1.'

        # Traiter les réponses ici
        # ...

        return 'Réponses enregistrées avec succès !'

    return render_template('formulaireAT.html')

if __name__ == '__main__':
    app.run(debug=True)

