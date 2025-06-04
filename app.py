from flask import Flask, render_template, request
import pandas as pd
import os
from openpyxl import load_workbook

app = Flask(__name__)

def sauvegarder_excel(nom, prenom, presence):
    fichier_excel = "reponses.xlsx"
    
    try:
        # Créer un nouveau DataFrame avec la réponse
        nouvelle_reponse = pd.DataFrame({
            'Nom': [nom],
            'Prénom': [prenom],
            'Présence': [presence]
        })
        
        # Si le fichier existe, lire et ajouter la nouvelle réponse
        if os.path.exists(fichier_excel):
            df_existant = pd.read_excel(fichier_excel)
            df_final = pd.concat([df_existant, nouvelle_reponse], ignore_index=True)
        else:
            df_final = nouvelle_reponse

        # Sauvegarder dans Excel
        df_final.to_excel(fichier_excel, index=False)

        # Calculer le total des "Oui" et le mettre dans la cellule D2
        total_presents = len(df_final[df_final['Présence'] == 'Oui'])
        
        # Utiliser openpyxl pour modifier la cellule D2
        wb = load_workbook(fichier_excel)
        ws = wb.active
        ws['D1'] = 'Total Participants'
        ws['D2'] = total_presents
        wb.save(fichier_excel)
        
        return True, total_presents
    except Exception as e:
        print(f"Erreur lors de la sauvegarde: {e}")
        return False, 0

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    nom = request.form.get('nom')
    prenom = request.form.get('prenom')
    presence = request.form.get('presence')

    if not all([nom, prenom, presence]):
        return render_template('index.html', 
                             message="Veuillez remplir tous les champs", 
                             message_type="error")

    succes, total = sauvegarder_excel(nom, prenom, presence)
    if succes:
        return render_template('index.html', 
                             message=f"Votre réponse a été enregistrée.",
                             message_type="success")
    else:
        return render_template('index.html', 
                             message="Une erreur est survenue lors de l'enregistrement.", 
                             message_type="error")

if __name__ == '__main__':
    app.run(debug=True) 