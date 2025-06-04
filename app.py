from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from twilio.rest import Client
from dotenv import load_dotenv

app = Flask(__name__)
load_dotenv()

# Configuration Twilio
TWILIO_ACCOUNT_SID = os.getenv('TWILIO_ACCOUNT_SID')
TWILIO_AUTH_TOKEN = os.getenv('TWILIO_AUTH_TOKEN')
TWILIO_PHONE_NUMBER = os.getenv('TWILIO_PHONE_NUMBER')

client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

def sauvegarder_excel(nom, prenom, presence):
    fichier_excel = "reponses.xlsx"
    
    # Créer un nouveau DataFrame avec la réponse
    nouvelle_reponse = pd.DataFrame({
        'Nom': [nom],
        'Prénom': [prenom],
        'Présence': [presence]
    })
    
    try:
        # Si le fichier existe, lire et ajouter la nouvelle réponse
        if os.path.exists(fichier_excel):
            df_existant = pd.read_excel(fichier_excel)
            df_final = pd.concat([df_existant, nouvelle_reponse], ignore_index=True)
        else:
            df_final = nouvelle_reponse
        
        # Sauvegarder dans Excel
        df_final.to_excel(fichier_excel, index=False)
        return True
    except Exception as e:
        print(f"Erreur lors de la sauvegarde: {e}")
        return False

def envoyer_sms(numero, nom, prenom):
    try:
        message = client.messages.create(
            body=f"Bonjour! Merci de confirmer votre présence en cliquant sur ce lien: http://votre-site.com/rsvp",
            from_=TWILIO_PHONE_NUMBER,
            to=numero
        )
        return True
    except Exception as e:
        print(f"Erreur lors de l'envoi du SMS: {e}")
        return False

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

    if sauvegarder_excel(nom, prenom, presence):
        return render_template('index.html', 
                             message="Votre réponse a été enregistrée. Merci!", 
                             message_type="success")
    else:
        return render_template('index.html', 
                             message="Une erreur est survenue lors de l'enregistrement.", 
                             message_type="error")

@app.route('/statistiques')
def statistiques():
    try:
        if os.path.exists("reponses.xlsx"):
            df = pd.read_excel("reponses.xlsx")
            stats = {
                'total_reponses': len(df),
                'presents': len(df[df['Présence'] == 'Oui']),
                'absents': len(df[df['Présence'] == 'Non'])
            }
            return render_template('statistiques.html', stats=stats)
        else:
            return render_template('statistiques.html', 
                                 message="Aucune réponse enregistrée pour le moment.",
                                 stats={'total_reponses': 0, 'presents': 0, 'absents': 0})
    except Exception as e:
        return render_template('statistiques.html', 
                             message=f"Erreur lors de la lecture des statistiques: {e}",
                             stats={'total_reponses': 0, 'presents': 0, 'absents': 0})

if __name__ == '__main__':
    app.run(debug=True) 