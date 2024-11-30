from flask import Flask, render_template, request, send_file, jsonify
import os
from werkzeug.utils import secure_filename
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from datetime import datetime
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-123')

# Créer les dossiers nécessaires s'ils n'existent pas
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}

# Définition des colonnes attendues dans l'ordre
EXPECTED_COLUMNS = [
    'Facture Numero',
    'Date de facture',
    'Client',
    'Date de Depart',
    'Date de Retour',
    'Marque du Vehicule',
    'Matricule',
    'Prix Total Ht',
    'TVA',
    'Prix TTC'
]

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def format_date(date_str):
    """Convertir la date du format Excel au format j-m-a"""
    try:
        # Convertir la chaîne de date en objet datetime
        if isinstance(date_str, str):
            date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        else:
            date_obj = date_str
        # Formater la date au format j-m-a
        return date_obj.strftime('%d-%m-%Y')
    except Exception as e:
        return date_str

def format_amount(amount):
    """Formater un montant avec deux décimales"""
    try:
        return f"{float(amount):,.2f}".replace(",", " ").replace(".", ",")
    except:
        return str(amount)

def create_invoice_pdf(data, output_path):
    """Créer une facture PDF avec les données fournies"""
    # Création du document PDF
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    # Ajout d'espace en haut pour le papier à en-tête (200 points = environ 7 cm)
    top_margin = 200

    # En-tête avec logo et informations de l'entreprise
    c.setFont("Helvetica-Bold", 22)
    c.drawString(50, height - (top_margin + 50), "FACTURE DE LOCATION")
    
    # Informations de l'entreprise (côté gauche)
    c.setFont("Helvetica", 10)
    c.drawString(50, height - (top_margin + 80), "7 RUE MOHAMED DIOURI ETG 3 N°149 CASABLANCA")
    c.drawString(50, height - (top_margin + 95), "Tel : +212 5 22 54 00 22")
    c.drawString(50, height - (top_margin + 110), "Capital : 7 000 000 DHS - RC : 309011 - I.F : 15186686")
    c.drawString(50, height - (top_margin + 125), "Taxe Professionnelle : 33066321 - C.N.S.S : 4052594")
    c.drawString(50, height - (top_margin + 140), "ICE : 0000 349 590 00014")
    
    # Cadre pour le numéro de facture (côté droit)
    frame_x = 400
    frame_y = height - (top_margin + 140)
    frame_width = 150
    frame_height = 90
    
    # Dessiner le cadre
    c.rect(frame_x, frame_y, frame_width, frame_height)
    
    # Centrer le texte dans le cadre
    c.setFont("Helvetica-Bold", 10)
    
    # Centrer "N° Facture"
    facture_text = f"N° Facture: {data['Facture Numero']}"
    facture_width = c.stringWidth(facture_text, "Helvetica-Bold", 10)
    facture_x = frame_x + (frame_width - facture_width) / 2
    facture_y = frame_y + frame_height - 30  # Position verticale ajustée
    c.drawString(facture_x, facture_y, facture_text)
    
    # Centrer la date
    date_text = f"Date: {format_date(data['Date de facture'])}"
    date_width = c.stringWidth(date_text, "Helvetica-Bold", 10)
    date_x = frame_x + (frame_width - date_width) / 2
    date_y = frame_y + frame_height - 50  # Position verticale ajustée
    c.drawString(date_x, date_y, date_text)
    
    # Informations client
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - (top_margin + 180), "FACTURER À:")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - (top_margin + 200), f"{data['Client']}")
    
    # Informations de location
    y = height - (top_margin + 250)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "INFORMATIONS DE LOCATION")
    
    # Tableau des informations
    y -= 30
    c.setFont("Helvetica", 10)
    # Première ligne
    c.drawString(50, y, "Date de départ:")
    c.drawString(150, y, f"{format_date(data['Date de Depart'])}")
    c.drawString(300, y, "Marque:")
    c.drawString(400, y, f"{data['Marque du Vehicule']}")
    
    # Deuxième ligne
    y -= 20
    c.drawString(50, y, "Date de retour:")
    c.drawString(150, y, f"{format_date(data['Date de Retour'])}")
    c.drawString(300, y, "Matricule:")
    c.drawString(400, y, f"{data['Matricule']}")
    
    # Tableau des prix
    y -= 50
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "DÉTAIL DES PRIX")
    
    # En-têtes du tableau
    y -= 30
    c.setStrokeColor(colors.black)
    c.setFillColor(colors.black)
    c.rect(50, y - 20, 500, 30, fill=0)
    
    # Lignes du tableau
    c.setFont("Helvetica", 10)
    # Prix HT
    c.drawString(60, y - 10, "Prix Total HT")
    c.drawRightString(520, y - 10, f"{format_amount(data['Prix Total Ht'])} MAD")
    
    # TVA
    y -= 30
    c.rect(50, y - 20, 500, 30, fill=0)
    c.drawString(60, y - 10, "TVA")
    c.drawRightString(520, y - 10, "20 %")
    
    # Prix TTC
    y -= 30
    c.rect(50, y - 20, 500, 30, fill=0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(60, y - 10, "Prix TTC")
    c.drawRightString(520, y - 10, f"{format_amount(data['Prix TTC'])} MAD")
    
    # Pied de page
    c.setFont("Helvetica", 8)
    footer_text1 = "PREPAID CAR RENTAL SARL AU - 7 RUE MOHAMED DIOURI ETG 3 N°149 CASABLANCA"
    footer_text2 = "Tel : +212 5 22 54 00 22 - Capital : 7 000 000 DHS - RC : 309011 - I.F : 15186686"
    footer_text3 = "Taxe Professionnelle : 33066321 - C.N.S.S : 4052594 - ICE : 0000 349 590 00014"
    
    c.drawCentredString(width/2, 40, footer_text1)
    c.drawCentredString(width/2, 30, footer_text2)
    c.drawCentredString(width/2, 20, footer_text3)
    
    c.save()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier n\'a été envoyé'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Aucun fichier sélectionné'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Lire le fichier Excel
            wb = openpyxl.load_workbook(filepath, data_only=True)  # data_only=True pour obtenir les valeurs calculées
            ws = wb.active
            
            # Vérifier les en-têtes
            headers = [cell.value for cell in ws[1]]
            if headers != EXPECTED_COLUMNS:
                return jsonify({
                    'error': 'Format de fichier incorrect. Les colonnes doivent être : ' + 
                            ', '.join(EXPECTED_COLUMNS)
                }), 400
            
            # Générer les PDFs
            generated_files = []
            
            # Pour chaque ligne de données
            for row in ws.iter_rows(min_row=2):
                # Créer un dictionnaire avec les données
                data = {headers[i]: cell.value for i, cell in enumerate(row)}
                
                # Vérifier que toutes les données nécessaires sont présentes
                if not all(data.values()):
                    continue
                
                # Générer le nom du fichier PDF
                output_filename = f"facture_{data['Facture Numero']}.pdf"
                output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
                
                # Créer le PDF
                create_invoice_pdf(data, output_path)
                generated_files.append(output_filename)
            
            return jsonify({
                'message': f'{len(generated_files)} facture(s) générée(s) avec succès',
                'files': generated_files
            })
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
        finally:
            # Nettoyer le fichier uploadé
            os.remove(filepath)
    
    return jsonify({'error': 'Type de fichier non autorisé'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(
            os.path.join(app.config['OUTPUT_FOLDER'], filename),
            as_attachment=True
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 404

if __name__ == '__main__':
    app.run(debug=True)
