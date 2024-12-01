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
        return jsonify({'success': False, 'error': 'Aucun fichier trouvé'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'Aucun fichier sélectionné'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Type de fichier non autorisé'})
    
    try:
        # Sauvegarder le fichier Excel
        filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(excel_path)
        
        # Charger le workbook
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb.active
        
        # Vérifier les en-têtes
        headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]
        
        # Créer un dictionnaire de correspondance des colonnes
        column_mapping = {}
        for expected_col in EXPECTED_COLUMNS:
            for i, header in enumerate(headers):
                if header.lower() == expected_col.lower():
                    column_mapping[expected_col] = i
                    break
        
        # Vérifier si toutes les colonnes requises sont présentes
        if len(column_mapping) != len(EXPECTED_COLUMNS):
            missing_columns = set(EXPECTED_COLUMNS) - set(column_mapping.keys())
            return jsonify({
                'success': False, 
                'error': f'Colonnes manquantes dans le fichier Excel: {", ".join(missing_columns)}'
            })
        
        # Créer un PDF pour chaque ligne
        pdf_files = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2), 2):
            # Extraire les données de la ligne en utilisant le mapping
            row_data = {}
            skip_row = False
            
            for col_name, col_idx in column_mapping.items():
                value = row[col_idx].value
                if value is None or value == '':
                    skip_row = True
                    break
                row_data[col_name] = value
            
            if skip_row:
                continue
            
            try:
                # Générer un nom unique pour le PDF
                invoice_num = str(row_data['Facture Numero']).strip()
                pdf_filename = f"facture_{invoice_num}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
                pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
                
                # Créer le PDF
                create_invoice_pdf(row_data, pdf_path)
                pdf_files.append(pdf_filename)
                
            except Exception as e:
                print(f"Erreur lors de la génération de la facture à la ligne {row_idx}: {str(e)}")
                continue
        
        # Nettoyer le fichier Excel
        os.remove(excel_path)
        
        if not pdf_files:
            return jsonify({
                'success': False,
                'error': 'Aucune facture n\'a pu être générée. Vérifiez le format de vos données.'
            })
        
        return jsonify({
            'success': True,
            'files': pdf_files,
            'message': f'{len(pdf_files)} factures générées avec succès'
        })
        
    except Exception as e:
        if os.path.exists(excel_path):
            os.remove(excel_path)
        return jsonify({'success': False, 'error': str(e)})

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
