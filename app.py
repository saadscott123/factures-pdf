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
    'Nombre de jours',
    'Prix par jour HT',
    'Prix location total HT',
    'Surclassement HT',
    'Sup 2eme Conducteur HT',
    'Out of Hours HT',
    'CDW HT',
    'TPC HT',
    'PAI HT',
    'SUPER CDW HT',
    'GPS HT',
    'Siege Bebe HT',
    'One Way HT',
    'Total Location HT',
    'TVA 20 %',
    'TOTAL TTC'
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
    if isinstance(amount, str):
        # Supprimer les espaces et remplacer la virgule par un point
        amount = amount.replace(' ', '').replace(',', '.')
        try:
            amount = float(amount)
        except ValueError:
            return "0,00"
    return f"{float(amount):,.2f}".replace(",", " ").replace(".", ",")

def create_invoice_pdf(data, output_path):
    """Créer une facture PDF avec les données fournies"""
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    # Augmenter considérablement l'espace en haut pour le papier à en-tête (8cm)
    top_margin = height - 8*cm

    # En-tête de la facture
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, top_margin, "FACTURE")
    
    # Numéro de facture et date
    c.setFont("Helvetica", 11)
    c.drawString(400, top_margin, f"N° : {data['Facture Numero']}")
    c.drawString(400, top_margin - 20, f"Date : {format_date(data['Date de facture'])}")

    # Informations du client
    c.drawString(30, top_margin - 60, "Client:")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(30, top_margin - 80, str(data['Client']))

    # Détails de la location
    c.setFont("Helvetica", 11)
    y = top_margin - 120
    c.drawString(30, y, f"Véhicule : {data['Marque du Vehicule']}")
    c.drawString(30, y - 20, f"Immatriculation : {data['Matricule']}")
    c.drawString(30, y - 40, f"Période de location : Du {format_date(data['Date de Depart'])} au {format_date(data['Date de Retour'])}")
    c.drawString(30, y - 60, f"Nombre de jours : {data['Nombre de jours']}")

    # Tableau des prestations
    y = y - 100
    c.setFont("Helvetica", 10)
    
    # Première ligne du tableau
    c.drawString(30, y, "Désignation")
    c.drawString(450, y, "Montant HT")
    y -= 20
    c.line(30, y + 15, 550, y + 15)  # Ligne de séparation

    # Prix de location
    c.drawString(30, y, "Prix location")
    c.drawString(450, y, f"{format_amount(data['Prix location total HT'])} MAD")
    
    # Autres prestations
    prestations = [
        ("Surclassement", "Surclassement HT"),
        ("2ème Conducteur", "Sup 2eme Conducteur HT"),
        ("Out of Hours", "Out of Hours HT"),
        ("CDW", "CDW HT"),
        ("TPC", "TPC HT"),
        ("PAI", "PAI HT"),
        ("SUPER CDW", "SUPER CDW HT"),
        ("GPS", "GPS HT"),
        ("Siège Bébé", "Siege Bebe HT"),
        ("One Way", "One Way HT")
    ]

    for label, key in prestations:
        if data[key] and float(data[key]) != 0:
            y -= 20
            c.drawString(30, y, label)
            c.drawString(450, y, f"{format_amount(data[key])} MAD")

    # Total HT
    y -= 30
    c.line(350, y + 25, 550, y + 25)  # Ligne de séparation
    c.drawString(350, y + 10, "Total HT")
    c.drawString(450, y + 10, f"{format_amount(data['Total Location HT'])} MAD")

    # TVA
    y -= 20
    c.drawString(350, y, "TVA 20%")
    c.drawString(450, y, f"{format_amount(data['TVA 20 %'])} MAD")

    # Total TTC
    y -= 20
    c.line(350, y + 15, 550, y + 15)  # Ligne de séparation
    c.drawString(350, y, "Total TTC")
    c.drawString(450, y, f"{format_amount(data['TOTAL TTC'])} MAD")

    # Pied de page centré avec plus d'espace en bas
    c.setFont("Helvetica", 8)
    footer_text = [
        "PREPAID CAR RENTAL S.A.R.L A.U, 7 RUE MOHAMED DIOURI ETG 3 N°149, CASABLANCA. Taxe Professionnelle : 33066321 - CNSS : 4052594 - ICE : 0000 349 590 00014.",
        "TEL : (+212) 5 22 54 00 22. Capital : 7 000 000 DHS - RC : 309011 - IF : 15186686."
    ]
    
    # Position de départ pour le pied de page (1.5cm du bas)
    footer_start_y = 1.5*cm
    line_spacing = 12  # Espacement entre les lignes
    
    for i, text in enumerate(footer_text):
        text_width = c.stringWidth(text, "Helvetica", 8)
        x = (width - text_width) / 2
        c.drawString(x, footer_start_y + (i * line_spacing), text)

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
