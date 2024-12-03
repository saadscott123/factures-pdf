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

# Création des dossiers s'ils n'existent pas
for folder in ['uploads', 'output']:
    os.makedirs(folder, exist_ok=True)

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
    """Formater les montants avec deux décimales"""
    if amount is None:
        return "0,00"
    try:
        # Convertir en float si c'est une chaîne
        if isinstance(amount, str):
            amount = float(amount.replace(' ', '').replace(',', '.'))
        # Formater avec deux décimales et remplacer le point par une virgule
        return "{:,.2f}".format(amount).replace(',', ' ').replace('.', ',')
    except (ValueError, TypeError):
        return "0,00"

def number_to_letters(number):
    """Convertir un nombre en lettres"""
    units = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf"]
    tens = ["", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante-dix", "quatre-vingt", "quatre-vingt-dix"]
    teens = ["dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
    
    def convert_less_than_thousand(n):
        if n == 0:
            return ""
        
        result = []
        # Centaines
        if n >= 100:
            if n // 100 == 1:
                result.append("cent")
            else:
                result.extend([units[n // 100], "cent"])
            n = n % 100
        
        # Dizaines et unités
        if n >= 10:
            if n < 20:
                result.append(teens[n - 10])
                return " ".join(result)
            else:
                ten_digit = n // 10
                unit_digit = n % 10
                if ten_digit == 7 or ten_digit == 9:
                    result.append(tens[ten_digit - 1])
                    if unit_digit == 1:
                        result.append("et")
                    result.append(teens[unit_digit])
                else:
                    result.append(tens[ten_digit])
                    if unit_digit == 1 and ten_digit != 8:
                        result.append("et")
                    if unit_digit > 0:
                        result.append(units[unit_digit])
        elif n > 0:
            result.append(units[n])
        
        return " ".join(result)
    
    if number == 0:
        return "zéro"
    
    # Séparer les parties entière et décimale
    parts = str(number).replace(',', '.').split('.')
    integer_part = int(parts[0])
    decimal_part = int(parts[1]) if len(parts) > 1 else 0
    
    result = []
    
    # Traiter la partie entière
    if integer_part == 0:
        result.append("zéro")
    else:
        # Millions
        if integer_part >= 1000000:
            millions = integer_part // 1000000
            if millions == 1:
                result.append("un million")
            else:
                result.extend([convert_less_than_thousand(millions), "millions"])
            integer_part = integer_part % 1000000
        
        # Milliers
        if integer_part >= 1000:
            thousands = integer_part // 1000
            if thousands == 1:
                result.append("mille")
            else:
                result.extend([convert_less_than_thousand(thousands), "mille"])
            integer_part = integer_part % 1000
        
        # Reste
        if integer_part > 0:
            result.append(convert_less_than_thousand(integer_part))
    
    # Ajouter "dirhams"
    result.append("dirhams")
    
    # Traiter la partie décimale
    if decimal_part > 0:
        result.append("et")
        result.append(convert_less_than_thousand(decimal_part))
        result.append("centimes")
    
    return " ".join(result).capitalize()

def create_invoice_pdf(data, output_path):
    """Créer une facture PDF avec les données fournies"""
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    # Configuration de la page
    page_width = 595
    page_height = 842
    
    # Définir la marge supérieure (ajustée pour redescendre légèrement)
    top_margin = height - 6*cm  # Réduit de 7cm à 6cm
    
    # En-tête de la facture
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, top_margin, "FACTURE")
    
    # Numéro de facture et date (à gauche sous FACTURE)
    c.setFont("Helvetica-Bold", 11)  
    c.drawString(30, top_margin - 30, f"N° : {data['Facture Numero']}.")
    c.drawString(30, top_margin - 50, f"Date : {format_date(data['Date de facture'])}.")
    
    # Informations du client (à droite)
    c.setFont("Helvetica-Bold", 11)  
    c.drawString(400, top_margin - 30, "Client:")
    c.setFont("Helvetica", 11)  
    c.drawString(400, top_margin - 50, f"{str(data['Client'])}.")
    
    # Détails de la location
    c.setFont("Helvetica-Bold", 11)
    y = top_margin - 80
    c.drawString(30, y, f"Véhicule : {data['Marque du Vehicule']}.")
    c.drawString(30, y - 20, f"Immatriculation : {data['Matricule']}.")
    c.drawString(30, y - 40, f"Période de location : Du {format_date(data['Date de Depart'])} au {format_date(data['Date de Retour'])}.")
    
    # Calcul du prix par jour avec TVA
    prix_location_ht = float(str(data['Prix location total HT']).replace(' ', '').replace(',', '.'))
    nombre_jours = int(data['Nombre de jours'])
    prix_par_jour_ht = prix_location_ht / nombre_jours if nombre_jours > 0 else 0
    prix_par_jour_ttc = prix_par_jour_ht * 1.20  # TVA 20%
    
    # Affichage du nombre de jours et prix par jour
    c.setFont("Helvetica-Bold", 11)
    c.drawString(30, y - 60, "Nombre de jours :")
    c.setFont("Helvetica", 11)
    c.drawString(120, y - 60, f"{data['Nombre de jours']}.")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(150, y - 60, f"Prix par jour TTC : {format_amount(prix_par_jour_ttc)} MAD.")
    
    # Tableau des prestations
    y = y - 120  # Augmenté de -100 à -120 pour encore plus d'espace
    
    # Définir les dimensions du tableau
    table_left = 30
    table_right = 550
    col_montant = 450

    # En-tête du tableau avec fond gris
    header_height = 20
    c.setFillColorRGB(0.9, 0.9, 0.9)
    c.rect(table_left, y + 15, table_right - table_left, header_height, fill=1)
    c.setFillColorRGB(0, 0, 0)

    # En-tête du tableau
    c.setFont("Helvetica-Bold", 10)
    c.drawString(table_left + 5, y + 20, "Désignation.")
    c.drawString(col_montant, y + 20, "Montant HT.")

    # Ligne horizontale sous l'en-tête
    c.line(table_left, y + 15, table_right, y + 15)

    # Prix de location
    y -= 5
    c.setFont("Helvetica", 10)
    c.drawString(table_left + 5, y, "Prix location")
    c.drawString(col_montant, y, f"{format_amount(data['Prix location total HT'])} MAD")
    c.line(table_left, y - 5, table_right, y - 5)

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

    # Point de départ pour les bordures verticales
    start_y = y + header_height + 20

    # Afficher toutes les prestations
    for label, key in prestations:
        y -= 20
        c.drawString(table_left + 5, y, label)
        c.drawString(col_montant, y, f"{format_amount(data[key])} MAD")
        c.line(table_left, y - 5, table_right, y - 5)

    # Bordures verticales du tableau principal
    c.line(table_left, start_y, table_left, y - 5)  # Gauche
    c.line(col_montant - 20, start_y, col_montant - 20, y - 5)  # Avant montant
    c.line(table_right, start_y, table_right, y - 5)  # Droite

    # Section des totaux
    y -= 20  # Espace avant les totaux
    totals_start_y = y + 15  # Ajusté pour un meilleur alignement
    
    # Dimensions de la section totaux
    totals_width = table_right - (col_montant - 120)
    
    # Total HT
    c.setFont("Helvetica-Bold", 10)
    c.drawString(col_montant - 100, y, "Total HT")
    c.drawString(col_montant, y, f"{format_amount(data['Total Location HT'])} MAD")
    
    # TVA
    y -= 20
    c.drawString(col_montant - 100, y, "TVA 20%")
    c.drawString(col_montant, y, f"{format_amount(data['TVA 20 %'])} MAD")
    
    # Total TTC
    y -= 20
    totals_end_y = y - 5

    # Bordures de la section totaux
    # Rectangle principal pour Total HT et TVA
    c.rect(col_montant - 120, y + 15, totals_width, totals_start_y - y - 15)  # Fond blanc pour la zone des totaux
    
    # Ligne de séparation entre Total HT et TVA
    c.line(col_montant - 120, y + 35, table_right, y + 35)
    
    # Ligne verticale de séparation pour les montants
    c.line(col_montant - 20, totals_start_y, col_montant - 20, y + 15)
    
    # Total TTC avec fond gris
    c.setFillColorRGB(0.9, 0.9, 0.9)
    c.rect(col_montant - 120, totals_end_y, totals_width, 20, fill=1)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(col_montant - 100, y, "Total TTC")
    c.drawString(col_montant, y, f"{format_amount(data['TOTAL TTC'])} MAD")
    
    # Ligne finale du bas
    c.line(col_montant - 120, totals_end_y, table_right, totals_end_y)

    # Montant en lettres
    y -= 40  # Espace après le Total TTC
    montant_lettres = number_to_letters(float(str(data['TOTAL TTC']).replace(' ', '').replace(',', '.')))
    
    # Texte en gras
    c.setFont("Helvetica-Bold", 10)
    texte_complet = f"Arrêtée la présente facture à la somme de : {montant_lettres}."
    
    # Dessiner le texte
    c.drawString(30, y, texte_complet)
    
    # Ajouter le soulignement
    text_width = c.stringWidth(texte_complet, "Helvetica-Bold", 10)
    c.line(30, y - 2, 30 + text_width, y - 2)

    # Ajouter "Signature" en gras à droite après un espace
    y -= 40  # Espace d'une ligne
    c.setFont("Helvetica-Bold", 11)
    signature_text = "Signature"
    text_width = c.stringWidth(signature_text, "Helvetica-Bold", 11)
    c.drawString(520 - text_width, y, signature_text)  # 520 est proche du bord droit, ajusté pour la marge

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
