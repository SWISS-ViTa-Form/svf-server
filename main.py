from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io, os, subprocess, tempfile, base64, requests
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl

app = Flask(__name__)
CORS(app, origins=['https://portail.swissvf.ch', 'http://localhost:3000', '*'])

CERT_COMPLET = base64.b64decode(open('/app/cert_complet.b64').read())
CERT_COMPACT = base64.b64decode(open('/app/cert_compact.b64').read())
FICHE_PRESENCE = base64.b64decode(open('/app/fiche_presence.b64').read())
FICHE_SALAIRE = base64.b64decode(open('/app/fiche_salaire.b64').read())
ORS_API_KEY = os.environ.get('ORS_API_KEY', '')

def clear_para(para):
    for run in para.runs:
        run.text = ''

def set_center(para):
    pPr = para._p.get_or_add_pPr()
    # Enlever jc existant
    for existing in pPr.findall(qn('w:jc')):
        pPr.remove(existing)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)

def add_run(para, text, size_pt, bold=False, italic=False, color=None):
    run = para.add_run(text)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run

def fill_complet(data):
    doc = Document(io.BytesIO(CERT_COMPLET))
    civ = 'Madame' if data['civilite'] == 'F' else 'Monsieur'
    nom_complet = f"{data['prenom']} {data['nom']}"
    date_cours = datetime.strptime(data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')

    # Para 8 = civilité (vide dans complet)
    para8 = doc.paragraphs[8]
    clear_para(para8)
    set_center(para8)
    add_run(para8, civ, 14, color=(0x5B, 0x5B, 0x5B))

    # Para 9 = nom (RecipientName, vide dans complet)
    para9 = doc.paragraphs[9]
    clear_para(para9)
    set_center(para9)
    add_run(para9, nom_complet, 20, bold=True, color=(0xC0, 0x39, 0x2B))

    # Para 14 = "Le :" + date
    para14 = doc.paragraphs[14]
    for run in para14.runs:
        if run.text.strip() == '':
            run.text = f' {date_cours}'
            break
    else:
        para14.add_run(f' {date_cours}')

    # Table row 0
    table = doc.tables[0]
    _fill_table(table, data['formateur'], data.get('formateur2'))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def fill_compact(data):
    doc = Document(io.BytesIO(CERT_COMPACT))
    civ = 'Madame' if data['civilite'] == 'F' else 'Monsieur'
    nom_complet = f"{data['prenom']} {data['nom']}"
    date_cours = datetime.strptime(data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')

    # Para 5 = "Monsieur/Madame" → remplacer par civilité
    para5 = doc.paragraphs[5]
    clear_para(para5)
    set_center(para5)
    add_run(para5, civ, 14, color=(0x5B, 0x5B, 0x5B))

    # Para 7 = "Tartenpion marcel" → remplacer par nom
    para7 = doc.paragraphs[7]
    clear_para(para7)
    set_center(para7)
    add_run(para7, nom_complet, 20, bold=True, color=(0xC0, 0x39, 0x2B))

    # Para 12 = "Le : " → ajouter date
    para12 = doc.paragraphs[12]
    for run in para12.runs:
        if 'Le' in run.text:
            run.text = f'Le : {date_cours}'
            break

    # Table row 0
    table = doc.tables[0]
    _fill_table(table, data['formateur'], data.get('formateur2'))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def _fill_table(table, formateur, formateur2=None):
    # Row 0 cell 0 = 2e instructeur (si présent), sinon vide
    cell_left = table.rows[0].cells[0]
    for para in cell_left.paragraphs:
        clear_para(para)
    if formateur2:
        p = cell_left.paragraphs[0]
        set_center(p)
        r0 = add_run(p, formateur2, 13, italic=True, color=(0xC0, 0x39, 0x2B))
        r0.font.name = 'Brush Script MT'

    # Row 0 cell 2 = formateur
    cell_sign = table.rows[0].cells[2]
    for para in cell_sign.paragraphs:
        clear_para(para)
    p2 = cell_sign.paragraphs[0]
    set_center(p2)
    r = add_run(p2, formateur, 13, italic=True, color=(0xC0, 0x39, 0x2B))
    r.font.name = 'Brush Script MT'

    # Vider row 1
    for ci in [0, 2]:
        for para in table.rows[1].cells[ci].paragraphs:
            clear_para(para)

def fill_fiche_presence(data):
    wb = openpyxl.load_workbook(io.BytesIO(FICHE_PRESENCE))
    ws = wb.active

    societe = data.get('societe', '') or ''
    adresse = data.get('adresse', '') or ''
    date_cours = data.get('date_cours', '')

    ws['B3'] = societe
    ws['B4'] = adresse
    if date_cours:
        try:
            ws['G6'] = datetime.strptime(date_cours, '%Y-%m-%d').strftime('%d.%m.%Y')
        except ValueError:
            ws['G6'] = date_cours

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def convert_to_pdf(docx_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'cert.docx')
        with open(docx_path, 'wb') as f:
            f.write(docx_bytes)
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            raise Exception(f'LibreOffice error: {result.stderr}')
        pdf_path = docx_path.replace('.docx', '.pdf')
        with open(pdf_path, 'rb') as f:
            return f.read()

def convert_xlsx_to_pdf(xlsx_bytes):
    """Identique a convert_to_pdf mais pour un fichier .xlsx (fiche de salaire)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, 'fiche.xlsx')
        with open(xlsx_path, 'wb') as f:
            f.write(xlsx_bytes)
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, xlsx_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            raise Exception(f'LibreOffice error: {result.stderr}')
        pdf_path = xlsx_path.replace('.xlsx', '.pdf')
        with open(pdf_path, 'rb') as f:
            return f.read()


MOIS_LABELS = {
    1: 'Janvier', 2: 'Fevrier', 3: 'Mars', 4: 'Avril', 5: 'Mai', 6: 'Juin',
    7: 'Juillet', 8: 'Aout', 9: 'Septembre', 10: 'Octobre', 11: 'Novembre', 12: 'Decembre',
    13: 'Gratification'
}


def fill_fiche_salaire(data):
    wb = openpyxl.load_workbook(io.BytesIO(FICHE_SALAIRE))
    ws = wb['fiche']

    civ = 'Madame' if data.get('civilite') == 'F' else 'Monsieur'
    date_emission = datetime.now().strftime('%d.%m.%Y')

    ws['D1'] = f'Yverdon-les-Bains le {date_emission}'
    ws['D9'] = civ
    ws['D10'] = data.get('nom_complet', '')
    ws['D11'] = data.get('adresse', '')
    ws['D12'] = data.get('npa_localite', '')
    ws['B15'] = data.get('avs_no', '')
    ws['B16'] = data.get('date_naissance', '')
    ws['B19'] = MOIS_LABELS.get(int(data.get('mois', 1)), '')
    ws['B20'] = f"Annee {data.get('annee', '')}"

    ws['B25'] = data.get('taux_horaire', 0)
    ws['C25'] = data.get('heures_total', 0)
    ws['E27'] = data.get('forfaits_ponctuels_montant', 0)
    ws['C28'] = data.get('materiel_nombre', 0)

    ws['C40'] = data.get('repas_nombre', 0)
    ws['C41'] = data.get('km_total', 0)
    ws['E41'] = data.get('km_montant', 0)

    ws['B44'] = data.get('recap_cours', '')

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate-cert', methods=['POST'])
def generate_cert():
    try:
        data = request.json
        for field in ['prenom', 'nom', 'civilite', 'cours', 'date_cours', 'formateur']:
            if field not in data:
                return jsonify({'error': f'Champ manquant: {field}'}), 400

        if 'Complet' in data['cours']:
            docx_bytes = fill_complet(data)
        else:
            docx_bytes = fill_compact(data)

        pdf_bytes = convert_to_pdf(docx_bytes)
        filename = f"Certificat_{data['prenom']}_{data['nom']}.pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/generate-fiche-presence', methods=['POST'])
def generate_fiche_presence():
    try:
        data = request.json or {}
        xlsx_bytes = fill_fiche_presence(data)
        date_cours = data.get('date_cours', '')
        filename = f"Fiche_presence_{date_cours}.xlsx" if date_cours else "Fiche_presence.xlsx"
        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== EMAIL ====================

BREVO_API_KEY = os.environ.get('BREVO_API_KEY', '')

def send_email_formateur(formateur_email, formateur_nom, cours_data):
    """Envoyer un email de notification au formateur"""
    if not formateur_email:
        return False
    
    date_f = datetime.strptime(cours_data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')
    
    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Nouveau cours assigné</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Un nouveau cours vous a été assigné. Voici les informations :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Type de cours</td><td style="padding: 8px 0; font-weight: bold;">{cours_data['type_cours']}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Date</td><td style="padding: 8px 0; font-weight: bold;">{date_f}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Horaire</td><td style="padding: 8px 0;">{cours_data['heure_debut']} – {cours_data['heure_fin']}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Lieu</td><td style="padding: 8px 0;">{cours_data['lieu']}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Participants prévus</td><td style="padding: 8px 0;">{cours_data.get('nb_participants_prevus', '—')}</td></tr>
                </table>
            </div>
            {f'<div style="background: #e6f1fb; border-radius: 8px; padding: 16px; margin: 16px 0;"><strong>Notes :</strong><br><pre style="font-family: Arial; white-space: pre-wrap; margin: 8px 0 0 0;">{cours_data["notes"]}</pre></div>' if cours_data.get('notes') else ''}
            <p>Veuillez confirmer votre présence en vous connectant au portail :</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """
    
    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": f"Cours assigné : {cours_data['type_cours']} — {date_f}",
        "htmlContent": html_content
    }
    
    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO formateur] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text

@app.route('/send-notification', methods=['POST'])
def send_notification():
    try:
        data = request.json
        formateur_email = data.get('formateur_email', '')
        formateur_nom = data.get('formateur_nom', '')
        cours = data.get('cours', {})
        
        if not formateur_email:
            return jsonify({'error': 'Email formateur manquant'}), 400
        
        success, detail = send_email_formateur(formateur_email, formateur_nom, cours)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_modification_cours(formateur_email, formateur_nom, cours_data, changements):
    """Informe un formateur/une formatrice d'un changement de date/heure/lieu sur un cours déjà assigné.
    changements : liste de dicts {label, ancien, nouveau}"""
    if not formateur_email:
        return False, 'email manquant'

    date_f = datetime.strptime(cours_data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')

    lignes_changements = ''.join([
        f"""<tr>
            <td style="padding: 8px 0; color: #888; width: 140px;">{c['label']}</td>
            <td style="padding: 8px 0;"><span style="text-decoration: line-through; color: #a32d2d;">{c['ancien']}</span> → <span style="font-weight: bold; color: #3b6d11;">{c['nouveau']}</span></td>
        </tr>"""
        for c in changements
    ])

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Modification d'un cours assigné</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Un cours qui vous est assigné a été modifié :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Type de cours</td><td style="padding: 8px 0; font-weight: bold;">{cours_data['type_cours']}</td></tr>
                    {lignes_changements}
                </table>
            </div>
            <p>Aucune nouvelle confirmation de votre part n'est nécessaire — cette information est juste à titre indicatif.</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": f"Cours modifié : {cours_data['type_cours']} — {date_f}",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO modification cours] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-modification-cours', methods=['POST'])
def send_modification_cours():
    """Envoie un email groupé aux formateurs/formatrices déjà assignés à un cours modifié (date/heure/lieu)"""
    try:
        data = request.json or {}
        formateurs = data.get('formateurs', [])
        cours = data.get('cours', {})
        changements = data.get('changements', [])

        if not formateurs:
            return jsonify({'error': 'Liste de formateurs vide'}), 400
        if not cours.get('date_cours') or not cours.get('type_cours'):
            return jsonify({'error': 'Données du cours incomplètes'}), 400
        if not changements:
            return jsonify({'error': 'Aucun changement fourni'}), 400

        sent = 0
        errors = []
        for f in formateurs:
            email = f.get('email', '')
            nom = f.get('nom', '')
            if not email:
                continue
            success, detail = send_email_modification_cours(email, nom, cours, changements)
            if success:
                sent += 1
            else:
                errors.append({'email': email, 'detail': detail})

        return jsonify({'status': 'done', 'sent': sent, 'total': len(formateurs), 'errors': errors})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_client_welcome(client_email, client_nom, login, mot_de_passe):
    """Envoyer un email de bienvenue au nouveau client avec ses identifiants"""
    if not client_email:
        return False

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Bienvenue</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour,</p>
            <p>Bienvenue chez Swiss ViTa Form ! Un espace client a été créé pour <strong>{client_nom}</strong> sur notre portail.</p>
            <p>Voici vos identifiants de connexion :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Identifiant</td><td style="padding: 8px 0; font-weight: bold;">{login}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Mot de passe</td><td style="padding: 8px 0; font-weight: bold;">{mot_de_passe}</td></tr>
                </table>
            </div>
            <p>Vous pouvez vous connecter dès maintenant pour consulter vos cours à venir et télécharger les certificats de vos participants :</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": client_email, "name": client_nom}],
        "subject": "Bienvenue chez Swiss ViTa Form — vos identifiants",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO client welcome] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-welcome-client', methods=['POST'])
def send_welcome_client():
    try:
        data = request.json
        client_email = data.get('client_email', '')
        client_nom = data.get('client_nom', '')
        login = data.get('login', '')
        mot_de_passe = data.get('mot_de_passe', '')

        if not client_email:
            return jsonify({'error': 'Email client manquant'}), 400

        success, detail = send_email_client_welcome(client_email, client_nom, login, mot_de_passe)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


WEBHOOK_SECRET = os.environ.get('WEBHOOK_SECRET', '')


def send_email_formateur_welcome(formateur_email, formateur_nom, login, mot_de_passe):
    """Envoyer un email de bienvenue au nouveau formateur avec ses identifiants"""
    if not formateur_email:
        return False, 'Email manquant'

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Bienvenue</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Bienvenue chez Swiss ViTa Form ! Un accès au portail formateur a été créé pour vous.</p>
            <p>Voici vos identifiants de connexion :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Identifiant</td><td style="padding: 8px 0; font-weight: bold;">{login}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Mot de passe</td><td style="padding: 8px 0; font-weight: bold;">{mot_de_passe}</td></tr>
                </table>
            </div>
            <p>Vous pouvez vous connecter dès maintenant pour consulter vos cours assignés :</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": "Bienvenue chez Swiss ViTa Form — vos identifiants",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO formateur welcome] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/webhook-new-formateur', methods=['POST'])
def webhook_new_formateur():
    try:
        # Vérification du secret partagé (header configuré côté Supabase)
        if WEBHOOK_SECRET:
            incoming_secret = request.headers.get('X-Webhook-Secret', '')
            if incoming_secret != WEBHOOK_SECRET:
                return jsonify({'error': 'Non autorisé'}), 401

        data = request.json or {}
        # Supabase envoie {"type":"INSERT","table":"formateurs","record":{...}}
        record = data.get('record', data)

        formateur_email = record.get('email', '')
        prenom = record.get('prenom', '')
        nom = record.get('nom', '')
        formateur_nom = f'{prenom} {nom}'.strip()
        login = record.get('login', '')
        mot_de_passe = record.get('mot_de_passe', '')

        if not formateur_email:
            return jsonify({'error': 'Email formateur manquant dans le webhook'}), 400

        success, detail = send_email_formateur_welcome(formateur_email, formateur_nom, login, mot_de_passe)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_nouveau_cours_disponible(formateur_email, formateur_nom, cours_data):
    """Informe un formateur/une formatrice qu'un nouveau cours est ouvert aux candidatures"""
    if not formateur_email:
        return False, 'email manquant'

    date_f = datetime.strptime(cours_data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Nouveau cours disponible</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Un nouveau cours est ouvert aux inscriptions des formateurs et formatrices. Voici les informations :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Type de cours</td><td style="padding: 8px 0; font-weight: bold;">{cours_data['type_cours']}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Date</td><td style="padding: 8px 0; font-weight: bold;">{date_f}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Horaire</td><td style="padding: 8px 0;">{cours_data.get('heure_debut', '—')} – {cours_data.get('heure_fin', '—')}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Lieu</td><td style="padding: 8px 0;">{cours_data.get('lieu', '—')}</td></tr>
                </table>
            </div>
            <p>Si vous êtes disponible et intéressé(e), vous pouvez vous porter candidat(e) directement depuis le portail. Votre demande sera ensuite validée par l'administration.</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": f"Nouveau cours disponible : {cours_data['type_cours']} — {date_f}",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO nouveau cours disponible] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-nouveau-cours-formateurs', methods=['POST'])
def send_nouveau_cours_formateurs():
    """Envoie un email groupé à une liste de formateurs/formatrices pour un cours nouvellement visible"""
    try:
        data = request.json or {}
        formateurs = data.get('formateurs', [])
        cours = data.get('cours', {})

        if not formateurs:
            return jsonify({'error': 'Liste de formateurs vide'}), 400
        if not cours.get('date_cours') or not cours.get('type_cours'):
            return jsonify({'error': 'Données du cours incomplètes'}), 400

        sent = 0
        errors = []
        for f in formateurs:
            email = f.get('email', '')
            nom = f.get('nom', '')
            if not email:
                continue
            success, detail = send_email_nouveau_cours_disponible(email, nom, cours)
            if success:
                sent += 1
            else:
                errors.append({'email': email, 'detail': detail})

        return jsonify({'status': 'done', 'sent': sent, 'total': len(formateurs), 'errors': errors})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_candidature_refusee(formateur_email, formateur_nom, cours_data):
    """Informe un formateur/une formatrice que sa candidature n'a pas été retenue"""
    if not formateur_email:
        return False, 'email manquant'

    date_f = datetime.strptime(cours_data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Réponse à votre candidature</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Merci pour votre intérêt pour le cours ci-dessous. Malheureusement, votre candidature n'a pas pu être retenue cette fois-ci :</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 140px;">Type de cours</td><td style="padding: 8px 0; font-weight: bold;">{cours_data['type_cours']}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Date</td><td style="padding: 8px 0; font-weight: bold;">{date_f}</td></tr>
                </table>
            </div>
            <p>N'hésitez pas à consulter le portail régulièrement pour découvrir d'autres cours disponibles.</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": f"Candidature non retenue : {cours_data['type_cours']} — {date_f}",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO candidature refusee] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-candidature-refusee', methods=['POST'])
def send_candidature_refusee():
    try:
        data = request.json or {}
        formateur_email = data.get('formateur_email', '')
        formateur_nom = data.get('formateur_nom', '')
        cours = data.get('cours', {})

        if not formateur_email:
            return jsonify({'error': 'Email formateur manquant'}), 400

        success, detail = send_email_candidature_refusee(formateur_email, formateur_nom, cours)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/send-welcome-formateur', methods=['POST'])
def send_welcome_formateur_manual():
    """Renvoi manuel de l'email de bienvenue depuis le portail admin (pour les formateurs déjà existants)"""
    try:
        data = request.json or {}
        formateur_email = data.get('formateur_email', '')
        formateur_nom = data.get('formateur_nom', '')
        login = data.get('login', '')
        mot_de_passe = data.get('mot_de_passe', '')

        if not formateur_email:
            return jsonify({'error': 'Email formateur manquant'}), 400

        success, detail = send_email_formateur_welcome(formateur_email, formateur_nom, login, mot_de_passe)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_signalement(sig_type, description, formateur_nom):
    """Envoyer un email pour un signalement (matériel) ou une idée"""
    label = 'Idée' if sig_type == 'idee' else 'Problème signalé'
    emoji = '💡' if sig_type == 'idee' else '⚠️'

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">{emoji} {label}</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Formateur : <strong>{formateur_nom or '—'}</strong></p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <pre style="font-family: Arial; white-space: pre-wrap; margin: 0;">{description}</pre>
            </div>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Ouvrir le portail admin</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": "info@swissvf.ch", "name": "Vincent"}],
        "subject": f"{emoji} {label} — {formateur_nom or 'Formateur'}",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO signalement] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-signalement', methods=['POST'])
def send_signalement():
    try:
        data = request.json or {}
        sig_type = data.get('type', 'materiel')
        description = data.get('description', '')
        formateur_nom = data.get('formateur_nom', '')

        if not description:
            return jsonify({'error': 'Description manquante'}), 400

        success, detail = send_email_signalement(sig_type, description, formateur_nom)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_demande_formation(client_nom, type_cours, nb_participants, date_souhaitee, message):
    """Envoyer un email pour une demande de formation depuis le portail client"""
    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">📋 Nouvelle demande de formation</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <div style="background: white; border-radius: 8px; padding: 20px; margin: 20px 0; border-left: 4px solid #c0392b;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px 0; color: #888; width: 160px;">Client</td><td style="padding: 8px 0; font-weight: bold;">{client_nom or '—'}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Type de cours</td><td style="padding: 8px 0;">{type_cours or '—'}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Participants estimés</td><td style="padding: 8px 0;">{nb_participants or '—'}</td></tr>
                    <tr><td style="padding: 8px 0; color: #888;">Date souhaitée</td><td style="padding: 8px 0;">{date_souhaitee or '—'}</td></tr>
                </table>
            </div>
            {f'<div style="background: #e6f1fb; border-radius: 8px; padding: 16px; margin: 16px 0;"><strong>Message :</strong><br>{message}</div>' if message else ''}
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Ouvrir le portail admin</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": "info@swissvf.ch", "name": "Vincent"}],
        "subject": f"📋 Demande de formation — {client_nom or ''}",
        "htmlContent": html_content
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO demande formation] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-demande-formation', methods=['POST'])
def send_demande_formation():
    try:
        data = request.json or {}
        client_nom = data.get('client_nom', '')
        type_cours = data.get('type_cours', '')
        nb_participants = data.get('nb_participants', '')
        date_souhaitee = data.get('date_souhaitee', '')
        message = data.get('message', '')

        success, detail = send_email_demande_formation(client_nom, type_cours, nb_participants, date_souhaitee, message)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== KM / DISTANCE ====================

def bareme_km(km_aller_retour, params):
    if km_aller_retour <= 30:
        return params.get('km_bareme_0_30', 21)
    elif km_aller_retour <= 60:
        return params.get('km_bareme_30_60', 45)
    elif km_aller_retour <= 100:
        return params.get('km_bareme_60_100', 70)
    else:
        return params.get('km_bareme_plus_100', 90)


def geocode_ors(address):
    resp = requests.get(
        'https://api.openrouteservice.org/geocode/search',
        params={'api_key': ORS_API_KEY, 'text': address, 'size': 1, 'boundary.country': 'CH'},
        timeout=15
    )
    resp.raise_for_status()
    features = resp.json().get('features', [])
    if not features:
        raise Exception(f'Adresse introuvable: {address}')
    lon, lat = features[0]['geometry']['coordinates']
    return lon, lat


@app.route('/calculate-distance-km', methods=['POST'])
def calculate_distance_km():
    try:
        data = request.json or {}
        origin = data.get('origin_address')
        destination = data.get('destination_address')
        params = data.get('parametres_rh', {})

        if not origin or not destination:
            return jsonify({'error': 'Adresse manquante'}), 400
        if not ORS_API_KEY:
            return jsonify({'error': 'ORS_API_KEY non configuree sur le serveur'}), 500

        lon1, lat1 = geocode_ors(origin)
        lon2, lat2 = geocode_ors(destination)

        resp = requests.post(
            'https://api.openrouteservice.org/v2/directions/driving-car',
            headers={'Authorization': ORS_API_KEY, 'Content-Type': 'application/json'},
            json={'coordinates': [[lon1, lat1], [lon2, lat2]]},
            timeout=15
        )
        resp.raise_for_status()
        route = resp.json()
        distance_m = route['routes'][0]['summary']['distance']
        km_aller_simple = distance_m / 1000
        km_aller_retour = round(km_aller_simple * 2, 1)

        montant = bareme_km(km_aller_retour, params)

        return jsonify({
            'km_aller_simple': round(km_aller_simple, 1),
            'km_aller_retour': km_aller_retour,
            'montant': montant
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== FICHE DE SALAIRE ====================

@app.route('/generate-fiche-salaire', methods=['POST'])
def generate_fiche_salaire():
    try:
        data = request.json or {}
        for field in ['nom_complet', 'mois', 'annee', 'taux_horaire']:
            if field not in data:
                return jsonify({'error': f'Champ manquant: {field}'}), 400

        xlsx_bytes = fill_fiche_salaire(data)
        pdf_bytes = convert_xlsx_to_pdf(xlsx_bytes)

        mois_label = MOIS_LABELS.get(int(data['mois']), data['mois'])
        filename = f"Fiche_salaire_{mois_label}_{data['annee']}_{data['nom_complet'].replace(' ', '_')}.pdf"

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def send_email_fiche_salaire(formateur_email, formateur_nom, mois_label, annee, pdf_base64, filename):
    if not formateur_email:
        return False, 'email manquant'

    html_content = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Fiche de salaire</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9;">
            <p>Bonjour {formateur_nom},</p>
            <p>Veuillez trouver ci-joint votre fiche de salaire pour <strong>{mois_label} {annee}</strong>.</p>
            <div style="text-align: center; margin: 24px 0;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Acceder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """

    payload = {
        "sender": {"name": "Swiss ViTa Form", "email": "info@swissvf.ch"},
        "to": [{"email": formateur_email, "name": formateur_nom}],
        "subject": f"Fiche de salaire — {mois_label} {annee}",
        "htmlContent": html_content,
        "attachment": [{"content": pdf_base64, "name": filename}]
    }

    response = requests.post(
        'https://api.brevo.com/v3/smtp/email',
        headers={'api-key': BREVO_API_KEY, 'Content-Type': 'application/json'},
        json=payload
    )
    print(f'[BREVO fiche salaire] status={response.status_code} body={response.text}')
    return response.status_code == 201, response.text


@app.route('/send-fiche-salaire', methods=['POST'])
def send_fiche_salaire():
    try:
        data = request.json or {}
        formateur_email = data.get('formateur_email', '')
        formateur_nom = data.get('formateur_nom', '')
        mois_label = data.get('mois_label', '')
        annee = data.get('annee', '')
        pdf_base64 = data.get('pdf_base64', '')
        filename = data.get('filename', 'Fiche_salaire.pdf')

        if not formateur_email or not pdf_base64:
            return jsonify({'error': 'Email ou PDF manquant'}), 400

        success, detail = send_email_fiche_salaire(formateur_email, formateur_nom, mois_label, annee, pdf_base64, filename)
        if success:
            return jsonify({'status': 'sent'})
        else:
            return jsonify({'error': 'Echec envoi email', 'detail': detail}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
