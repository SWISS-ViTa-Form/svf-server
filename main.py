from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import zipfile, io, os, subprocess, tempfile, base64
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

app = Flask(__name__)
CORS(app)

CERT_COMPLET = base64.b64decode(open('/app/cert_complet.b64').read())
CERT_COMPACT = base64.b64decode(open('/app/cert_compact.b64').read())

def fill_certificate(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    civ = 'Madame' if data['civilite'] == 'F' else 'Monsieur'
    nom_complet = f"{data['prenom']} {data['nom']}"
    date_cours = data['date_cours']
    date_signature = datetime.today().strftime('%d.%m.%Y')
    formateur = data['formateur']

    # Civilité (para 8)
    para8 = doc.paragraphs[8]
    para8.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = para8._p.get_or_add_pPr()
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    run8 = para8.add_run(civ)
    run8.font.size = Pt(14)
    run8.font.color.rgb = RGBColor(0x5B, 0x5B, 0x5B)

    # Nom (para 9)
    para9 = doc.paragraphs[9]
    pPr9 = para9._p.get_or_add_pPr()
    jc9 = OxmlElement('w:jc')
    jc9.set(qn('w:val'), 'center')
    pPr9.append(jc9)
    run9 = para9.add_run(nom_complet)
    run9.font.size = Pt(20)
    run9.bold = True
    run9.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    # Date cours (para 14)
    para14 = doc.paragraphs[14]
    if len(para14.runs) > 1:
        para14.runs[1].text = f' {date_cours}'
    else:
        para14.add_run(f' {date_cours}')

    # Table row 0 — date signature et formateur sur la ligne
    table = doc.tables[0]

    # Date signature
    cell_date = table.rows[0].cells[0]
    for para in cell_date.paragraphs:
        for run in para.runs:
            run.text = ''
    if not cell_date.paragraphs[0].runs:
        run_d = cell_date.paragraphs[0].add_run(date_signature)
    else:
        run_d = cell_date.paragraphs[0].runs[0]
        run_d.text = date_signature
    run_d.font.size = Pt(11)
    run_d.italic = True
    run_d.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    cell_date.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Formateur
    cell_sign = table.rows[0].cells[2]
    for para in cell_sign.paragraphs:
        for run in para.runs:
            run.text = ''
    if not cell_sign.paragraphs[0].runs:
        run_f = cell_sign.paragraphs[0].add_run(formateur)
    else:
        run_f = cell_sign.paragraphs[0].runs[0]
        run_f.text = formateur
    run_f.font.size = Pt(13)
    run_f.italic = True
    run_f.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    cell_sign.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Vider row 1
    for ci in [0, 2]:
        cell = table.rows[1].cells[ci]
        for para in cell.paragraphs:
            for run in para.runs:
                run.text = ''

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def convert_to_pdf(docx_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'cert.docx')
        pdf_path = os.path.join(tmpdir, 'cert.pdf')
        with open(docx_path, 'wb') as f:
            f.write(docx_bytes)
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            raise Exception(f'LibreOffice error: {result.stderr}')
        with open(pdf_path, 'rb') as f:
            return f.read()

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate-cert', methods=['POST'])
def generate_cert():
    try:
        data = request.json
        required = ['prenom', 'nom', 'civilite', 'cours', 'date_cours', 'formateur']
        for field in required:
            if field not in data:
                return jsonify({'error': f'Champ manquant: {field}'}), 400

        template = CERT_COMPLET if 'Complet' in data['cours'] else CERT_COMPACT
        docx_bytes = fill_certificate(template, data)
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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
