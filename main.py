from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io, os, subprocess, tempfile, base64
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
CORS(app)

CERT_COMPLET = base64.b64decode(open('/app/cert_complet.b64').read())
CERT_COMPACT = base64.b64decode(open('/app/cert_compact.b64').read())

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
    date_sig = datetime.today().strftime('%d.%m.%Y')

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
    _fill_table(table, date_sig, data['formateur'])

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def fill_compact(data):
    doc = Document(io.BytesIO(CERT_COMPACT))
    civ = 'Madame' if data['civilite'] == 'F' else 'Monsieur'
    nom_complet = f"{data['prenom']} {data['nom']}"
    date_cours = datetime.strptime(data['date_cours'], '%Y-%m-%d').strftime('%d.%m.%Y')
    date_sig = datetime.today().strftime('%d.%m.%Y')

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
    _fill_table(table, date_sig, data['formateur'])

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def _fill_table(table, date_sig, formateur):
    # Row 0 cell 0 = date signature
    cell_date = table.rows[0].cells[0]
    for para in cell_date.paragraphs:
        clear_para(para)
    p = cell_date.paragraphs[0]
    set_center(p)
    add_run(p, date_sig, 11, italic=True, color=(0xC0, 0x39, 0x2B))

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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
