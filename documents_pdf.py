"""Génération PDF des devis et factures SWISS ViTa Form (module du svf-server).

Intégration dans main.py (2 lignes, après la création de `app` et le chargement de LOGO_SVF) :
    from documents_pdf import register_document_routes
    register_document_routes(app, LOGO_SVF, BREVO_API_KEY)
"""
import io
import base64
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import requests
from flask import request, send_file, jsonify
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (CondPageBreak, BaseDocTemplate, Frame, PageTemplate, Paragraph,
                                Spacer, Table, TableStyle, KeepTogether, Image)

# Police Unicode (Liberation Sans, installée avec LibreOffice) pour les caractères hors WinAnsi (ć, etc.)
FONT, FONT_BOLD = 'Helvetica', 'Helvetica-Bold'
try:
    _dir = '/usr/share/fonts/truetype/liberation/'
    pdfmetrics.registerFont(TTFont('LibSans', _dir + 'LiberationSans-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('LibSans-Bold', _dir + 'LiberationSans-Bold.ttf'))
    pdfmetrics.registerFontFamily('LibSans', normal='LibSans', bold='LibSans-Bold')
    FONT, FONT_BOLD = 'LibSans', 'LibSans-Bold'
except Exception:
    pass

SOCIETE = {
    'nom': 'SWISS ViTa Form',
    'lignes': ['Avenue Kiener 29', '1400 Yverdon-les-Bains', 'Suisse', 'info@swissvf.ch',
               'Téléphone : 078 892 02 63', "Numéro d'identification de la société :", 'CHE-108.957.698'],
}
# Coordonnées de paiement (identiques aux factures actuelles)
QR_IBAN = 'CH31 0076 7000 C561 3150 4'
QR_CREDITOR = {'name': 'Detta Vincent', 'street': 'Rue des Remparts', 'house_num': '17',
               'pcode': '1400', 'city': 'Yverdon-les-Bains', 'country': 'CH'}
PAIEMENT = ['Detta Vincent', 'SWISS ViTa Form Detta', 'Rue des Remparts 17', '1400 Yverdon-les-Bains',
            'IBAN : CH31 0076 7000 C561 3150 4', 'Banque : Banque Cantonale Vaudoise']

NAVY = colors.HexColor('#0E1A3A')
GREY = colors.HexColor('#8A8FA8')
HEAD_BG = colors.HexColor('#D6E4FB')
BOX_BG = colors.HexColor('#F3F8FE')
LINE = colors.HexColor('#DADDE6')
BADGES = {
    'paye': ('Payé', '#CDEBD8'), 'facture': ('Facturé', '#CDEBD8'),
    'expire': ('Expiré', '#F9CFCF'), 'annule': ('Annulé', '#F9CFCF'),
    'refuse': ('Refusé', '#F9CFCF'), 'accepte': ('Accepté', '#CDEBD8'),
}


def _d(v):
    return Decimal(str(v or 0))


def _num(v):
    t = f'{_d(v):f}'
    return t.rstrip('0').rstrip('.') if '.' in t else t


def _chf(v):
    q = _d(v).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    entier, dec = f'{abs(q):.2f}'.split('.')
    entier = f'{int(entier):,}'.replace(',', ' ')
    return f"{'-' if q < 0 else ''}{entier},{dec} CHF"


def _date_fr(v):
    mois = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.']
    if isinstance(v, str):
        v = datetime.strptime(v[:10], '%Y-%m-%d').date()
    return f'{v.day} {mois[v.month - 1]} {v.year}'


def calcul_totaux(doc, lignes):
    sous = sum((_d(l.get('quantite')) * _d(l.get('prix_unitaire')) for l in lignes), Decimal(0))
    remise = (sous * _d(doc.get('remise_pct')) / 100).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    taxes = _d(doc.get('taxes'))
    total = sous - remise + taxes
    paye = _d(doc.get('montant_paye'))
    return {'sous_total': sous, 'remise': remise, 'taxes': taxes, 'total': total,
            'paye': paye, 'reste': total - paye}


def statut_affiche(doc):
    """Devis envoyé dont la date d'expiration est dépassée => 'expire'."""
    st = doc.get('statut')
    if doc.get('type') == 'devis' and st == 'envoye' and doc.get('date_echeance'):
        ech = doc['date_echeance']
        ech = datetime.strptime(ech[:10], '%Y-%m-%d').date() if isinstance(ech, str) else ech
        if ech < date.today():
            return 'expire'
    return st


def _esc(t):
    return (t or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')


def build_document_pdf(doc, lignes, logo_bytes=None):
    est_facture = doc['type'] == 'facture'
    avec_qr = False
    numero = f"{int(doc['numero']):07d}"
    tot = calcul_totaux(doc, lignes)

    s = lambda name, **kw: ParagraphStyle(name, fontName=kw.pop('font', FONT),
                                          fontSize=kw.pop('size', 9), leading=kw.pop('lead', 12),
                                          textColor=kw.pop('color', NAVY), **kw)
    st_norm, st_small = s('n'), s('sm', size=8.5, lead=11.5)
    st_bold = s('b', font=FONT_BOLD)
    st_title = s('t', font=FONT_BOLD, size=11, lead=14)
    st_desc = s('d', size=8.5, lead=11.5, color=GREY)
    st_right = s('r', alignment=TA_RIGHT)
    st_hnum = s('hn', font=FONT_BOLD, size=18 if est_facture else 9, lead=22 if est_facture else 12,
                alignment=TA_RIGHT)

    buf = io.BytesIO()
    pdf = BaseDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                          topMargin=15 * mm, bottomMargin=15 * mm,
                          title=f"{'Facture' if est_facture else 'Devis'} n° {numero}", author='SWISS ViTa Form')
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='f', leftPadding=0,
                  rightPadding=0, topPadding=0, bottomPadding=0)
    pdf.addPageTemplates([PageTemplate(id='p', frames=[frame])])
    W = pdf.width
    story = []

    # --- En-tête : logo + société | numéro + dates ---
    logo = ''
    if logo_bytes:
        w, h = ImageReader(io.BytesIO(logo_bytes)).getSize()
        logo = Image(io.BytesIO(logo_bytes), width=24 * mm, height=24 * mm * h / w)
    societe = [Paragraph(f'<b>{SOCIETE["nom"]}</b>', st_small)] + [Paragraph(x, st_small) for x in SOCIETE['lignes']]
    droite = [Paragraph(f"{'Facture' if est_facture else 'Devis'} n° {numero}", st_hnum)]
    dates = [f"Date d'émission : {_date_fr(doc['date_emission'])}" if est_facture else f"Émis le : {_date_fr(doc['date_emission'])}"]
    if doc.get('date_echeance'):
        dates.append(f"Échéance : {_date_fr(doc['date_echeance'])}" if est_facture else f"Expire le : {_date_fr(doc['date_echeance'])}")
    droite += [Paragraph(x, s('dt', size=8.5, lead=12, alignment=TA_RIGHT)) for x in dates]
    badge = BADGES.get(statut_affiche(doc))
    if badge and not (est_facture and statut_affiche(doc) == 'facture'):
        droite.insert(0, Table([[Paragraph(badge[0], s('bd', size=8.5, alignment=1))]], colWidths=[22 * mm],
                               style=[('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(badge[1])),
                                      ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)],
                               hAlign='RIGHT'))
        droite.insert(1, Spacer(1, 4))
    head = Table([[logo, societe, droite]], colWidths=[30 * mm, W * 0.5 - 30 * mm, W * 0.5])
    head.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                              ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story += [head, Spacer(1, 8 * mm)]

    # --- Objet + client ---
    if est_facture:
        story += [Paragraph('<b>Facturer à :</b>' if doc.get('client_adresse') else '<b>Infos client :</b>',
                            s('lbl', size=8.5, color=GREY))]
    else:
        if doc.get('titre'):
            story += [Paragraph(f"<b>{_esc(doc['titre'])}</b>", st_norm), Spacer(1, 4 * mm)]
    story += [Paragraph(f"<b>{_esc(doc['client_nom'])}</b>", st_norm)]
    for k in ('client_email', 'client_adresse', 'client_tel'):
        if doc.get(k):
            story.append(Paragraph(_esc(doc[k]), st_norm))
    story.append(Spacer(1, 6 * mm))
    if est_facture and doc.get('titre'):
        story += [Paragraph(_esc(doc['titre']), s('ft', font=FONT_BOLD, size=13, lead=16)), Spacer(1, 2 * mm)]

    # --- Tableau des lignes ---
    cw = [W - 100 * mm, 20 * mm, 38 * mm, 42 * mm]
    rows = [[Paragraph('<b>Article ou service</b>', st_norm), Paragraph('<b>Quantité</b>', st_norm),
             Paragraph('<b>Prix</b>', st_norm), Paragraph('<b>Total</b>', st_right)]]
    for l in sorted(lignes, key=lambda x: x.get('position', 0)):
        cell = [Paragraph(_esc(l['titre']), st_norm)]
        desc = (l.get('description') or '')
        if l.get('offert'):
            desc = (desc + ' *OFFERT*').strip()
        if desc:
            cell.append(Paragraph(_esc(desc), st_desc))
        total_l = _d(l['quantite']) * _d(l['prix_unitaire'])
        q = _d(l['quantite'])
        rows.append([cell, Paragraph(_num(q), st_norm),
                     Paragraph(_chf(l['prix_unitaire']), st_norm), Paragraph(_chf(total_l), st_right)])
    t = Table(rows, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HEAD_BG), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, LINE), ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7), ('LEFTPADDING', (0, 0), (0, -1), 6),
        ('RIGHTPADDING', (-1, 0), (-1, -1), 6)]))
    story += [t, Spacer(1, 8 * mm)]

    # --- Totaux ---
    lignes_tot = [[Paragraph('Sous-total', st_norm), Paragraph(_chf(tot['sous_total']), st_right)]]
    if tot['remise'] > 0:
        pct = _num(doc.get('remise_pct'))
        lignes_tot.append([Paragraph(f'Réduction ({pct} %)', st_norm), Paragraph(_chf(tot['remise']), st_right)])
    if est_facture:
        lignes_tot.append([Paragraph('Taxes', st_norm), Paragraph(_chf(tot['taxes']), st_right)])
        lignes_tot.append([Paragraph('Total de la facture', st_norm), Paragraph(_chf(tot['total']), st_right)])
        lignes_tot.append([Paragraph('Montant payé', st_norm), Paragraph(_chf(tot['paye']), st_right)])
        final = ('Reste à payer', tot['reste'])
    else:
        final = ('Prix total :', tot['total'])
    tt = Table(lignes_tot, colWidths=[40 * mm, 42 * mm], hAlign='RIGHT')
    tt.setStyle(TableStyle([('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]))
    box = Table([[Paragraph(final[0], st_norm), Paragraph(f'<b>{_chf(final[1])}</b>', s('fx', size=13, lead=16, alignment=TA_RIGHT))]],
                colWidths=[40 * mm, 42 * mm], hAlign='RIGHT')
    box.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), BOX_BG), ('BOX', (0, 0), (-1, -1), 0.5, HEAD_BG),
                             ('TOPPADDING', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    story.append(KeepTogether([tt, Spacer(1, 3 * mm), box]))

    # --- Notes / mentions / paiement ---
    if doc.get('notes'):
        story += [Spacer(1, 8 * mm), Paragraph('<b>Notes</b>', st_small), Paragraph(_esc(doc['notes']), st_small)]
    story.append(Spacer(1, 5 * mm if est_facture else 8 * mm))
    if est_facture:
        gauche = [Paragraph('Merci pour votre confiance,', st_small), Spacer(1, 2 * mm),
                  Paragraph('SWISS ViTa Form.', st_small), Spacer(1, 4 * mm),
                  Paragraph(f'Le paiement doit être effectué aux coordonnées suivantes <b>avec la mention du numéro de facture</b> : {int(doc["numero"]):05d}', st_small)]
        droite = [Paragraph(x, st_small) for x in PAIEMENT]
        bloc = Table([[gauche, droite]], colWidths=[W * 0.52, W * 0.48])
        bloc.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                  ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))
        story += [bloc]
        avec_qr = doc.get('statut') != 'annule' and tot['reste'] > 0
        if avec_qr:  # place pour le bulletin de paiement (105 mm depuis le bas de la page)
            story += [CondPageBreak(91 * mm), Spacer(1, 1)]
    else:
        story += [Paragraph('<b>Mentions légales</b>', st_small),
                  Paragraph('En acceptant ce devis, vous acceptez nos conditions générales.', st_small)]

    pdf.build(story)
    data = buf.getvalue()
    if avec_qr:
        try:
            data = _ajouter_qr(data, numero, tot['reste'])
        except Exception as e:  # la facture reste utilisable sans bulletin
            print('QR-facture non générée :', e)
    return data


def _ajouter_qr(pdf_bytes, numero, montant):
    """Superpose la section de paiement QR (norme suisse) en bas de la dernière page."""
    from qrbill import QRBill
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPDF
    from pypdf import PdfReader, PdfWriter
    bill = QRBill(account=QR_IBAN, creditor=QR_CREDITOR, amount=f'{Decimal(str(montant)):.2f}',
                  currency='CHF', additional_information=f'Facture {int(numero)}', language='fr')
    svg = io.StringIO()
    bill.as_svg(svg, full_page=True)
    drawing = svg2rlg(io.BytesIO(svg.getvalue().encode('utf-8')))
    qr_pdf = PdfReader(io.BytesIO(renderPDF.drawToString(drawing)))
    base = PdfReader(io.BytesIO(pdf_bytes))
    w = PdfWriter()
    for i, page in enumerate(base.pages):
        if i == len(base.pages) - 1:
            page.merge_page(qr_pdf.pages[0], over=False)
        w.add_page(page)
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()


def nom_fichier(doc):
    return f"{'Facture' if doc['type'] == 'facture' else 'Devis'}_{int(doc['numero']):07d}.pdf"


def register_document_routes(app, logo_bytes, brevo_api_key):
    @app.route('/generate-document-pdf', methods=['POST'])
    def generate_document_pdf():
        try:
            data = request.json or {}
            doc, lignes = data.get('document'), data.get('lignes', [])
            if not doc or doc.get('type') not in ('devis', 'facture') or not doc.get('numero'):
                return jsonify({'error': 'Document invalide'}), 400
            pdf = build_document_pdf(doc, lignes, logo_bytes)
            return send_file(io.BytesIO(pdf), mimetype='application/pdf', as_attachment=True,
                             download_name=nom_fichier(doc))
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @app.route('/send-document', methods=['POST'])
    def send_document():
        """Envoie le devis/la facture en pièce jointe au client (Brevo)."""
        try:
            data = request.json or {}
            doc, lignes = data.get('document'), data.get('lignes', [])
            dest = data.get('destinataire') or (doc or {}).get('client_email')
            if not doc or not dest:
                return jsonify({'error': 'Document ou destinataire manquant'}), 400
            pdf = build_document_pdf(doc, lignes, logo_bytes)
            tot = calcul_totaux(doc, lignes)
            est_facture = doc['type'] == 'facture'
            numero = f"{int(doc['numero']):07d}"
            libelle = 'Facture' if est_facture else 'Devis'
            html = f"""
            <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
              <div style="background: #c0392b; padding: 20px; text-align: center;">
                <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
                <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">{libelle} n° {numero}</p>
              </div>
              <div style="padding: 30px; background: #f9f9f9;">
                <p>Bonjour,</p>
                <p>Veuillez trouver ci-joint {'la facture' if est_facture else 'le devis'} n° {numero}
                {('« ' + _esc(doc.get('titre'))) + ' »' if doc.get('titre') else ''}
                d'un montant de <strong>{_chf(tot['total'])}</strong>.</p>
                <p>Vous pouvez aussi le retrouver à tout moment sur votre espace client :</p>
                <div style="text-align: center; margin: 24px 0;">
                  <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
                </div>
              </div>
              <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
                Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
              </div>
            </div>"""
            payload = {
                'sender': {'name': 'Swiss ViTa Form', 'email': 'info@swissvf.ch'},
                'to': [{'email': dest, 'name': doc.get('client_nom', '')}],
                'subject': f'{libelle} n° {numero} — SWISS ViTa Form',
                'htmlContent': html,
                'attachment': [{'content': base64.b64encode(pdf).decode(), 'name': nom_fichier(doc)}],
            }
            r = requests.post('https://api.brevo.com/v3/smtp/email',
                              headers={'api-key': brevo_api_key, 'Content-Type': 'application/json'}, json=payload)
            print(f'[BREVO document] status={r.status_code} body={r.text}')
            if r.status_code == 201:
                return jsonify({'status': 'sent'})
            return jsonify({'error': 'Echec envoi email', 'detail': r.text}), 500
        except Exception as e:
            return jsonify({'error': str(e)}), 500
