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
from reportlab.platypus import (CondPageBreak, HRFlowable, Flowable, BaseDocTemplate, Frame, PageTemplate, Paragraph,
                                Spacer, Table, TableStyle, KeepTogether, Image)

# Polices du site / des certificats : Poppins (texte) + Lora (titres). Repli sur Liberation Sans puis Helvetica.
import os
FONT, FONT_BOLD, FONT_TITLE = 'Helvetica', 'Helvetica-Bold', 'Helvetica'
_ici = os.path.dirname(os.path.abspath(__file__))
for _d in (os.path.join(_ici, 'fonts'), '/app/fonts', '/usr/share/fonts/truetype/custom',
           '/usr/share/fonts/truetype/google-fonts'):
    try:
        pdfmetrics.registerFont(TTFont('PopL', os.path.join(_d, 'Poppins-Light.ttf')))
        pdfmetrics.registerFont(TTFont('PopR', os.path.join(_d, 'Poppins-Regular.ttf')))
        pdfmetrics.registerFont(TTFont('PopM', os.path.join(_d, 'Poppins-Medium.ttf')))
        pdfmetrics.registerFontFamily('PopL', normal='PopL', bold='PopM', italic='PopL', boldItalic='PopM')
        FONT, FONT_BOLD = 'PopL', 'PopM'
        break
    except Exception:
        continue
else:
    try:
        _dir = '/usr/share/fonts/truetype/liberation/'
        pdfmetrics.registerFont(TTFont('LibSans', _dir + 'LiberationSans-Regular.ttf'))
        pdfmetrics.registerFont(TTFont('LibSans-Bold', _dir + 'LiberationSans-Bold.ttf'))
        pdfmetrics.registerFontFamily('LibSans', normal='LibSans', bold='LibSans-Bold')
        FONT, FONT_BOLD = 'LibSans', 'LibSans-Bold'
    except Exception:
        pass
FONT_TITLE = FONT_BOLD
for _d in (os.path.join(_ici, 'fonts'), '/app/fonts', '/usr/share/fonts/truetype/custom',
           '/usr/share/fonts/truetype/google-fonts'):
    try:
        pdfmetrics.registerFont(TTFont('LoraT', os.path.join(_d, 'Lora-Variable.ttf')))
        FONT_TITLE = 'LoraT'
        break
    except Exception:
        continue

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

NAVY = colors.HexColor('#3A2C2C')      # texte (palette des certificats)
DARK = colors.HexColor('#2A2020')
GREY = colors.HexColor('#8A7A7A')
RED = colors.HexColor('#B8050F')
FRAME = colors.HexColor('#C0392B')
HEAD_BG = colors.HexColor('#F6ECE9')
BOX_BG = colors.HexColor('#FBF5F3')
LINE = colors.HexColor('#E3C9C2')
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
    remise_pct = (sous * _d(doc.get('remise_pct')) / 100).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    remise_chf = _d(doc.get('remise_montant')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    remise = min(remise_pct + remise_chf, sous)
    taxes = _d(doc.get('taxes'))
    total = sous - remise + taxes
    paye = _d(doc.get('montant_paye'))
    return {'sous_total': sous, 'remise': remise, 'remise_pct': remise_pct, 'remise_chf': remise_chf,
            'taxes': taxes, 'total': total, 'paye': paye, 'reste': total - paye}


def statut_affiche(doc):
    """Devis envoyé dont la date d'expiration est dépassée => 'expire'."""
    st = doc.get('statut')
    if doc.get('type') == 'devis' and st == 'envoye' and doc.get('date_echeance'):
        ech = doc['date_echeance']
        ech = datetime.strptime(ech[:10], '%Y-%m-%d').date() if isinstance(ech, str) else ech
        if ech < date.today():
            return 'expire'
    return st


def numero_affiche(n):
    """AAAA-NNNN (numero = AAAA*10000 + NNNN) ; anciens numéros sur 7 chiffres."""
    n = int(n)
    return f'{n // 10000}-{n % 10000:04d}' if n >= 10000 else f'{n:07d}'


def _esc(t):
    return (t or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')


class _TitreEspace(Flowable):
    """Titre en capitales espacées (comme « CERTIFICAT »), aligné à droite."""
    def __init__(self, texte, largeur, taille=20, espace=4.5, couleur=DARK):
        super().__init__()
        self.texte, self.largeur, self.taille, self.espace, self.couleur = texte, largeur, taille, espace, couleur
        self.height = taille * 1.35

    def wrap(self, aw, ah):
        return self.largeur, self.height

    def draw(self):
        c = self.canv
        c.setFont(FONT_TITLE, self.taille)
        c.setFillColor(self.couleur)
        w = pdfmetrics.stringWidth(self.texte, FONT_TITLE, self.taille) + self.espace * (len(self.texte) - 1)
        t = c.beginText(self.largeur - w, self.taille * 0.3)
        t.setFont(FONT_TITLE, self.taille)
        t.setCharSpace(self.espace)
        t.textOut(self.texte)
        c.drawText(t)


FOOTER = 'SWISS ViTa Form Vincent Detta  ·  Av. Kiener 29  ·  1400 Yverdon-les-Bains  ·  078 892 02 63  ·  info@swissvf.ch'
QR_H = 105 * mm


def _decorer(pdf_bytes, avec_qr, numero=''):
    """Cadre rouge fin + pied de page (style des certificats). Sur la page du bulletin QR,
    le cadre s'arrête à la ligne de découpe et le pied de page est omis."""
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas
    base = PdfReader(io.BytesIO(pdf_bytes))
    n = len(base.pages)
    pw, ph = A4
    m = 8.5 * mm
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    for i in range(n):
        c.setStrokeColor(FRAME)
        c.setLineWidth(0.75)
        vide = avec_qr and i == n - 1 and not (base.pages[i].extract_text() or '').strip()
        if vide:  # page réservée au bulletin QR : pas de cadre, juste un rappel
            c.setFont(FONT_TITLE, 14)
            c.setFillColor(RED)
            c.drawString(20 * mm, ph - 28 * mm, f'Facture n° {numero} — bulletin de paiement')
            c.setFont(FONT, 8.5)
            c.setFillColor(GREY)
            c.drawString(20 * mm, ph - 35 * mm, 'Merci de régler au moyen du bulletin ci-dessous (scan avec votre application bancaire).')
        elif avec_qr and i == n - 1:
            c.line(m, ph - m, pw - m, ph - m)
            c.line(m, ph - m, m, QR_H + 3 * mm)
            c.line(pw - m, ph - m, pw - m, QR_H + 3 * mm)
            c.line(m, QR_H + 3 * mm, pw - m, QR_H + 3 * mm)
        else:
            c.rect(m, m, pw - 2 * m, ph - 2 * m)
            c.setFont(FONT, 6.5)
            c.setFillColor(GREY)
            c.drawCentredString(pw / 2, m + 4.5 * mm, FOOTER)
        c.showPage()
    c.save()
    deco = PdfReader(io.BytesIO(buf.getvalue()))
    w = PdfWriter()
    for i, page in enumerate(base.pages):
        page.merge_page(deco.pages[i])
        w.add_page(page)
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()


def build_document_pdf(doc, lignes, logo_bytes=None):
    est_facture = doc['type'] == 'facture'
    avec_qr = False
    numero = numero_affiche(doc['numero'])
    tot = calcul_totaux(doc, lignes)

    s = lambda name, **kw: ParagraphStyle(name, fontName=kw.pop('font', FONT),
                                          fontSize=kw.pop('size', 9), leading=kw.pop('lead', 12.5),
                                          textColor=kw.pop('color', NAVY), **kw)
    st_norm, st_small = s('n'), s('sm', size=8.5, lead=12)
    st_desc = s('d', size=8, lead=11, color=GREY)
    st_right = s('r', alignment=TA_RIGHT)
    st_th = s('th', font=FONT_BOLD, size=7.5, lead=10, color=RED)
    st_th_r = s('thr', font=FONT_BOLD, size=7.5, lead=10, color=RED, alignment=TA_RIGHT)
    st_lbl = s('lbl', size=7.5, lead=10, color=RED, font=FONT_BOLD)

    titre_doc = 'Facture' if est_facture else 'Devis'
    buf = io.BytesIO()
    pdf = BaseDocTemplate(buf, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm,
                          topMargin=19 * mm, bottomMargin=19 * mm,
                          title=f"{titre_doc} n° {numero}", author='SWISS ViTa Form')
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='f', leftPadding=0,
                  rightPadding=0, topPadding=0, bottomPadding=0)
    pdf.addPageTemplates([PageTemplate(id='p', frames=[frame])])
    W = pdf.width
    story = []

    # --- En-tête : logo + société | titre espacé + numéro + dates ---
    logo = ''
    if logo_bytes:
        w, h = ImageReader(io.BytesIO(logo_bytes)).getSize()
        logo = Image(io.BytesIO(logo_bytes), width=30 * mm, height=30 * mm * h / w)
    societe = [Paragraph(f'<b>{SOCIETE["nom"]}</b>', s('sn', size=8.5, lead=12, color=DARK))] + \
              [Paragraph(x, s('sg', size=7.5, lead=10.5, color=GREY)) for x in SOCIETE['lignes']]
    droite_w = W * 0.5
    droite = [_TitreEspace(titre_doc.upper(), droite_w),
              Paragraph(f'N° {numero}', s('num', font=FONT_BOLD, size=11, lead=16, color=RED, alignment=TA_RIGHT))]
    dates = [f"Date d'émission : {_date_fr(doc['date_emission'])}" if est_facture else f"Émis le : {_date_fr(doc['date_emission'])}"]
    if doc.get('date_echeance'):
        dates.append(f"Échéance : {_date_fr(doc['date_echeance'])}" if est_facture else f"Expire le : {_date_fr(doc['date_echeance'])}")
    droite += [Spacer(1, 1.5 * mm)] + [Paragraph(x, s('dt', size=8, lead=11.5, color=GREY, alignment=TA_RIGHT)) for x in dates]
    badge = BADGES.get(statut_affiche(doc))
    if badge and not (est_facture and statut_affiche(doc) == 'facture'):
        droite += [Spacer(1, 2 * mm), Table([['', Paragraph(badge[0], s('bd', size=8, alignment=1))]],
                               colWidths=[droite_w - 22 * mm, 22 * mm],
                               style=[('BACKGROUND', (1, 0), (1, 0), colors.HexColor(badge[1])),
                                      ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                                      ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)])]
    head = Table([[[logo, Spacer(1, 2 * mm)] + societe, droite]], colWidths=[W * 0.5, droite_w])
    head.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                              ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story += [head, Spacer(1, 5 * mm),
              HRFlowable(width='100%', thickness=0.6, color=LINE, spaceAfter=5 * mm)]

    # --- Client + objet ---
    story += [Paragraph('FACTURER À' if (est_facture and doc.get('client_adresse')) else
                        'INFOS CLIENT' if est_facture else 'DESTINATAIRE', st_lbl), Spacer(1, 1.5 * mm),
              Paragraph(f"<b>{_esc(doc['client_nom'])}</b>", s('cn', size=10, lead=14, color=DARK))]
    for k in ('client_email', 'client_adresse', 'client_tel'):
        if doc.get(k):
            story.append(Paragraph(_esc(doc[k]), st_norm))
    story.append(Spacer(1, 7 * mm))
    if doc.get('titre'):
        story += [Paragraph(_esc(doc['titre']), s('ft', font=FONT_TITLE, size=14, lead=18, color=RED)), Spacer(1, 4 * mm)]

    # --- Tableau des lignes ---
    cw = [W - 100 * mm, 20 * mm, 38 * mm, 42 * mm]
    rows = [[Paragraph('ARTICLE OU SERVICE', st_th), Paragraph('QUANTITÉ', st_th),
             Paragraph('PRIX', st_th), Paragraph('TOTAL', st_th_r)]]
    for l in sorted(lignes, key=lambda x: x.get('position', 0)):
        cell = [Paragraph(f"<b>{_esc(l['titre'])}</b>", st_norm)]
        desc = (l.get('description') or '')
        if l.get('offert'):
            desc = (desc + ' *OFFERT*').strip()
        if desc:
            cell.append(Paragraph(_esc(desc), st_desc))
        total_l = _d(l['quantite']) * _d(l['prix_unitaire'])
        rows.append([cell, Paragraph(_num(_d(l['quantite'])), st_norm),
                     Paragraph(_chf(l['prix_unitaire']), st_norm), Paragraph(_chf(total_l), st_right)])
    t = Table(rows, colWidths=cw, repeatRows=1)
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEBELOW', (0, 0), (-1, 0), 0.8, FRAME),
        ('LINEBELOW', (0, 1), (-1, -1), 0.4, LINE), ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6), ('LEFTPADDING', (0, 0), (0, -1), 0),
        ('RIGHTPADDING', (-1, 0), (-1, -1), 0)]))
    story += [t, Spacer(1, 6 * mm)]

    # --- Totaux ---
    lignes_tot = [[Paragraph('Sous-total', st_norm), Paragraph(_chf(tot['sous_total']), st_right)]]
    if tot['remise_pct'] > 0:
        lignes_tot.append([Paragraph(f"Réduction ({_num(doc.get('remise_pct'))} %)", st_norm),
                           Paragraph('-' + _chf(tot['remise_pct']), st_right)])
    if tot['remise_chf'] > 0:
        lignes_tot.append([Paragraph('Réduction', st_norm), Paragraph('-' + _chf(tot['remise_chf']), st_right)])
    if est_facture:
        lignes_tot.append([Paragraph('Taxes', st_norm), Paragraph(_chf(tot['taxes']), st_right)])
        lignes_tot.append([Paragraph('Total de la facture', st_norm), Paragraph(_chf(tot['total']), st_right)])
        lignes_tot.append([Paragraph('Montant payé', st_norm), Paragraph(_chf(tot['paye']), st_right)])
        final = ('Reste à payer', tot['reste'])
    else:
        final = ('Prix total', tot['total'])
    tt = Table(lignes_tot, colWidths=[44 * mm, 40 * mm], hAlign='RIGHT')
    tt.setStyle(TableStyle([('TOPPADDING', (0, 0), (-1, -1), 2.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
                            ('RIGHTPADDING', (-1, 0), (-1, -1), 0), ('LEFTPADDING', (0, 0), (0, -1), 0)]))
    box = Table([[Paragraph(final[0].upper(), st_th), Paragraph(_chf(final[1]), s('fx', font=FONT_TITLE, size=14, lead=18, color=RED, alignment=TA_RIGHT))]],
                colWidths=[44 * mm, 40 * mm], hAlign='RIGHT')
    box.setStyle(TableStyle([('LINEABOVE', (0, 0), (-1, 0), 0.8, FRAME), ('LINEBELOW', (0, 0), (-1, 0), 0.4, LINE),
                             ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                             ('LEFTPADDING', (0, 0), (0, -1), 0), ('RIGHTPADDING', (-1, 0), (-1, -1), 0),
                             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    story.append(KeepTogether([tt, Spacer(1, 2 * mm), box]))

    # --- Notes / conditions / paiement ---
    notes = (doc.get('notes') or '').strip()
    if not est_facture and 'conditions générales' not in notes.lower():
        notes = (notes + '\n' if notes else '') + 'En acceptant ce devis, vous acceptez nos conditions générales.'
    if notes:
        story += [Spacer(1, 7 * mm), Paragraph('NOTES', st_lbl), Spacer(1, 1.5 * mm), Paragraph(_esc(notes), st_small)]
    story.append(Spacer(1, 6 * mm))
    if est_facture:
        gauche = [Paragraph('Merci pour votre confiance,', st_small), Spacer(1, 1.5 * mm),
                  Paragraph('SWISS ViTa Form.', st_small), Spacer(1, 3 * mm),
                  Paragraph(f'Le paiement doit être effectué aux coordonnées suivantes <b>avec la mention du numéro de facture</b> : {numero}', st_small)]
        droite_p = [Paragraph(x, st_small) for x in PAIEMENT]
        bloc = Table([[gauche, droite_p]], colWidths=[W * 0.52, W * 0.48])
        bloc.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                  ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))
        story += [bloc]
        avec_qr = doc.get('statut') != 'annule' and tot['reste'] > 0
        if avec_qr:  # place pour le bulletin de paiement (105 mm depuis le bas de la page)
            story += [CondPageBreak(QR_H - pdf.bottomMargin + 3 * mm), Spacer(1, 1)]

    pdf.build(story)
    data = buf.getvalue()
    try:
        data = _decorer(data, avec_qr, numero)
    except Exception as e:
        print('Décor (cadre/pied de page) non appliqué :', e)
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
                  currency='CHF', additional_information=f'Facture {numero}', language='fr')
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
    return f"{'Facture' if doc['type'] == 'facture' else 'Devis'}_{numero_affiche(doc['numero'])}.pdf"


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
            numero = numero_affiche(doc['numero'])
            libelle = 'Facture' if est_facture else 'Devis'
            note_cg = '' if est_facture else "<p>Nos conditions générales sont jointes à ce devis : en l'acceptant, vous les acceptez.</p>"
            # Lien portail uniquement si le client a un espace client (flag envoyé par le portail).
            # Un futur client sans login ne reçoit ni la phrase ni le bouton.
            bloc_portail = ''
            if data.get('avec_portail'):
                bloc_portail = """<p>Vous pouvez aussi le retrouver à tout moment sur votre espace client :</p>
                <div style="text-align: center; margin: 24px 0;">
                  <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
                </div>"""
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
                {note_cg}
                {bloc_portail}
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
            # Devis : on joint les conditions générales (conditions_generales.pdf à côté de ce fichier)
            cg_path = os.path.join(_ici, 'conditions_generales.pdf')
            if not est_facture and os.path.exists(cg_path):
                with open(cg_path, 'rb') as fh:
                    payload['attachment'].append({'content': base64.b64encode(fh.read()).decode(),
                                                  'name': 'Conditions_generales_SWISS_ViTa_Form.pdf'})
            r = requests.post('https://api.brevo.com/v3/smtp/email',
                              headers={'api-key': brevo_api_key, 'Content-Type': 'application/json'}, json=payload)
            print(f'[BREVO document] status={r.status_code} body={r.text}')
            if r.status_code == 201:
                return jsonify({'status': 'sent'})
            return jsonify({'error': 'Echec envoi email', 'detail': r.text}), 500
        except Exception as e:
            return jsonify({'error': str(e)}), 500
