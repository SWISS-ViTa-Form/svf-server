"""Bons cadeaux — cours de premiers secours pour parents, futurs parents, grands-parents… à domicile.

/bon-cadeau-pdf   : renvoie le PDF du bon (A4 paysage, sans prix)
/send-bon-cadeau  : envoie le bon par email (message personnalisé), copie cachée à info@swissvf.ch
"""
import base64
import html
import io
import re
from datetime import datetime

import requests
from flask import jsonify, request, send_file
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas

from documents_pdf import FONT, FONT_BOLD, FONT_TITLE

ROUGE = HexColor('#c0392b')
TEXTE = HexColor('#222222')
GRIS = HexColor('#777777')
CREME = HexColor('#ffffff')  # fond blanc : le logo actuel a un fond blanc
ADMIN = 'info@swissvf.ch'
# Réservation des cours parents / futurs parents
CONTACT_NOM = 'Pepita Widmer'
CONTACT_TEL = '079 375 23 14'
CONTACT_MAIL = 'info@swissvf.ch'
PUBLIC = 'pour parents, futurs parents, grands-parents et toute personne qui garde des enfants'
EMAIL_RX = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def _fr(iso):
    try:
        return datetime.strptime(str(iso)[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
    except Exception:
        return str(iso or '')


def _centre(c, txt, y, font, size, color, w):
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawCentredString(w / 2, y, txt)


def _espace(c, txt, y, font, size, color, w, sp=4):
    c.setFont(font, size)
    c.setFillColor(color)
    tw = c.stringWidth(txt, font, size) + sp * (len(txt) - 1)
    x = (w - tw) / 2
    for ch in txt:
        c.drawString(x, y, ch)
        x += c.stringWidth(ch, font, size) + sp


def duree_libelle(d):
    return {'2h30': '2 h 30', '4h': '4 heures'}.get(str(d or ''), str(d or ''))


def build_bon_pdf(bon, logo_bytes=None):
    w, h = landscape(A4)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    c.setTitle(f"Bon cadeau {bon.get('numero', '')}")
    # fond et cadres
    c.setFillColor(CREME)
    c.rect(0, 0, w, h, stroke=0, fill=1)
    c.setStrokeColor(ROUGE)
    c.setLineWidth(1.2)
    c.rect(22, 22, w - 44, h - 44)
    c.setLineWidth(0.4)
    c.rect(28, 28, w - 56, h - 56)

    y = h - 78
    if logo_bytes:
        try:
            lw, lh = 150, 72
            c.drawImage(ImageReader(io.BytesIO(logo_bytes)), (w - lw) / 2, y - lh, width=lw, height=lh,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
    y -= 112
    _espace(c, 'BON CADEAU', y, FONT_TITLE, 38, TEXTE, w, 9)
    y -= 30
    _centre(c, 'Cours de premiers secours', y, FONT, 14, ROUGE, w)
    y -= 17
    _centre(c, PUBLIC, y, FONT, 11, ROUGE, w)
    y -= 19
    _centre(c, f"{duree_libelle(bon.get('duree'))}  ·  à domicile  ·  donné par une infirmière", y, FONT, 10.5, GRIS, w)

    # Pour / Offert par
    y -= 56
    pour = str(bon.get('beneficiaire') or '').strip()
    offert = str(bon.get('offert_par') or '').strip()
    if pour:
        _centre(c, 'pour', y, FONT, 10.5, GRIS, w)
        y -= 28
        _centre(c, pour, y, FONT_TITLE, 24, ROUGE, w)
        y -= 14
        c.setStrokeColor(HexColor('#e2cfc7'))
        c.setLineWidth(0.7)
        c.line(w / 2 - 190, y, w / 2 + 190, y)
        y -= 24
    if offert:
        _centre(c, f'offert par {offert}', y, FONT, 11.5, TEXTE, w)
        y -= 22
    msg = str(bon.get('message') or '').strip()
    if msg:
        y -= 4
        for ligne in simpleSplit(f'« {msg} »', FONT, 11, 520)[:4]:
            _centre(c, ligne, y, FONT, 11, TEXTE, w)
            y -= 16

    # Bandeau réservation
    bh = 66
    by = 50
    c.setFillColor(HexColor('#ffffff'))
    c.setStrokeColor(HexColor('#e2cfc7'))
    c.setLineWidth(0.8)
    c.roundRect(70, by, w - 140, bh, 8, stroke=1, fill=1)
    _centre(c, 'Pour réserver votre cours, contactez', by + bh - 20, FONT, 9.5, GRIS, w)
    _centre(c, f'{CONTACT_NOM}  ·  {CONTACT_TEL}  ·  {CONTACT_MAIL}', by + bh - 39, FONT_BOLD, 12.5, TEXTE, w)
    _centre(c, f"en indiquant le n° du bon  ·  valable jusqu'au {_fr(bon.get('valable_jusqu'))}", by + 12, FONT, 9.5, ROUGE, w)

    c.setFont(FONT, 8)
    c.setFillColor(GRIS)
    c.drawString(40, 34, f"N° {bon.get('numero', '')}")
    c.drawRightString(w - 40, 34, f"Émis le {_fr(bon.get('date_emission'))}  ·  SWISS ViTa Form · Av. Kiener 29 · 1400 Yverdon-les-Bains · www.swissvf.ch")
    c.showPage()
    c.save()
    return buf.getvalue()


def nom_fichier(bon):
    return f"Bon_cadeau_{str(bon.get('numero') or 'SVF').replace(' ', '_')}.pdf"


def register_bons_cadeaux_routes(app, logo_bytes, brevo_key):
    @app.route('/bon-cadeau-pdf', methods=['POST'])
    def bon_cadeau_pdf():
        try:
            bon = (request.json or {}).get('bon') or {}
            pdf = build_bon_pdf(bon, logo_bytes)
            return send_file(io.BytesIO(pdf), mimetype='application/pdf', as_attachment=False, download_name=nom_fichier(bon))
        except Exception as e:
            return jsonify({'error': str(e)[:300]}), 500

    @app.route('/send-bon-cadeau', methods=['POST'])
    def send_bon_cadeau():
        try:
            b = request.json or {}
            bon = b.get('bon') or {}
            dest = str(b.get('destinataire') or '').strip()
            message = str(b.get('message') or '').strip()
            if not EMAIL_RX.match(dest) or not message:
                return jsonify({'error': 'Destinataire ou message manquant'}), 400
            pdf = build_bon_pdf(bon, logo_bytes)
            corps = f"""<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
              <div style="background:#c0392b;padding:20px;text-align:center"><h1 style="color:#fff;margin:0;font-size:24px">SWISS ViTa Form</h1>
                <p style="color:rgba(255,255,255,.85);margin:5px 0 0">Bon cadeau</p></div>
              <div style="padding:26px 30px;background:#f9f9f9;font-size:14px;line-height:1.6;color:#222">{html.escape(message).replace(chr(10), '<br>')}</div>
              <div style="background:#f0f0f0;padding:16px;text-align:center;font-size:12px;color:#888">Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63</div></div>"""
            payload = {'sender': {'name': 'Swiss ViTa Form', 'email': ADMIN},
                       'replyTo': {'name': 'Vincent Detta — SWISS ViTa Form', 'email': ADMIN},
                       'to': [{'email': dest, 'name': str(b.get('nom') or '')[:80]}],
                       'subject': str(b.get('sujet') or '').strip() or 'Votre bon cadeau — SWISS ViTa Form',
                       'htmlContent': corps,
                       'attachment': [{'content': base64.b64encode(pdf).decode(), 'name': nom_fichier(bon)}]}
            if b.get('copie', True) and dest.lower() != ADMIN:
                payload['bcc'] = [{'email': ADMIN, 'name': 'SWISS ViTa Form'}]
            r = requests.post('https://api.brevo.com/v3/smtp/email',
                              headers={'api-key': brevo_key, 'Content-Type': 'application/json'}, json=payload, timeout=20)
            print(f'[BREVO bon cadeau] status={r.status_code} body={r.text[:200]}')
            if r.status_code == 201:
                return jsonify({'status': 'sent'})
            return jsonify({'error': 'Echec envoi email', 'detail': r.text[:300]}), 500
        except Exception as e:
            return jsonify({'error': str(e)[:300]}), 500
