"""Rappel d'échéance des certificats BLS-AED-SRC (validité 2 ans).

- /send-rappel-bls : envoie le rappel validé dans le portail (au contact du client, ou au participant privé),
  copie cachée à info@swissvf.ch.
- compter_a_envoyer() : nombre de rappels en attente, pour le mail du lundi (mêmes règles que le portail).
"""
import calendar
import datetime as dt
import html
import os
import re
import unicodedata

import requests
from flask import jsonify, request

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
ADMIN = 'info@swissvf.ch'
EMAIL_RX = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


def _fr(iso):
    s = str(iso or '')[:10]
    return f'{s[8:10]}.{s[5:7]}.{s[0:4]}' if len(s) == 10 else s


def ajouter_mois(d, n):
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def _norm(s):
    s = unicodedata.normalize('NFD', str(s or '').lower())
    return re.sub(r'[^a-z]', '', ''.join(c for c in s if unicodedata.category(c) != 'Mn'))


def _sb(table, params):
    key = os.environ.get('SUPABASE_SERVICE_KEY', '')
    r = requests.get(f'{SB_URL}/rest/v1/{table}', params=params,
                     headers={'apikey': key, 'Authorization': f'Bearer {key}'}, timeout=30)
    r.raise_for_status()
    return r.json()


def compter_a_envoyer(aujourdhui):
    """Nombre de rappels BLS en attente de validation (mêmes règles que le portail)."""
    parts = _sb('participants', {'select': 'id,client_id,prenom,nom,cours,date_cours,confirme,absent,email',
                                 'cours': 'ilike.*BLS*'})
    clients = {str(c['id']): c for c in _sb('clients', {'select': 'id,nom,contact,rappel_bls_actif'})}
    deja = {(str(r['participant_id']), r['niveau']) for r in _sb('rappels_bls', {'select': 'participant_id,niveau'})}
    par_nom = {}
    for p in parts:
        par_nom.setdefault((_norm(p.get('prenom')), _norm(p.get('nom'))), []).append(p)
    groupes = set()
    for p in parts:
        if not p.get('confirme') or p.get('absent') or not p.get('date_cours'):
            continue
        c = clients.get(str(p.get('client_id')))
        if c and c.get('rappel_bls_actif') is False:
            continue
        d = dt.date.fromisoformat(str(p['date_cours'])[:10])
        if d >= aujourdhui:
            continue
        ech = ajouter_mois(d, 24)
        if ech < aujourdhui - dt.timedelta(days=30):
            continue
        if any(q['id'] != p['id'] and str(q.get('date_cours'))[:10] > d.isoformat()
               and (q.get('client_id') == p.get('client_id') or (p.get('email') and q.get('email') == p.get('email')))
               for q in par_nom.get((_norm(p.get('prenom')), _norm(p.get('nom'))), [])):
            continue
        niveau = 0
        if aujourdhui >= ajouter_mois(ech, -3):
            niveau = 0 if (str(p['id']), 2) in deja else 2
        elif aujourdhui >= ajouter_mois(ech, -6):
            niveau = 0 if (str(p['id']), 1) in deja else 1
        if niveau:
            prive = not c or 'cours public' in (c.get('nom') or '').lower() or not c.get('contact')
            groupes.add((p['id'] if prive else str(p.get('client_id')), niveau))
    return len(groupes)


def _mail(dest_nom, niveau, prive, civilite, participants, espace_client):
    delai = '6 mois' if niveau == 1 else '3 mois'
    lignes = ''.join(
        f'<tr><td style="padding:6px 8px;border-bottom:1px solid #eee">{html.escape(p.get("prenom", ""))} {html.escape(p.get("nom", ""))}</td>'
        f'<td style="padding:6px 8px;border-bottom:1px solid #eee">{_fr(p.get("date_cours"))}</td>'
        f'<td style="padding:6px 8px;border-bottom:1px solid #eee;font-weight:bold">{_fr(p.get("echeance"))}</td></tr>'
        for p in participants)
    tableau = ('<table style="border-collapse:collapse;font-size:14px;margin:10px 0 16px;width:100%">'
               '<tr style="background:#f5f5f5"><th style="text-align:left;padding:6px 8px">Participant·e</th>'
               '<th style="text-align:left;padding:6px 8px">Cours suivi le</th><th style="text-align:left;padding:6px 8px">Échéance</th></tr>'
               + lignes + '</table>')
    if prive:
        p = participants[0]
        salut = 'Madame' if civilite == 'F' else ('Monsieur' if civilite == 'M' else '')
        intro = (f"<p>Bonjour {salut + ' ' if salut else ''}{html.escape(p.get('nom', ''))},</p>"
                 f"<p>Votre certificat BLS-AED-SRC, obtenu le {_fr(p.get('date_cours'))}, arrive à échéance le "
                 f"<b>{_fr(p.get('echeance'))}</b> (validité de 2 ans).</p>")
        suite = ("<p>Pour conserver une certification valable, nous vous invitons à vous inscrire à l'un de nos prochains cours. "
                 "Répondez simplement à cet e-mail : nous vous proposerons les prochaines dates.</p>")
        bouton = ('<p style="text-align:center;margin:22px 0"><a href="https://www.swissvf.ch" style="background:#c0392b;color:#fff;'
                  'padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Voir nos cours</a></p>')
    else:
        intro = ("<p>Bonjour,</p>"
                 f"<p>Les certificats BLS-AED-SRC des personnes ci-dessous arrivent à échéance dans environ <b>{delai}</b> "
                 "(validité de 2 ans) :</p>" + tableau)
        suite = ("<p>Pour assurer la continuité de la certification de vos collaborateurs et collaboratrices, "
                 + ("nous vous proposons d'organiser dès maintenant leur prochain cours." if niveau == 1 else
                    "il est temps de planifier leur prochain cours : les dates se remplissent vite.")
                 + " Répondez simplement à cet e-mail"
                 + (" ou faites une demande de formation depuis votre espace client" if espace_client else '')
                 + " : nous vous proposerons des dates.</p>")
        bouton = ('<p style="text-align:center;margin:22px 0"><a href="https://portail.swissvf.ch" style="background:#c0392b;color:#fff;'
                  'padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Demander une formation</a></p>'
                  if espace_client else '')
    return f"""<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:#c0392b;padding:20px;text-align:center"><h1 style="color:#fff;margin:0;font-size:24px">SWISS ViTa Form</h1>
        <p style="color:rgba(255,255,255,.85);margin:5px 0 0">Échéance des certificats BLS-AED-SRC</p></div>
      <div style="padding:26px 30px;background:#f9f9f9;font-size:14px;line-height:1.6;color:#222">{intro}{suite}{bouton}
        <p>Meilleures salutations,<br><br>Vincent Detta<br>SWISS ViTa Form</p></div>
      <div style="background:#f0f0f0;padding:16px;text-align:center;font-size:12px;color:#888">Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63</div>
    </div>"""


def register_rappels_bls_routes(app, brevo_key):
    @app.route('/send-rappel-bls', methods=['POST'])
    def send_rappel_bls():
        try:
            b = request.json or {}
            dest = str(b.get('destinataire') or '').strip()
            parts = [p for p in (b.get('participants') or []) if isinstance(p, dict)][:200]
            niveau = 2 if b.get('niveau') == 2 else 1
            if not EMAIL_RX.match(dest) or not parts:
                return jsonify({'error': 'Destinataire ou participants manquants'}), 400
            prive = bool(b.get('prive'))
            corps = _mail(b.get('nom') or '', niveau, prive, b.get('civilite'), parts, bool(b.get('espace_client')))
            sujet = ('Votre certificat BLS-AED-SRC arrive à échéance' if prive else
                     f"Certificats BLS-AED-SRC : échéance dans {'6' if niveau == 1 else '3'} mois") + ' — SWISS ViTa Form'
            r = requests.post('https://api.brevo.com/v3/smtp/email',
                              headers={'api-key': brevo_key, 'Content-Type': 'application/json'},
                              json={'sender': {'name': 'Swiss ViTa Form', 'email': ADMIN},
                                    'replyTo': {'name': 'Vincent Detta — SWISS ViTa Form', 'email': ADMIN},
                                    'to': [{'email': dest, 'name': str(b.get('nom') or '')[:80]}],
                                    **({'bcc': [{'email': ADMIN, 'name': 'SWISS ViTa Form'}]} if dest.lower() != ADMIN else {}),
                                    'subject': sujet, 'htmlContent': corps}, timeout=20)
            print(f'[BREVO rappel BLS] status={r.status_code} body={r.text[:200]}')
            if r.status_code == 201:
                return jsonify({'status': 'sent'})
            return jsonify({'error': 'Echec envoi email', 'detail': r.text[:300]}), 500
        except Exception as e:
            return jsonify({'error': str(e)[:300]}), 500
