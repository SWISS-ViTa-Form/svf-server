"""Rappel d'anniversaire : la veille, un mail part sur info@swissvf.ch.

Appelé une fois par jour par un déclencheur externe (Supabase pg_cron ou cron-job.org)
avec l'en-tête  X-Cron-Secret: <CRON_SECRET>.
Variables Render : CRON_SECRET (obligatoire), SUPABASE_SERVICE_KEY (déjà là), BREVO_API_KEY (déjà là).
"""
import os
import hmac
import html
import datetime as dt
from zoneinfo import ZoneInfo

import requests
from flask import request, jsonify

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
TZ = ZoneInfo('Europe/Zurich')
DEST = 'info@swissvf.ch'
MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août',
        'septembre', 'octobre', 'novembre', 'décembre']

_last_sent = {'date': None}   # évite un double envoi si le déclencheur est appelé deux fois le même jour


def _est_bissextile(y):
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


def cibles_demain(aujourdhui):
    """Liste de (jour, mois) à fêter demain. Les nés un 29 février sont fêtés le 28 les années non bissextiles."""
    demain = aujourdhui + dt.timedelta(days=1)
    cibles = [(demain.day, demain.month)]
    if demain.month == 2 and demain.day == 28 and not _est_bissextile(demain.year):
        cibles.append((29, 2))
    return demain, cibles


def construire_mail(demain, personnes):
    noms = [f"{p['prenom']} {p['nom']}".strip() for p in personnes]
    prenoms = [p['prenom'] for p in personnes]
    if len(noms) == 1:
        sujet = f"🎂 Demain, c'est l'anniversaire de {noms[0]}"
        phrase = f"Demain, c'est l'anniversaire de <b>{html.escape(noms[0])}</b>."
    else:
        sujet = f"🎂 Demain, c'est l'anniversaire de {' et '.join(prenoms)}"
        phrase = "Demain, c'est l'anniversaire de " + ' et de '.join(f"<b>{html.escape(n)}</b>" for n in noms) + '.'
    date_txt = f"{demain.day} {MOIS[demain.month - 1]} {demain.year}"
    corps = (
        '<div style="font-family:Arial,sans-serif;font-size:15px;color:#222;max-width:520px">'
        f'<p>{phrase}</p>'
        f'<p style="color:#666;font-size:13px">({html.escape(date_txt)}) Un petit message, un coup de fil ? 🎉</p>'
        '<p style="color:#999;font-size:11px">Rappel automatique du portail Swiss ViTa Form.</p></div>'
    )
    return sujet, corps


def register_anniversaires_routes(app, brevo_key):
    cron_secret = os.environ.get('CRON_SECRET', '')

    @app.route('/cron/anniversaires', methods=['GET', 'POST'])
    def cron_anniversaires():
        recu = request.headers.get('X-Cron-Secret', '') or request.args.get('secret', '')
        if not cron_secret:
            return jsonify({'error': 'CRON_SECRET non configuré'}), 503
        if not hmac.compare_digest(recu.encode(), cron_secret.encode()):
            return jsonify({'error': 'non autorisé'}), 401

        # ?date=AAAA-MM-JJ permet de simuler « aujourd'hui » pour tester ; ?test=1 n'enregistre pas l'envoi du jour
        test = request.args.get('test') == '1'
        try:
            aujourdhui = dt.date.fromisoformat(request.args['date']) if request.args.get('date') else dt.datetime.now(TZ).date()
        except ValueError:
            return jsonify({'error': 'date invalide'}), 400
        if not test and _last_sent['date'] == aujourdhui:
            return jsonify({'status': 'deja_envoye', 'date': aujourdhui.isoformat()})

        key = os.environ.get('SUPABASE_SERVICE_KEY', '')
        if not key:
            return jsonify({'error': 'SUPABASE_SERVICE_KEY manquante'}), 503
        r = requests.get(
            f'{SB_URL}/rest/v1/formateurs',
            params={'select': 'prenom,nom,anniv_jour,anniv_mois', 'actif': 'eq.true', 'anniv_jour': 'not.is.null'},
            headers={'apikey': key, 'Authorization': f'Bearer {key}'}, timeout=20)
        if r.status_code != 200:
            return jsonify({'error': 'lecture formateurs impossible', 'detail': r.text[:200]}), 502

        demain, cibles = cibles_demain(aujourdhui)
        personnes = [p for p in r.json() if (p.get('anniv_jour'), p.get('anniv_mois')) in cibles]
        if not personnes:
            if not test:
                _last_sent['date'] = aujourdhui
            return jsonify({'status': 'aucun', 'demain': demain.isoformat()})

        sujet, corps = construire_mail(demain, personnes)
        m = requests.post(
            'https://api.brevo.com/v3/smtp/email',
            headers={'api-key': brevo_key, 'Content-Type': 'application/json'},
            json={'sender': {'name': 'Swiss ViTa Form', 'email': DEST},
                  'to': [{'email': DEST, 'name': 'Swiss ViTa Form'}],
                  'subject': sujet, 'htmlContent': corps}, timeout=20)
        print(f'[BREVO anniversaires] status={m.status_code} body={m.text[:200]}')
        if m.status_code != 201:
            return jsonify({'error': 'envoi mail échoué', 'detail': m.text[:200]}), 502
        if not test:
            _last_sent['date'] = aujourdhui
        return jsonify({'status': 'envoye', 'personnes': [f"{p['prenom']} {p['nom']}" for p in personnes]})
