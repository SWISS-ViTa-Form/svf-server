"""Rappel de facturation : chaque lundi, un mail part sur info@swissvf.ch avec
- les cours terminés qui n'ont pas encore de facture,
- les factures créées mais pas encore envoyées (brouillons).

Déclenché par le même appel quotidien que le rappel d'anniversaire (/cron/anniversaires),
ou directement par /cron/facturation (même en-tête X-Cron-Secret ; ?test=1 pour envoyer tout de suite).
Requiert la migration migration_facturation_cours.sql (colonnes cours.facture_id et cours.facturation).
"""
import datetime as dt
import hmac
import html
import os
from zoneinfo import ZoneInfo

import requests
from flask import jsonify, request

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
TZ = ZoneInfo('Europe/Zurich')
DEST = 'info@swissvf.ch'
_dernier_envoi = {'date': None}


def _sb(table, params):
    key = os.environ.get('SUPABASE_SERVICE_KEY', '')
    r = requests.get(f'{SB_URL}/rest/v1/{table}', params=params,
                     headers={'apikey': key, 'Authorization': f'Bearer {key}'}, timeout=20)
    r.raise_for_status()
    return r.json()


def _fr(iso):
    s = str(iso or '')[:10]
    return f'{s[8:10]}.{s[5:7]}.{s[0:4]}' if len(s) == 10 else s


def _num(n):
    n = int(n or 0)
    return f'{n // 10000}-{n % 10000:04d}' if n >= 10000 else str(n)


def contenu(aujourdhui):
    cours = _sb('cours', {
        'select': 'id,client_id,type_cours,date_cours,heure_debut,heure_fin,statut',
        'date_cours': f'lt.{aujourdhui.isoformat()}', 'facture_id': 'is.null', 'facturation': 'is.null',
        'or': '(statut.is.null,statut.neq.annulé)', 'order': 'date_cours.asc,heure_debut.asc'})
    brouillons = _sb('documents_commerciaux', {
        'select': 'numero,client_nom,titre', 'type': 'eq.facture', 'statut': 'eq.brouillon', 'order': 'numero.asc'})
    clients = {c['id']: c.get('nom') or '' for c in _sb('clients', {'select': 'id,nom'})} if cours else {}
    # une ligne par client et par jour
    groupes = {}
    for c in cours:
        k = (c.get('client_id'), str(c['date_cours'])[:10])
        groupes.setdefault(k, []).append(c)
    return groupes, brouillons, clients


def construire_mail(aujourdhui, groupes, brouillons, clients):
    n1, n2 = len(groupes), len(brouillons)
    sujet = '🧾 Facturation : ' + ' · '.join(x for x in (
        f"{n1} cours à facturer" if n1 else '', f"{n2} facture{'s' if n2 > 1 else ''} à envoyer" if n2 else '') if x)
    lignes1 = ''.join(
        f'<li><b>{html.escape(clients.get(cid) or "Sans client")}</b> — {_fr(date)} '
        f'<span style="color:#888">({html.escape(", ".join((c.get("type_cours") or "") + " " + (c.get("heure_debut") or "") for c in cs))})</span>'
        + (f' <span style="color:#c0392b">· il y a {(aujourdhui - dt.date.fromisoformat(date)).days} jours</span>'
           if (aujourdhui - dt.date.fromisoformat(date)).days > 14 else '') + '</li>'
        for (cid, date), cs in sorted(groupes.items(), key=lambda x: x[0][1]))
    lignes2 = ''.join(
        f'<li><b>F-{_num(d["numero"])}</b> — {html.escape(d.get("client_nom") or "")} '
        f'<span style="color:#888">{html.escape(d.get("titre") or "")}</span></li>' for d in brouillons)
    corps = ('<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px">'
             '<p>Bonjour Vincent,</p>'
             + (f'<p><b>🧾 Cours terminés à facturer ({n1})</b></p><ul>{lignes1}</ul>' if n1 else '')
             + (f'<p><b>✉️ Factures créées mais pas encore envoyées ({n2})</b></p><ul>{lignes2}</ul>' if n2 else '')
             + '<p style="text-align:center;margin:22px 0"><a href="https://portail.swissvf.ch" style="background:#c0392b;color:#fff;'
               'padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Ouvrir Devis &amp; factures</a></p>'
             '<p style="color:#999;font-size:11px">Rappel automatique du portail Swiss ViTa Form, chaque lundi.</p></div>')
    return sujet, corps


def envoyer(brevo_key, aujourdhui, forcer=False):
    """Envoie le rappel si on est lundi (ou si forcer) et s'il y a quelque chose à faire."""
    if not forcer and (aujourdhui.weekday() != 0 or _dernier_envoi['date'] == aujourdhui):
        return {'status': 'pas_lundi_ou_deja_envoye'}
    if not os.environ.get('SUPABASE_SERVICE_KEY'):
        return {'status': 'erreur', 'error': 'SUPABASE_SERVICE_KEY manquante'}
    groupes, brouillons, clients = contenu(aujourdhui)
    if not groupes and not brouillons:
        if not forcer:
            _dernier_envoi['date'] = aujourdhui
        return {'status': 'rien_a_faire'}
    sujet, corps = construire_mail(aujourdhui, groupes, brouillons, clients)
    m = requests.post('https://api.brevo.com/v3/smtp/email',
                      headers={'api-key': brevo_key, 'Content-Type': 'application/json'},
                      json={'sender': {'name': 'Swiss ViTa Form', 'email': DEST},
                            'to': [{'email': DEST, 'name': 'Swiss ViTa Form'}],
                            'subject': sujet, 'htmlContent': corps}, timeout=20)
    print(f'[BREVO rappel facturation] status={m.status_code} body={m.text[:200]}')
    if m.status_code != 201:
        return {'status': 'erreur', 'error': m.text[:200]}
    if not forcer:
        _dernier_envoi['date'] = aujourdhui
    return {'status': 'envoye', 'cours': len(groupes), 'brouillons': len(brouillons)}


def envoyer_sans_erreur(brevo_key):
    """Appelé depuis le déclencheur quotidien existant : ne doit jamais le faire échouer."""
    try:
        return envoyer(brevo_key, dt.datetime.now(TZ).date())
    except Exception as e:  # migration pas encore faite, Supabase indisponible…
        print(f'[rappel facturation] ignoré : {e}')
        return {'status': 'erreur', 'error': str(e)[:200]}


def register_rappel_facturation_routes(app, brevo_key):
    @app.route('/cron/facturation', methods=['GET', 'POST'])
    def cron_facturation():
        secret = os.environ.get('CRON_SECRET', '')
        recu = request.headers.get('X-Cron-Secret', '') or request.args.get('secret', '')
        if not secret:
            return jsonify({'error': 'CRON_SECRET non configuré'}), 503
        if not hmac.compare_digest(recu.encode(), secret.encode()):
            return jsonify({'error': 'non autorisé'}), 401
        try:
            return jsonify(envoyer(brevo_key, dt.datetime.now(TZ).date(), forcer=request.args.get('test') == '1'))
        except Exception as e:
            return jsonify({'error': str(e)[:300]}), 500
