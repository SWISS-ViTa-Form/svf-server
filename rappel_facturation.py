"""Rappel hebdomadaire (lundi) sur info@swissvf.ch :
- cours terminés sans facture (et frais d'annulation à facturer),
- factures créées mais pas encore envoyées,
- factures en retard de paiement,
- rappels d'échéance BLS en attente de validation.

Déclenché par l'appel quotidien existant (/cron/anniversaires) ou par /cron/facturation
(en-tête X-Cron-Secret ; ?test=1 pour envoyer tout de suite). Une seule fois par semaine :
la semaine du dernier envoi est mémorisée dans la table app_etat ; si le lundi est manqué
(serveur indisponible), le rappel part au premier appel suivant de la même semaine.
Requiert migration_facturation_cours.sql et migration_rappels.sql.
"""
import datetime as dt
import hmac
import html
import os
from zoneinfo import ZoneInfo

import requests
from flask import jsonify, request

import rappels_bls

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
TZ = ZoneInfo('Europe/Zurich')
DEST = 'info@swissvf.ch'
CLE_SEMAINE = 'rappel_lundi_semaine'
_memoire = {'semaine': None}


def _h():
    key = os.environ.get('SUPABASE_SERVICE_KEY', '')
    return {'apikey': key, 'Authorization': f'Bearer {key}'}


def _sb(table, params):
    r = requests.get(f'{SB_URL}/rest/v1/{table}', params=params, headers=_h(), timeout=30)
    r.raise_for_status()
    return r.json()


def _semaine(d):
    y, w, _ = d.isocalendar()
    return f'{y}-W{w:02d}'


def _deja_envoye(sem):
    try:
        rows = _sb('app_etat', {'cle': f'eq.{CLE_SEMAINE}', 'select': 'valeur'})
        return bool(rows) and rows[0].get('valeur') == sem
    except Exception:
        return _memoire['semaine'] == sem


def _marquer_envoye(sem):
    _memoire['semaine'] = sem
    try:
        requests.post(f'{SB_URL}/rest/v1/app_etat', json={'cle': CLE_SEMAINE, 'valeur': sem, 'updated_at': dt.datetime.utcnow().isoformat()},
                      headers={**_h(), 'Prefer': 'resolution=merge-duplicates,return=minimal'}, timeout=20)
    except Exception as e:
        print(f'[rappel lundi] mémorisation impossible : {e}')


def _fr(iso):
    s = str(iso or '')[:10]
    return f'{s[8:10]}.{s[5:7]}.{s[0:4]}' if len(s) == 10 else s


def _num(n):
    n = int(n or 0)
    return f'{n // 10000}-{n % 10000:04d}' if n >= 10000 else str(n)


def contenu(aujourdhui):
    cours = _sb('cours', {
        'select': 'id,client_id,type_cours,date_cours,heure_debut,heure_fin,statut,facturation',
        'date_cours': f'lt.{aujourdhui.isoformat()}', 'facture_id': 'is.null',
        'or': '(and(facturation.is.null,statut.is.null),and(facturation.is.null,statut.neq.annulé),facturation.eq.frais_annulation)',
        'order': 'date_cours.asc,heure_debut.asc'})
    brouillons = _sb('documents_commerciaux', {
        'select': 'numero,client_nom,titre', 'type': 'eq.facture', 'statut': 'eq.brouillon', 'order': 'numero.asc'})
    retards = _sb('documents_commerciaux', {
        'select': 'numero,client_nom,date_echeance,rappel_niveau', 'type': 'eq.facture', 'statut': 'eq.envoye',
        'date_echeance': f'lt.{aujourdhui.isoformat()}', 'order': 'date_echeance.asc'})
    clients = {c['id']: c.get('nom') or '' for c in _sb('clients', {'select': 'id,nom'})} if cours else {}
    try:
        nb_bls = rappels_bls.compter_a_envoyer(aujourdhui)
    except Exception as e:
        print(f'[rappel lundi] BLS ignoré : {e}')
        nb_bls = 0
    groupes = {}
    for c in cours:
        groupes.setdefault((c.get('client_id'), str(c['date_cours'])[:10]), []).append(c)
    return groupes, brouillons, retards, nb_bls, clients


def construire_mail(aujourdhui, groupes, brouillons, retards, nb_bls, clients):
    n1, n2, n3 = len(groupes), len(brouillons), len(retards)
    morceaux = [f'{n1} cours à facturer' if n1 else '', f"{n2} facture{'s' if n2 > 1 else ''} à envoyer" if n2 else '',
                f"{n3} en retard" if n3 else '', f"{nb_bls} rappel{'s' if nb_bls > 1 else ''} BLS" if nb_bls else '']
    sujet = '🧾 À faire cette semaine : ' + ' · '.join(x for x in morceaux if x)

    def age(date):
        j = (aujourdhui - dt.date.fromisoformat(date)).days
        return f' <span style="color:#c0392b">· il y a {j} jours</span>' if j > 14 else ''

    l1 = ''.join(
        f'<li><b>{html.escape(clients.get(cid) or "Sans client")}</b> — {_fr(date)} <span style="color:#888">('
        + html.escape(', '.join(("frais d'annulation " if c.get('facturation') == 'frais_annulation' else '')
                                + (c.get('type_cours') or '') + ' ' + (c.get('heure_debut') or '') for c in cs))
        + f')</span>{age(date)}</li>'
        for (cid, date), cs in sorted(groupes.items(), key=lambda x: x[0][1]))
    l2 = ''.join(f'<li><b>F-{_num(d["numero"])}</b> — {html.escape(d.get("client_nom") or "")} '
                 f'<span style="color:#888">{html.escape(d.get("titre") or "")}</span></li>' for d in brouillons)
    l3 = ''.join(
        f'<li><b>F-{_num(d["numero"])}</b> — {html.escape(d.get("client_nom") or "")} · échue le {_fr(d.get("date_echeance"))} '
        f'<span style="color:#c0392b">({(aujourdhui - dt.date.fromisoformat(str(d["date_echeance"])[:10])).days} j)</span>'
        + {0: ' — <b>1er rappel à envoyer</b>', 1: ' — 1er rappel envoyé', 2: ' — 2 rappels envoyés'}.get(d.get('rappel_niveau') or 0, '')
        + '</li>' for d in retards)
    corps = ('<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px"><p>Bonjour Vincent,</p>'
             + (f'<p><b>🧾 Cours terminés à facturer ({n1})</b></p><ul>{l1}</ul>' if n1 else '')
             + (f'<p><b>✉️ Factures créées mais pas encore envoyées ({n2})</b></p><ul>{l2}</ul>' if n2 else '')
             + (f'<p><b>⏰ Factures en retard de paiement ({n3})</b></p><ul>{l3}</ul>' if n3 else '')
             + (f'<p><b>🔔 Rappels d\'échéance BLS à valider : {nb_bls}</b> (onglet Clients &amp; certificats)</p>' if nb_bls else '')
             + '<p style="text-align:center;margin:22px 0"><a href="https://portail.swissvf.ch" style="background:#c0392b;color:#fff;'
               'padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Ouvrir le portail</a></p>'
             '<p style="color:#999;font-size:11px">Rappel automatique du portail Swiss ViTa Form, une fois par semaine (le lundi).</p></div>')
    return sujet, corps


def envoyer(brevo_key, aujourdhui, forcer=False):
    """Envoie le rappel une fois par semaine (dès le lundi), s'il y a quelque chose à faire."""
    sem = _semaine(aujourdhui)
    if not forcer and _deja_envoye(sem):
        return {'status': 'deja_envoye_cette_semaine'}
    if not os.environ.get('SUPABASE_SERVICE_KEY'):
        return {'status': 'erreur', 'error': 'SUPABASE_SERVICE_KEY manquante'}
    groupes, brouillons, retards, nb_bls, clients = contenu(aujourdhui)
    if not (groupes or brouillons or retards or nb_bls):
        if not forcer:
            _marquer_envoye(sem)
        return {'status': 'rien_a_faire'}
    sujet, corps = construire_mail(aujourdhui, groupes, brouillons, retards, nb_bls, clients)
    m = requests.post('https://api.brevo.com/v3/smtp/email',
                      headers={'api-key': brevo_key, 'Content-Type': 'application/json'},
                      json={'sender': {'name': 'Swiss ViTa Form', 'email': DEST},
                            'to': [{'email': DEST, 'name': 'Swiss ViTa Form'}],
                            'subject': sujet, 'htmlContent': corps}, timeout=20)
    print(f'[BREVO rappel lundi] status={m.status_code} body={m.text[:200]}')
    if m.status_code != 201:
        return {'status': 'erreur', 'error': m.text[:200]}
    if not forcer:
        _marquer_envoye(sem)
    return {'status': 'envoye', 'cours': len(groupes), 'brouillons': len(brouillons), 'retards': len(retards), 'bls': nb_bls}


def envoyer_sans_erreur(brevo_key):
    """Appelé depuis le déclencheur quotidien existant : ne doit jamais le faire échouer."""
    try:
        return envoyer(brevo_key, dt.datetime.now(TZ).date())
    except Exception as e:
        print(f'[rappel lundi] ignoré : {e}')
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
