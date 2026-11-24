"""Connexion sécurisée au portail (phase 1 du chantier sécurité).

- Les mots de passe ne sont plus comparés dans le navigateur : le serveur les vérifie.
- Ils sont stockés hachés (illisibles) dans la table `comptes_mdp`, que seule la clé
  service du serveur peut lire. La colonne `mot_de_passe` des tables admins / clients /
  formateurs est vidée au fur et à mesure (migration automatique, sans changer les mots
  de passe de personne).
- À la connexion, le serveur remet un jeton signé avec le secret JWT de Supabase :
  role=authenticated, app_role=admin|client|formateur, uid=<id>. C'est ce jeton que
  les règles d'accès (phase 2) utiliseront pour savoir qui voit quoi.

Variables Render : SUPABASE_URL, SUPABASE_SERVICE_KEY, SUPABASE_JWT_SECRET (secret JWT
« legacy » du projet Supabase : Project Settings → JWT Keys → Legacy JWT secret).
"""
import os
import threading
import time
from datetime import datetime, timedelta, timezone

import jwt
import requests
from flask import jsonify, request
from werkzeug.security import check_password_hash, generate_password_hash

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
DUREE_H = 12                 # durée d'une session
ESSAIS_MAX = 8               # échecs autorisés…
FENETRE_S = 15 * 60          # …par quart d'heure (par adresse IP et par identifiant)
TABLES = {'admin': 'admins', 'client': 'clients', 'formateur': 'formateurs'}

_echecs = {}
_verrou = threading.Lock()
_migration_lancee = False


# ---------------------------------------------------------------- Supabase
def _cle():
    return os.environ.get('SUPABASE_SERVICE_KEY', '')


def _secret():
    return os.environ.get('SUPABASE_JWT_SECRET', '')


def _sb(method, table, params=None, body=None, prefer=None):
    h = {'apikey': _cle(), 'Authorization': f'Bearer {_cle()}', 'Content-Type': 'application/json'}
    if prefer:
        h['Prefer'] = prefer
    r = requests.request(method, f'{SB_URL}/rest/v1/{table}', headers=h, params=params, json=body, timeout=20)
    if r.status_code >= 300:
        raise RuntimeError(f'Supabase {table} {r.status_code}: {r.text[:200]}')
    return r.json() if r.text else None


def _ref(type_, row):
    return str(row.get('id') if row.get('id') is not None else row.get('login'))


def _hash_de(type_, ref):
    rows = _sb('GET', 'comptes_mdp', {'type': f'eq.{type_}', 'ref_id': f'eq.{ref}', 'select': 'hash'})
    return rows[0]['hash'] if rows else None


def _ecrire_hash(type_, row, mdp):
    """Enregistre le hash, vérifie qu'il est bien écrit, PUIS vide l'ancien mot de passe en clair."""
    ref = _ref(type_, row)
    h = generate_password_hash(mdp)
    _sb('POST', 'comptes_mdp', {'on_conflict': 'type,ref_id'},
        {'type': type_, 'ref_id': ref, 'hash': h, 'maj': datetime.now(timezone.utc).isoformat()},
        prefer='resolution=merge-duplicates')
    if _hash_de(type_, ref) != h:
        raise RuntimeError('hash non enregistré')
    cle = {'id': f"eq.{row['id']}"} if row.get('id') is not None else {'login': f"eq.{row['login']}"}
    _sb('PATCH', TABLES[type_], cle, {'mot_de_passe': None})


def _trouver(login):
    """Même ordre que l'ancien portail : admin, puis client (email en minuscules), puis formateur."""
    login = (login or '').strip()
    if not login:
        return None, None
    for type_ in ('admin', 'client', 'formateur'):
        lg = login.lower() if (type_ == 'client' and '@' in login) else login
        rows = _sb('GET', TABLES[type_], {'login': f'eq.{lg}', 'select': '*', 'limit': '2'})
        if rows:
            return type_, rows[0]
    return None, None


def _charger(type_, uid):
    """Relit le compte à partir de l'uid du jeton (id, ou login pour un admin sans id)."""
    try:
        rows = _sb('GET', TABLES[type_], {'id': f'eq.{uid}', 'select': '*'})
        if rows:
            return rows[0]
    except Exception:
        pass
    if type_ == 'admin':
        rows = _sb('GET', 'admins', {'login': f'eq.{uid}', 'select': '*'})
        return rows[0] if rows else None
    return None


def _verifier(type_, row, mdp):
    h = _hash_de(type_, _ref(type_, row))
    if h:
        return check_password_hash(h, mdp)
    ancien = row.get('mot_de_passe')
    if ancien is not None and ancien != '' and ancien == mdp:
        try:
            _ecrire_hash(type_, row, mdp)    # migration à la volée
        except Exception as e:
            print('Migration mot de passe impossible :', e)
        return True
    return False


def migrer_tous():
    """Hache tous les mots de passe encore en clair (en arrière-plan, sans rien changer pour les utilisateurs)."""
    n = 0
    for type_, table in TABLES.items():
        try:
            rows = _sb('GET', table, {'mot_de_passe': 'not.is.null', 'select': '*'})
        except Exception as e:
            print('Migration', table, ':', e)
            continue
        for row in rows or []:
            if not row.get('mot_de_passe'):
                continue
            try:
                if _hash_de(type_, _ref(type_, row)):
                    # déjà haché : ne garder que le hash
                    cle = {'id': f"eq.{row['id']}"} if row.get('id') is not None else {'login': f"eq.{row['login']}"}
                    _sb('PATCH', table, cle, {'mot_de_passe': None})
                else:
                    _ecrire_hash(type_, row, row['mot_de_passe'])
                n += 1
            except Exception as e:
                print('Migration', table, ':', e)
    print(f'Migration des mots de passe : {n} compte(s) sécurisé(s)')
    return n


# ---------------------------------------------------------------- jetons
def creer_jeton(type_, row):
    maintenant = datetime.now(timezone.utc)
    uid = _ref(type_, row)
    return jwt.encode({
        'aud': 'authenticated', 'role': 'authenticated', 'iss': 'svf-portail',
        'sub': f'{type_}:{uid}', 'app_role': type_, 'uid': uid,
        'iat': int(maintenant.timestamp()), 'exp': int((maintenant + timedelta(hours=DUREE_H)).timestamp()),
    }, _secret(), algorithm='HS256')


def lire_jeton():
    """Renvoie les infos du jeton de la requête, ou None."""
    h = request.headers.get('Authorization', '')
    if not h.startswith('Bearer ') or not _secret():
        return None
    try:
        c = jwt.decode(h[7:], _secret(), algorithms=['HS256'], audience='authenticated')
    except jwt.PyJWTError:
        return None
    return c if c.get('iss') == 'svf-portail' else None


def _profil(type_, row):
    if type_ == 'admin':
        return {'role': 'admin', 'login': row.get('login')}
    if type_ == 'client':
        return {'role': 'client', 'login': row.get('login'), 'clientId': row.get('id'), 'clientNom': row.get('nom'),
                'doitChangerMdp': bool(row.get('doit_changer_mdp'))}
    return {'role': 'formateur', 'login': row.get('login'), 'formateurId': row.get('id'),
            'formateurNom': f"{row.get('prenom') or ''} {row.get('nom') or ''}".strip(),
            'accueilStatut': row.get('accueil_statut')}


def _reponse_session(type_, row):
    p = _profil(type_, row)
    p['token'] = creer_jeton(type_, row)
    p['exp'] = int(time.time()) + DUREE_H * 3600
    return p


# ---------------------------------------------------------------- limites
def _bloque(*cles):
    now = time.time()
    with _verrou:
        for k in cles:
            _echecs[k] = [t for t in _echecs.get(k, []) if now - t < FENETRE_S]
            if len(_echecs[k]) >= ESSAIS_MAX:
                return True
    return False


def _echec(*cles):
    with _verrou:
        for k in cles:
            _echecs.setdefault(k, []).append(time.time())


def _ip():
    return (request.headers.get('X-Forwarded-For', request.remote_addr or '').split(',')[0]).strip()


def mdp_probleme(m):
    if len(m or '') < 8:
        return 'au moins 8 caractères'
    if not any(c.isalpha() for c in m) or not any(c.isdigit() for c in m):
        return 'au moins une lettre et un chiffre'
    return None


# ---------------------------------------------------------------- routes
def register_auth_routes(app):

    @app.route('/auth/login', methods=['POST'])
    def auth_login():
        global _migration_lancee
        if not (_cle() and _secret()):
            return jsonify(error='Connexion sécurisée non configurée sur le serveur (SUPABASE_JWT_SECRET).'), 503
        d = request.get_json(silent=True) or {}
        login, mdp = (d.get('login') or '').strip(), d.get('mdp') or ''
        cles = ('ip:' + _ip(), 'lg:' + login.lower())
        if _bloque(*cles):
            return jsonify(error='Trop de tentatives. Réessayez dans 15 minutes.'), 429
        if not login or not mdp:
            return jsonify(error='Identifiants incorrects'), 401
        try:
            type_, row = _trouver(login)
            ok = bool(row) and _verifier(type_, row, mdp)
        except Exception as e:
            print('auth/login :', e)
            return jsonify(error='Serveur momentanément indisponible, réessayez.'), 502
        if not ok:
            _echec(*cles)
            time.sleep(0.5)
            return jsonify(error='Identifiants incorrects'), 401
        if type_ == 'admin' and not _migration_lancee:
            _migration_lancee = True
            threading.Thread(target=migrer_tous, daemon=True).start()
        return jsonify(_reponse_session(type_, row))

    @app.route('/auth/renouveler', methods=['POST'])
    def auth_renouveler():
        c = lire_jeton()
        if not c:
            return jsonify(error='Session expirée'), 401
        type_ = c['app_role']
        try:
            row = _charger(type_, c['uid'])
        except Exception:
            row = None
        if not row:
            return jsonify(error='Compte introuvable'), 401
        return jsonify(_reponse_session(type_, row))

    @app.route('/auth/changer-mdp', methods=['POST'])
    def auth_changer_mdp():
        """Changer SON propre mot de passe (client ou formateur connecté)."""
        c = lire_jeton()
        if not c:
            return jsonify(error='Session expirée, reconnectez-vous.'), 401
        d = request.get_json(silent=True) or {}
        nouveau = d.get('nouveau') or ''
        pb = mdp_probleme(nouveau)
        if pb:
            return jsonify(error=f'Le mot de passe doit contenir {pb}.'), 400
        type_ = c['app_role']
        row = _charger(type_, c['uid'])
        if not row:
            return jsonify(error='Compte introuvable'), 404
        if d.get('different_du_provisoire') and _verifier(type_, row, nouveau):
            return jsonify(error='Choisissez un mot de passe différent du mot de passe provisoire.'), 400
        _ecrire_hash(type_, row, nouveau)
        if type_ == 'client':
            _sb('PATCH', 'clients', {'id': f"eq.{row['id']}"}, {'doit_changer_mdp': False})
        return jsonify(ok=True)

    @app.route('/auth/definir-mdp', methods=['POST'])
    def auth_definir_mdp():
        """L'admin définit le mot de passe d'un client ou d'un formateur (création, réinitialisation)."""
        c = lire_jeton()
        if not c or c.get('app_role') != 'admin':
            return jsonify(error='Réservé à l\'administrateur'), 403
        d = request.get_json(silent=True) or {}
        type_, ident, mdp = d.get('type'), str(d.get('id') or ''), d.get('mdp') or ''
        if type_ not in ('client', 'formateur') or not ident or len(mdp) < 6:
            return jsonify(error='Données invalides (mot de passe : 6 caractères minimum)'), 400
        rows = _sb('GET', TABLES[type_], {'id': f'eq.{ident}', 'select': '*'})
        if not rows:
            return jsonify(error='Compte introuvable'), 404
        _ecrire_hash(type_, rows[0], mdp)
        return jsonify(ok=True)

    @app.route('/auth/etat', methods=['GET'])
    def auth_etat():
        """Contrôle de configuration (aucune donnée renvoyée)."""
        if not _secret():
            return jsonify(ok=False, probleme='SUPABASE_JWT_SECRET manquant dans Render'), 200
        if not _cle():
            return jsonify(ok=False, probleme='SUPABASE_SERVICE_KEY manquant dans Render'), 200
        try:
            _sb('GET', 'comptes_mdp', {'select': 'type', 'limit': '1'})
        except Exception:
            return jsonify(ok=False, probleme='Table comptes_mdp absente : lancer le SQL de la phase 1'), 200
        tok = creer_jeton('admin', {'login': '__test__'})
        r = requests.get(f'{SB_URL}/rest/v1/admins', params={'select': 'login', 'limit': '1'},
                         headers={'apikey': _cle(), 'Authorization': f'Bearer {tok}'}, timeout=20)
        if r.status_code == 401:
            return jsonify(ok=False, probleme='Supabase refuse le jeton : SUPABASE_JWT_SECRET incorrect'), 200
        return jsonify(ok=r.status_code == 200, supabase_http=r.status_code), 200
