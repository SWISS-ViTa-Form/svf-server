"""Dossier personnel sécurisé des formateurs — Swiss ViTa Form.

Sécurité :
- accès uniquement via ce serveur (clé service_role en variable d'environnement, jamais dans le code)
- mot de passe de dossier distinct du mot de passe du portail, haché (jamais en clair)
- activation par code à usage unique remis par l'admin
- jeton de session de 15 minutes ; blocage après 5 échecs
- infos personnelles (AVS, IBAN, adresse, banque...) chiffrées (Fernet) avant stockage
- fichiers dans un bucket privé, servis uniquement après contrôle des droits
- document déposé = verrouillé pour le formateur tant que l'admin ne le déverrouille pas
- journal de tous les accès
"""
import os, json, time, hmac, uuid, re, secrets
from datetime import datetime, timedelta, timezone
from functools import wraps

import jwt
import requests
from cryptography.fernet import Fernet, InvalidToken
from flask import request, jsonify, g, Response
from werkzeug.security import generate_password_hash, check_password_hash

SB_URL = os.environ.get('SUPABASE_URL', 'https://weguamuhrillncibedap.supabase.co').rstrip('/')
SB_KEY = os.environ.get('SUPABASE_SERVICE_KEY', '')
JWT_SECRET = os.environ.get('JWT_SECRET', '')
ENC_KEY = os.environ.get('DOSSIER_ENC_KEY', '')
ADMIN_PW = os.environ.get('ADMIN_DOSSIER_PASSWORD', '')
BUCKET = 'dossiers-formateurs'
TOKEN_MIN = 15
MAX_BYTES = 10 * 1024 * 1024

# (clé, libellé, libellé court, date d'expiration obligatoire)
FORMATEUR_TYPES = [
    ('cert_instructeur', 'Certificat / attestation instructeur BLS-AED-SRC', 'Instructeur', False),
    ('maj_instructeur', 'Mise à jour BLS-AED-SRC instructeur', 'Mise à jour', True),
    ('diplome_pro', 'Diplôme ou brevet : ambulancier / technicien ambulancier / infirmier', 'Diplôme', False),
    ('autre_formateur', 'Autres documents (à nommer)', 'Autre', False),
]
ADMIN_TYPES = [
    ('contrat', 'Contrat de travail signé', 'Contrat', False),
    ('reglement', 'Règlement signé', 'Règlement', False),
    ('avenant', 'Avenant au contrat', 'Avenant', False),
    ('autre_admin', 'Autres documents (à nommer)', 'Autre', False),
]
FORM_KEYS = {t[0] for t in FORMATEUR_TYPES}
ALL_KEYS = FORM_KEYS | {t[0] for t in ADMIN_TYPES}
# Tous les types acceptent PLUSIEURS documents (ajout toujours possible, jamais de remplacement).
MULTI_KEYS = set(ALL_KEYS)
# "Autres documents" : un nom est obligatoire
OTHER_KEYS = {'autre_formateur', 'autre_admin'}
EXPIRY_REQUIRED = {t[0] for t in FORMATEUR_TYPES + ADMIN_TYPES if t[3]}
PROFIL_FIELDS = ['civilite', 'date_naissance', 'adresse', 'npa_localite', 'telephone',
                 'avs_no', 'iban', 'banque_nom', 'banque_adresse']
DATE_RE = re.compile(r'\d{4}-\d{2}-\d{2}')

try:
    _fernet = Fernet(ENC_KEY.encode()) if ENC_KEY else None
except Exception:
    _fernet = None
CONFIG_OK = bool(SB_KEY and ADMIN_PW and _fernet and len(JWT_SECRET) >= 32)


class DossierError(Exception):
    def __init__(self, msg, code=400):
        self.msg, self.code = msg, code


def _err(msg, code=400):
    return jsonify({'error': msg}), code


def types_payload():
    f = lambda L: [{'key': k, 'label': l, 'court': c, 'expiry': e, 'multi': True, 'named': k in OTHER_KEYS} for k, l, c, e in L]
    return {'formateur': f(FORMATEUR_TYPES), 'admin': f(ADMIN_TYPES)}


# ---------- Supabase (clé service_role, côté serveur uniquement) ----------
def sb(method, table, params=None, body=None, prefer=None):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY, 'Content-Type': 'application/json'}
    if prefer:
        h['Prefer'] = prefer
    r = requests.request(method, f'{SB_URL}/rest/v1/{table}', headers=h, params=params,
                         data=json.dumps(body) if body is not None else None, timeout=20)
    if not r.ok:
        raise RuntimeError(f'supabase {r.status_code}: {r.text[:200]}')
    return r.json() if r.text else None


def _st_headers(mime=None):
    h = {'apikey': SB_KEY, 'Authorization': 'Bearer ' + SB_KEY}
    if mime:
        h['Content-Type'] = mime
    return h


def put_object(path, data, mime):
    r = requests.post(f'{SB_URL}/storage/v1/object/{BUCKET}/{path}', headers=_st_headers(mime), data=data, timeout=60)
    if not r.ok:
        raise RuntimeError(f'storage put {r.status_code}: {r.text[:200]}')


def get_object(path):
    r = requests.get(f'{SB_URL}/storage/v1/object/{BUCKET}/{path}', headers=_st_headers(), timeout=60)
    if not r.ok:
        raise RuntimeError(f'storage get {r.status_code}')
    return r.content


def delete_object(path):
    try:
        requests.delete(f'{SB_URL}/storage/v1/object/{BUCKET}/{path}', headers=_st_headers(), timeout=20)
    except Exception:
        pass


# ---------- utilitaires ----------
def _now():
    return datetime.now(timezone.utc)


def _ip():
    return (request.headers.get('X-Forwarded-For', '').split(',')[0].strip() or request.remote_addr or '')


def _uuid(s):
    try:
        return str(uuid.UUID(str(s)))
    except Exception:
        raise DossierError('Identifiant invalide.', 400)


def _token(role, sub=None):
    return jwt.encode({'role': role, 'sub': sub or role, 'exp': _now() + timedelta(minutes=TOKEN_MIN)}, JWT_SECRET, algorithm='HS256')


def log(acteur, acteur_id, fid, action, doc_id=None):
    try:
        sb('POST', 'dossier_acces_log', body={'acteur': acteur, 'acteur_id': acteur_id, 'formateur_id': fid,
                                              'action': action, 'doc_id': doc_id, 'ip': _ip()}, prefer='return=minimal')
    except Exception as e:
        print('[dossier] journal indisponible:', repr(e))


_fails = {}


def _locked(key, limit):
    e = _fails.get(key)
    if not e:
        return False
    n, until = e
    if until and until <= time.time():
        _fails.pop(key, None)
        return False
    return n >= limit and until > time.time()


def _fail(key, limit):
    n, until = _fails.get(key, (0, 0))
    n += 1
    _fails[key] = (n, time.time() + 900 if n >= limit else until)


def _ok(key):
    _fails.pop(key, None)


def enc(obj):
    return _fernet.encrypt(json.dumps(obj, ensure_ascii=False).encode()).decode()


def dec(txt):
    try:
        return json.loads(_fernet.decrypt(txt.encode()).decode())
    except (InvalidToken, ValueError):
        raise DossierError('Données illisibles (clé de chiffrement différente ?).', 500)


def clean_profil(d, strict=True):
    out = {k: str(d.get(k) or '').strip()[:200] for k in PROFIL_FIELDS}
    if not strict:
        return out
    if out['avs_no'] and not re.fullmatch(r'756\.\d{4}\.\d{4}\.\d{2}', out['avs_no']):
        raise DossierError('N° AVS invalide (format 756.1234.5678.90).')
    if out['iban']:
        iban = re.sub(r'\s', '', out['iban']).upper()
        if not re.fullmatch(r'[A-Z]{2}\d{2}[A-Z0-9]{11,30}', iban):
            raise DossierError('IBAN invalide.')
        out['iban'] = ' '.join(iban[i:i + 4] for i in range(0, len(iban), 4))
    if out['date_naissance'] and not DATE_RE.fullmatch(out['date_naissance']):
        raise DossierError('Date de naissance invalide.')
    return out


def _sniff(data):
    if data[:5] == b'%PDF-':
        return 'pdf', 'application/pdf'
    if data[:3] == b'\xff\xd8\xff':
        return 'jpg', 'image/jpeg'
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png', 'image/png'
    return None, None


def _formateur_by_login(login):
    rows = sb('GET', 'formateurs', {'login': 'eq.' + login, 'select': 'id,prenom,nom,actif'}) or []
    if rows and rows[0].get('actif') is not False:
        return rows[0]
    return None


def _auth_row(fid):
    rows = sb('GET', 'formateur_dossier_auth', {'formateur_id': 'eq.' + fid}) or []
    return rows[0] if rows else None


def _active_doc(fid, dtype):
    rows = sb('GET', 'formateur_dossier_docs', {'formateur_id': 'eq.' + fid, 'type': 'eq.' + dtype, 'archive': 'eq.false'}) or []
    return rows[0] if rows else None


DOC_COLS = 'id,formateur_id,type,libelle,nom_original,mime,taille,date_expiration,depose_par,deverrouille,archive,created_at'


def _docs(fid, with_archive=False):
    p = {'formateur_id': 'eq.' + fid, 'select': DOC_COLS, 'order': 'created_at.desc'}
    if not with_archive:
        p['archive'] = 'eq.false'
    return sb('GET', 'formateur_dossier_docs', p) or []


def _get_profil(fid):
    rows = sb('GET', 'formateur_dossier_profil', {'formateur_id': 'eq.' + fid}) or []
    if not rows:
        return {}, None
    return dec(rows[0]['profil_enc']), rows[0].get('updated_at')


def _put_profil(fid, profil, by):
    sb('POST', 'formateur_dossier_profil', body={'formateur_id': fid, 'profil_enc': enc(profil),
                                                  'updated_at': _now().isoformat(), 'updated_by': by},
       prefer='resolution=merge-duplicates,return=minimal')


def _save_doc(fid, dtype, f, depose_par, date_exp, libelle=''):
    if dtype not in ALL_KEYS:
        raise DossierError('Type de document inconnu.')
    if depose_par == 'formateur' and dtype not in FORM_KEYS:
        raise DossierError('Ce document est ajouté par Swiss ViTa Form.', 403)
    if date_exp and not DATE_RE.fullmatch(date_exp):
        raise DossierError("Date d'expiration invalide.")
    if dtype in EXPIRY_REQUIRED and not date_exp:
        raise DossierError("Indiquez la date d'expiration.")
    libelle = re.sub(r'\s+', ' ', libelle or '').strip()[:80]
    if dtype in OTHER_KEYS and not libelle:
        raise DossierError('Donnez un nom à ce document.')
    if not f:
        raise DossierError('Aucun fichier reçu.')
    data = f.read(MAX_BYTES + 1)
    if not data:
        raise DossierError('Fichier vide.')
    if len(data) > MAX_BYTES:
        raise DossierError('Fichier trop volumineux (10 Mo maximum).', 413)
    ext, mime = _sniff(data)
    if not ext:
        raise DossierError('Format non accepté (PDF, JPG ou PNG uniquement).')
    actif = None if dtype in MULTI_KEYS else _active_doc(fid, dtype)
    if actif and depose_par == 'formateur' and not actif.get('deverrouille'):
        raise DossierError('Document verrouillé. Pour le modifier, contactez Swiss ViTa Form.', 403)
    path = f'{fid}/{dtype}/{uuid.uuid4().hex}.{ext}'
    put_object(path, data, mime)
    try:
        if actif:
            sb('PATCH', 'formateur_dossier_docs', {'id': 'eq.' + actif['id']}, {'archive': True, 'deverrouille': False},
               prefer='return=minimal')
        row = sb('POST', 'formateur_dossier_docs', body={
            'formateur_id': fid, 'type': dtype, 'libelle': libelle or None, 'storage_path': path,
            'nom_original': re.sub(r'[^\w.\- ]', '_', (f.filename or 'document'))[:120],
            'mime': mime, 'taille': len(data), 'date_expiration': date_exp or None,
            'depose_par': depose_par}, prefer='return=representation')
    except Exception:
        try:
            if actif:
                sb('PATCH', 'formateur_dossier_docs', {'id': 'eq.' + actif['id']}, {'archive': False}, prefer='return=minimal')
        except Exception:
            pass
        delete_object(path)
        raise
    d = row[0]
    d.pop('storage_path', None)
    return d


def _send_file(doc):
    data = get_object(doc['storage_path'])
    ext = {'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png'}.get(doc['mime'], 'bin')
    return Response(data, mimetype=doc['mime'], headers={
        'Content-Disposition': f'inline; filename="{doc["type"]}.{ext}"',
        'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'})


def register_dossier_routes(app):
    P = '/dossier'

    def guard(role=None):
        def deco(fn):
            @wraps(fn)
            def w(*a, **k):
                if not CONFIG_OK:
                    return _err('Service dossier non configuré sur le serveur.', 503)
                if role:
                    h = request.headers.get('Authorization', '')
                    if not h.startswith('Bearer '):
                        return _err('Non autorisé.', 401)
                    try:
                        c = jwt.decode(h[7:], JWT_SECRET, algorithms=['HS256'])
                    except Exception:
                        return _err('Session expirée.', 401)
                    if c.get('role') != role:
                        return _err('Accès refusé.', 403)
                    g.claims = c
                try:
                    return fn(*a, **k)
                except DossierError as e:
                    return _err(e.msg, e.code)
                except Exception as e:
                    print('[dossier] erreur:', repr(e))
                    return _err('Erreur serveur.', 500)
            return w
        return deco

    def body():
        return request.get_json(silent=True) or {}

    def upload_guard():
        if request.content_length and request.content_length > MAX_BYTES + 200 * 1024:
            raise DossierError('Fichier trop volumineux (10 Mo maximum).', 413)

    # ---------------- connexion formateur ----------------
    @app.route(P + '/login', methods=['POST'])
    @guard()
    def d_login():
        b = body()
        login, pw = str(b.get('login') or '').strip(), str(b.get('mot_de_passe') or '')
        k1, k2 = f'{_ip()}|{login}', f'ip|{_ip()}'
        if _locked(k1, 5) or _locked(k2, 20):
            raise DossierError('Trop de tentatives. Réessayez dans 15 minutes.', 429)
        f = _formateur_by_login(login) if login else None
        a = _auth_row(f['id']) if f else None
        if not f or not a or not a.get('pwd_hash') or not check_password_hash(a['pwd_hash'], pw):
            _fail(k1, 5); _fail(k2, 20)
            if f and not (a and a.get('pwd_hash')):
                raise DossierError("Dossier non activé : demandez un code d'activation à Swiss ViTa Form.", 403)
            raise DossierError('Identifiants incorrects.', 401)
        _ok(k1)
        log('formateur', f['id'], f['id'], 'connexion')
        return jsonify({'token': _token('formateur', f['id'])})

    @app.route(P + '/activer', methods=['POST'])
    @guard()
    def d_activer():
        b = body()
        login = str(b.get('login') or '').strip()
        code = re.sub(r'[\s-]', '', str(b.get('code') or '')).upper()
        pw = str(b.get('mot_de_passe') or '')
        k1, k2 = f'act|{_ip()}|{login}', f'ip|{_ip()}'
        if _locked(k1, 5) or _locked(k2, 20):
            raise DossierError('Trop de tentatives. Réessayez dans 15 minutes.', 429)
        if len(pw) < 12:
            raise DossierError('Mot de passe trop court (12 caractères minimum).')
        f = _formateur_by_login(login) if login else None
        a = _auth_row(f['id']) if f else None
        valide = False
        if f and a and a.get('code_hash') and a.get('code_expire'):
            expire = datetime.fromisoformat(a['code_expire'].replace('Z', '+00:00'))
            if expire > _now() and (a.get('tentatives') or 0) < 5 and check_password_hash(a['code_hash'], code):
                valide = True
            elif f and expire > _now():
                n = (a.get('tentatives') or 0) + 1
                patch = {'tentatives': n}
                if n >= 5:
                    patch['code_hash'] = None
                sb('PATCH', 'formateur_dossier_auth', {'formateur_id': 'eq.' + f['id']}, patch, prefer='return=minimal')
        if not valide:
            _fail(k1, 5); _fail(k2, 20)
            raise DossierError("Code invalide ou expiré. Demandez-en un nouveau à Swiss ViTa Form.", 401)
        sb('PATCH', 'formateur_dossier_auth', {'formateur_id': 'eq.' + f['id']},
           {'pwd_hash': generate_password_hash(pw), 'code_hash': None, 'code_expire': None, 'tentatives': 0,
            'updated_at': _now().isoformat()}, prefer='return=minimal')
        _ok(k1)
        log('formateur', f['id'], f['id'], 'activation')
        return jsonify({'token': _token('formateur', f['id'])})

    # ---------------- espace formateur ----------------
    @app.route(P + '/me', methods=['GET'])
    @guard('formateur')
    def d_me():
        fid = g.claims['sub']
        profil, upd = _get_profil(fid)
        log('formateur', fid, fid, 'consultation_dossier')
        return jsonify({'profil': profil, 'profil_updated': upd, 'docs': _docs(fid), 'types': types_payload()})

    @app.route(P + '/me/profil', methods=['PUT'])
    @guard('formateur')
    def d_me_profil():
        fid = g.claims['sub']
        _put_profil(fid, clean_profil(body()), 'formateur')
        log('formateur', fid, fid, 'modif_profil')
        return jsonify({'status': 'ok'})

    @app.route(P + '/me/docs/<dtype>', methods=['POST'])
    @guard('formateur')
    def d_me_upload(dtype):
        fid = g.claims['sub']
        upload_guard()
        d = _save_doc(fid, dtype, request.files.get('fichier'), 'formateur', (request.form.get('date_expiration') or '').strip(), request.form.get('libelle'))
        log('formateur', fid, fid, 'depot_' + dtype, d['id'])
        return jsonify(d)

    @app.route(P + '/me/docs/<doc_id>/fichier', methods=['GET'])
    @guard('formateur')
    def d_me_fichier(doc_id):
        fid = g.claims['sub']
        rows = sb('GET', 'formateur_dossier_docs', {'id': 'eq.' + _uuid(doc_id)}) or []
        if not rows or rows[0]['formateur_id'] != fid or rows[0]['archive']:
            raise DossierError('Document introuvable.', 404)
        log('formateur', fid, fid, 'ouverture_doc', doc_id)
        return _send_file(rows[0])

    # ---------------- admin ----------------
    @app.route(P + '/admin/login', methods=['POST'])
    @guard()
    def d_admin_login():
        k1 = f'adm|{_ip()}'
        if _locked(k1, 5):
            raise DossierError('Trop de tentatives. Réessayez dans 15 minutes.', 429)
        pw = str(body().get('mot_de_passe') or '')
        if not hmac.compare_digest(pw.encode(), ADMIN_PW.encode()):
            _fail(k1, 5)
            raise DossierError('Mot de passe incorrect.', 401)
        _ok(k1)
        log('admin', 'admin', None, 'connexion_admin')
        return jsonify({'token': _token('admin')})

    @app.route(P + '/admin/apercu', methods=['GET'])
    @guard('admin')
    def d_admin_apercu():
        forms = sb('GET', 'formateurs', {'actif': 'eq.true', 'select': 'id,prenom,nom,login,email', 'order': 'nom.asc'}) or []
        auths = {a['formateur_id']: a for a in (sb('GET', 'formateur_dossier_auth', {'select': 'formateur_id,pwd_hash,code_expire,code_hash'}) or [])}
        profils = {p['formateur_id']: p['updated_at'] for p in (sb('GET', 'formateur_dossier_profil', {'select': 'formateur_id,updated_at'}) or [])}
        docs = sb('GET', 'formateur_dossier_docs', {'archive': 'eq.false', 'select': DOC_COLS}) or []
        by = {}
        for d in docs:
            if d['type'] in OTHER_KEYS:
                continue
            by.setdefault(d['formateur_id'], {})[d['type']] = {k: d[k] for k in ('id', 'date_expiration', 'deverrouille', 'depose_par', 'created_at')}
        out = []
        for f in forms:
            a = auths.get(f['id']) or {}
            out.append({**f, 'active': bool(a.get('pwd_hash')), 'code_pending': bool(a.get('code_hash')),
                        'profil_updated': profils.get(f['id']), 'docs': by.get(f['id'], {})})
        return jsonify({'formateurs': out, 'types': types_payload()})

    @app.route(P + '/admin/formateur/<fid>', methods=['GET'])
    @guard('admin')
    def d_admin_detail(fid):
        fid = _uuid(fid)
        rows = sb('GET', 'formateurs', {'id': 'eq.' + fid, 'select': 'id,prenom,nom'}) or []
        if not rows:
            raise DossierError('Formateur introuvable.', 404)
        profil, upd = _get_profil(fid)
        log('admin', 'admin', fid, 'consultation_dossier')
        return jsonify({'formateur': rows[0], 'profil': profil, 'profil_updated': upd,
                        'docs': _docs(fid, True), 'types': types_payload()})

    @app.route(P + '/admin/formateur/<fid>/profil', methods=['PUT'])
    @guard('admin')
    def d_admin_profil(fid):
        fid = _uuid(fid)
        _put_profil(fid, clean_profil(body()), 'admin')
        log('admin', 'admin', fid, 'modif_profil')
        return jsonify({'status': 'ok'})

    @app.route(P + '/admin/formateur/<fid>/docs/<dtype>', methods=['POST'])
    @guard('admin')
    def d_admin_upload(fid, dtype):
        fid = _uuid(fid)
        upload_guard()
        d = _save_doc(fid, dtype, request.files.get('fichier'), 'admin', (request.form.get('date_expiration') or '').strip(), request.form.get('libelle'))
        log('admin', 'admin', fid, 'depot_' + dtype, d['id'])
        return jsonify(d)

    @app.route(P + '/admin/docs/<doc_id>/fichier', methods=['GET'])
    @guard('admin')
    def d_admin_fichier(doc_id):
        rows = sb('GET', 'formateur_dossier_docs', {'id': 'eq.' + _uuid(doc_id)}) or []
        if not rows:
            raise DossierError('Document introuvable.', 404)
        log('admin', 'admin', rows[0]['formateur_id'], 'ouverture_doc', doc_id)
        return _send_file(rows[0])

    @app.route(P + '/admin/docs/<doc_id>/deverrouiller', methods=['POST'])
    @guard('admin')
    def d_admin_unlock(doc_id):
        rows = sb('GET', 'formateur_dossier_docs', {'id': 'eq.' + _uuid(doc_id), 'archive': 'eq.false'}) or []
        if not rows:
            raise DossierError('Document introuvable.', 404)
        sb('PATCH', 'formateur_dossier_docs', {'id': 'eq.' + doc_id}, {'deverrouille': True}, prefer='return=minimal')
        log('admin', 'admin', rows[0]['formateur_id'], 'deverrouillage', doc_id)
        return jsonify({'status': 'ok'})

    @app.route(P + '/admin/docs/<doc_id>', methods=['DELETE'])
    @guard('admin')
    def d_admin_archive(doc_id):
        rows = sb('GET', 'formateur_dossier_docs', {'id': 'eq.' + _uuid(doc_id)}) or []
        if not rows:
            raise DossierError('Document introuvable.', 404)
        sb('PATCH', 'formateur_dossier_docs', {'id': 'eq.' + doc_id}, {'archive': True, 'deverrouille': False}, prefer='return=minimal')
        log('admin', 'admin', rows[0]['formateur_id'], 'archivage', doc_id)
        return jsonify({'status': 'ok'})

    @app.route(P + '/admin/formateur/<fid>/code', methods=['POST'])
    @guard('admin')
    def d_admin_code(fid):
        fid = _uuid(fid)
        alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
        code = ''.join(secrets.choice(alphabet) for _ in range(10))
        expire = _now() + timedelta(days=7)
        # Générer un code désactive le mot de passe actuel (réinitialisation) jusqu'à nouvelle activation
        sb('POST', 'formateur_dossier_auth', body={'formateur_id': fid, 'pwd_hash': None, 'code_hash': generate_password_hash(code),
                                                    'code_expire': expire.isoformat(), 'tentatives': 0, 'updated_at': _now().isoformat()},
           prefer='resolution=merge-duplicates,return=minimal')
        log('admin', 'admin', fid, 'code_activation')
        return jsonify({'code': code[:5] + '-' + code[5:], 'expire': expire.isoformat()})

    @app.route(P + '/admin/journal', methods=['GET'])
    @guard('admin')
    def d_admin_journal():
        p = {'order': 'at.desc', 'limit': str(min(int(request.args.get('limit', 100) or 100), 300))}
        if request.args.get('formateur_id'):
            p['formateur_id'] = 'eq.' + _uuid(request.args['formateur_id'])
        return jsonify({'journal': sb('GET', 'dossier_acces_log', p) or []})

    @app.route(P + '/admin/migrer-legacy', methods=['POST'])
    @guard('admin')
    def d_admin_migrer():
        """Copie (chiffrée) les infos déjà saisies dans la table formateurs vers le dossier. Ne supprime rien."""
        rows = sb('GET', 'formateurs', {'select': 'id,civilite,date_naissance,adresse,npa_localite,telephone,avs_no,iban'}) or []
        deja = {p['formateur_id'] for p in (sb('GET', 'formateur_dossier_profil', {'select': 'formateur_id'}) or [])}
        n = 0
        for r in rows:
            if r['id'] in deja:
                continue
            profil = clean_profil({k: r.get(k) for k in PROFIL_FIELDS}, strict=False)
            if not any(profil.values()):
                continue
            _put_profil(r['id'], profil, 'migration')
            n += 1
        log('admin', 'admin', None, 'migration_legacy')
        return jsonify({'migres': n, 'deja_presents': len(deja)})
