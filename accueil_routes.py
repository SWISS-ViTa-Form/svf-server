"""Parcours d'accueil des nouveaux instructeurs — Swiss ViTa Form.

Étapes côté instructeur (dès sa première connexion, le reste du portail est bloqué) :
  1. lecture obligatoire : règlement, marche à suivre, ce que nous proposons (« J'ai lu » horodaté)
  2. fiche d'informations (enregistrée chiffrée dans son dossier sécurisé) + diplômes
  3. contrat de travail, règlement et avenant générés en PDF (comme les certificats),
     à signer, dater et redéposer
Ensuite : l'admin contresigne et valide.

Toutes les données personnelles passent par le dossier sécurisé (chiffrement, jeton 15 min).
Chaque réponse renvoie un nouveau jeton (expiration glissante) pour que la saisie du
formulaire ne coupe pas la session.
"""
import io, os, re, secrets, zipfile
from datetime import datetime, timedelta, date
from functools import wraps
from xml.sax.saxutils import escape

import jwt
import requests
from flask import request, jsonify, g, Response
from werkzeug.security import generate_password_hash
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

import dossier_routes as D

HERE = os.path.dirname(os.path.abspath(__file__))
MODELES = os.path.join(HERE, 'modeles')
def _chemin(nom):
    # Cherche le modèle dans modeles/ puis à la racine du dépôt
    p = os.path.join(MODELES, nom)
    return p if os.path.isfile(p) else os.path.join(HERE, nom)


F_CONTRAT = _chemin('contrat_travail.docx')
F_AVENANT = _chemin('avenant_apport.docx')
F_REGLEMENT = _chemin('reglement_instructeurs.pdf')
ADMIN_MAIL = os.environ.get('ADMIN_NOTIF_EMAIL', 'info@swissvf.ch')

LECTURES = {
    'reglement': ('Règlement des instructeurs', None),
    'marche': ('Marche à suivre', re.compile(r'marche\s*[àa]\s*suivre', re.I)),
    'offre': ('Ce que nous proposons', re.compile(r'ce\s*que\s*nous\s*proposons', re.I)),
}
QUALIFS = {
    'ambulancier': ('Ambulancier / ambulancière diplômé·e ES', 'diplome_pro'),
    'technicien': ('Technicien / technicienne ambulancier·ère', 'diplome_pro'),
    'infirmier': ('Infirmier / infirmière', 'diplome_pro'),
    'formateur_bls': ('Formateur / formatrice BLS-AED SRC', 'cert_instructeur'),
}
REQUIS = {
    'civilite': 'civilité', 'date_naissance': 'date de naissance', 'adresse': 'rue et numéro',
    'npa_localite': 'NPA / localité', 'telephone': 'téléphone', 'avs_no': 'N° AVS',
    'nationalite': 'nationalité', 'etat_civil': 'état civil', 'iban': 'IBAN',
    'banque_nom': 'nom de la banque', 'banque_titulaire': 'titulaire du compte',
    'taille_tshirt': 'taille de t-shirt',
}
SIGNES = ['contrat', 'reglement', 'avenant']
MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août',
        'septembre', 'octobre', 'novembre', 'décembre']


# ---------------- utilitaires ----------------
def _d(iso):
    """AAAA-MM-JJ -> JJ.MM.AAAA"""
    try:
        return datetime.strptime(str(iso)[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
    except Exception:
        return ''


def _d_long(dt):
    return f'{dt.day} {MOIS[dt.month - 1]} {dt.year}'


def _chf(v):
    return f'{float(v):.2f}'


def _taux(total):
    total = round(float(total or 0), 2)
    base = round(total / 1.0833, 2)
    return base, round(total - base, 2), total


def _get_accueil(fid):
    rows = D.sb('GET', 'formateur_accueil', {'formateur_id': 'eq.' + fid}) or []
    return rows[0] if rows else None


def _patch_accueil(fid, patch):
    D.sb('PATCH', 'formateur_accueil', {'formateur_id': 'eq.' + fid}, patch, prefer='return=minimal')


def _formateur(fid):
    rows = D.sb('GET', 'formateurs', {'id': 'eq.' + fid, 'select': 'id,prenom,nom,email,login,accueil_statut'}) or []
    if not rows:
        raise D.DossierError('Instructeur introuvable.', 404)
    return rows[0]


def reglement_version():
    try:
        txt = PdfReader(F_REGLEMENT).pages[0].extract_text() or ''
        m = re.search(r'Version du ([^\n·]+?\d{4})', txt)
        return m.group(1).strip() if m else ''
    except Exception:
        return ''


def _docs_partages():
    """Marche à suivre / ce que nous proposons : pris dans les documents communs du portail."""
    rows = D.sb('GET', 'formateur_documents', {'formateur_id': 'is.null',
                                               'select': 'nom_affichage,nom_fichier,url_storage,created_at',
                                               'order': 'created_at.desc'}) or []
    out = {}
    for key, (_, rx) in LECTURES.items():
        if not rx:
            continue
        for r in rows:
            if rx.search(str(r.get('nom_affichage') or r.get('nom_fichier') or '')):
                out[key] = r.get('url_storage')
                break
    return out


def _etat(fid, acc=None, profil=None, docs=None):
    """Calcule l'avancement des 3 étapes."""
    acc = acc or _get_accueil(fid) or {}
    if profil is None:
        profil, _ = D._get_profil(fid)
    docs = docs if docs is not None else D._docs(fid)
    lect = acc.get('lectures') or {}
    lect_ok = all(k in lect for k in LECTURES)
    manque = [lbl for k, lbl in REQUIS.items() if not str(profil.get(k) or '').strip()]
    types_docs = {d['type'] for d in docs}
    quals = [q for q in str(profil.get('qualifications') or '').split(',') if q in QUALIFS]
    for q in quals:
        if QUALIFS[q][1] not in types_docs:
            manque.append('copie du diplôme : ' + QUALIFS[q][0])
    requis_signes = SIGNES if acc.get('apporteur') else SIGNES[:2]
    depuis = acc.get('created_at') or ''
    signes = {t: [d for d in docs if d['type'] == t and d.get('depose_par') == 'formateur'
                  and (d.get('created_at') or '') >= depuis] for t in requis_signes}
    return {
        'lectures_ok': lect_ok,
        'infos_ok': not manque and bool(acc.get('infos_completes_at')),
        'infos_manque': manque,
        'requis_signes': requis_signes,
        'signes': {t: len(v) for t, v in signes.items()},
        'signes_ok': all(signes[t] for t in requis_signes),
    }


def fill_docx(path, values):
    """Remplace les jetons {{X}} dans le Word (corps, zones de texte, pied de page)."""
    zin = zipfile.ZipFile(path)
    out = io.BytesIO()
    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename.startswith('word/') and it.filename.endswith('.xml'):
            x = data.decode('utf-8')
            for k, v in values.items():
                x = x.replace('{{' + k + '}}', escape(str(v or '')))
            x = re.sub(r'\{\{[A-Z_]+\}\}', '', x)
            data = x.encode('utf-8')
        zout.writestr(it, data)
    zout.close()
    return out.getvalue()


def _valeurs(f, profil, acc):
    fem = (profil.get('civilite') or '').upper() in ('F', 'MME', 'MADAME')
    npa_loc = str(profil.get('npa_localite') or '').strip()
    m = re.match(r'^(\d{4,5})\s+(.*)$', npa_loc)
    npa, loc = (m.group(1), m.group(2)) if m else ('', npa_loc)
    base, vac, total = _taux(acc.get('taux_total') or 55)
    dc = acc.get('contrat_date') or date.today().isoformat()
    nat = profil.get('nationalite') or ''
    banque = profil.get('banque_nom') or ''
    tit = (profil.get('banque_titulaire') or '').strip()
    nom_complet = f"{f['prenom']} {f['nom']}"
    if tit and tit.lower() != nom_complet.lower():
        banque += f' (titulaire : {tit})'
    return {
        'NOM': f['nom'], 'PRENOM': f['prenom'], 'NOM_COMPLET': nom_complet,
        'CIVILITE': 'Madame' if fem else 'Monsieur', 'E': 'e' if fem else '',
        'COLLAB': 'Collaboratrice' if fem else 'Collaborateur',
        'FONCTION': 'Formatrice auxiliaire' if fem else 'Formateur auxiliaire',
        'DATE_NAISSANCE': _d(profil.get('date_naissance')), 'AVS': profil.get('avs_no'),
        'NATIONALITE': nat, 'LIEU_ORIGINE': profil.get('lieu_origine') or ('—' if nat and 'suisse' not in nat.lower() else ''),
        'ETAT_CIVIL': profil.get('etat_civil'), 'RUE': profil.get('adresse'), 'NPA': npa, 'LOCALITE': loc,
        'TELEPHONE': profil.get('telephone'), 'IBAN': profil.get('iban'), 'BANQUE': banque,
        'DATE_DEBUT': _d(acc.get('date_debut')) or 'dès la signature du présent contrat',
        'TAUX_TOTAL': _chf(total), 'TAUX_BASE': _chf(base), 'TAUX_VACANCES': _chf(vac),
        'TAUX_STANDARD': _chf(total), 'TAUX_MAJORE': _chf(acc.get('taux_majore') or 65),
        'COMMISSION': ('%g' % float(acc.get('commission_pct') or 8)),
        'REGLEMENT_VERSION': reglement_version(),
        'DATE_CONTRAT': _d_long(datetime.strptime(dc[:10], '%Y-%m-%d')),
    }


def _modele(path):
    if not os.path.isfile(path):
        raise D.DossierError(f"Modèle introuvable sur le serveur : {os.path.basename(path)}. "
                             "Vérifie qu'il est bien sur GitHub.", 500)
    return path


def reglement_pdf(nom_complet):
    """Règlement original + nom pré-rempli sur la ligne « Nom et prénom » (la date et la signature restent manuscrites)."""
    reader = PdfReader(_modele(F_REGLEMENT))
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    last = writer.pages[-1]
    pos = []

    def visit(text, cm, tm, fd, fs):
        if 'Nom et pr' in text and not pos:
            pos.append((tm[4], tm[5]))
    try:
        last.extract_text(visitor_text=visit)
    except Exception:
        pass
    if pos:
      try:
        x, y = pos[0]
        w, h = float(last.mediabox.width), float(last.mediabox.height)
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(w, h))
        c.setFont('Helvetica-Bold', 11)
        c.drawString(x + 95, y + 2, nom_complet)
        c.save()
        buf.seek(0)
        last.merge_page(PdfReader(buf).pages[0])
      except Exception as e:
        print('[accueil] nom non ajouté au règlement:', repr(e))
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _pdf(data, filename):
    return Response(data, mimetype='application/pdf', headers={
        'Content-Disposition': f'inline; filename="{filename}"', 'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff'})


def _brevo(api_key, to_email, to_name, subject, html):
    if not api_key or not to_email:
        return False
    try:
        r = requests.post('https://api.brevo.com/v3/smtp/email',
                          headers={'api-key': api_key, 'Content-Type': 'application/json'},
                          json={'sender': {'name': 'Swiss ViTa Form', 'email': 'info@swissvf.ch'},
                                'to': [{'email': to_email, 'name': to_name}], 'subject': subject,
                                'htmlContent': html}, timeout=20)
        print(f'[BREVO accueil] {subject} status={r.status_code}')
        return r.status_code == 201
    except Exception as e:
        print('[BREVO accueil] erreur', repr(e))
        return False


def _mail_cadre(contenu):
    return f"""<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:#c0392b;padding:20px;text-align:center"><h1 style="color:#fff;margin:0;font-size:24px">SWISS ViTa Form</h1></div>
      <div style="padding:28px;background:#f9f9f9;line-height:1.5">{contenu}</div>
      <div style="background:#f0f0f0;padding:14px;text-align:center;font-size:12px;color:#888">Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63</div></div>"""


def _new_code(fid):
    alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
    code = ''.join(secrets.choice(alphabet) for _ in range(10))
    expire = D._now() + timedelta(days=14)
    D.sb('POST', 'formateur_dossier_auth', body={'formateur_id': fid, 'pwd_hash': None,
                                                  'code_hash': generate_password_hash(code),
                                                  'code_expire': expire.isoformat(), 'tentatives': 0,
                                                  'updated_at': D._now().isoformat()},
         prefer='resolution=merge-duplicates,return=minimal')
    return code[:5] + '-' + code[5:], expire.isoformat()


# ---------------- routes ----------------
def register_accueil_routes(app, convert_to_pdf, brevo_key):
    P = '/accueil'

    def guard(role):
        def deco(fn):
            @wraps(fn)
            def w(*a, **k):
                if not D.CONFIG_OK:
                    return D._err('Service dossier non configuré sur le serveur.', 503)
                h = request.headers.get('Authorization', '')
                if not h.startswith('Bearer '):
                    return D._err('Non autorisé.', 401)
                try:
                    c = jwt.decode(h[7:], D.JWT_SECRET, algorithms=['HS256'])
                except Exception:
                    return D._err('Session expirée.', 401)
                if c.get('role') != role:
                    return D._err('Accès refusé.', 403)
                g.claims = c
                try:
                    resp = fn(*a, **k)
                except D.DossierError as e:
                    return D._err(e.msg, e.code)
                except Exception as e:
                    import traceback; traceback.print_exc()
                    return D._err('Erreur serveur (' + type(e).__name__ + ' : ' + str(e)[:160] + ').', 500)
                # expiration glissante : nouveau jeton à chaque appel réussi
                try:
                    resp = app.make_response(resp)
                    resp.headers['X-Token'] = D._token(role, c.get('sub'))
                    resp.headers['Access-Control-Expose-Headers'] = 'X-Token'
                except Exception:
                    pass
                return resp
            return w
        return deco

    def body():
        return request.get_json(silent=True) or {}

    def me():
        fid = g.claims['sub']
        acc = _get_accueil(fid)
        if not acc:
            raise D.DossierError("Aucun parcours d'accueil en cours.", 404)
        return fid, acc

    # ======== instructeur ========
    @app.route(P + '/me', methods=['GET'])
    @guard('formateur')
    def a_me():
        fid, acc = me()
        f = _formateur(fid)
        profil, _ = D._get_profil(fid)
        docs = D._docs(fid)
        return jsonify({'formateur': {k: f[k] for k in ('prenom', 'nom', 'email', 'accueil_statut')},
                        'accueil': {k: acc.get(k) for k in ('date_debut', 'apporteur', 'lectures', 'infos_completes_at',
                                                            'contrat_date', 'termine_at', 'valide_at')},
                        'profil': profil, 'docs': docs, 'etat': _etat(fid, acc, profil, docs),
                        'liens': _docs_partages(), 'qualifs': {k: v[0] for k, v in QUALIFS.items()},
                        'reglement_version': reglement_version()})

    @app.route(P + '/me/lu', methods=['POST'])
    @guard('formateur')
    def a_lu():
        fid, acc = me()
        k = str(body().get('doc') or '')
        if k not in LECTURES:
            raise D.DossierError('Document inconnu.')
        lect = dict(acc.get('lectures') or {})
        lect.setdefault(k, D._now().isoformat())
        _patch_accueil(fid, {'lectures': lect})
        D.log('formateur', fid, fid, 'accueil_lu_' + k)
        return jsonify({'lectures': lect})

    @app.route(P + '/me/infos', methods=['PUT'])
    @guard('formateur')
    def a_infos():
        fid, acc = me()
        b = body()
        q = [x for x in (b.get('qualifications') or []) if x in QUALIFS] if isinstance(b.get('qualifications'), list) else []
        b['qualifications'] = ','.join(q)
        if b.get('taille_tshirt') and b['taille_tshirt'] not in ('XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL'):
            raise D.DossierError('Taille de t-shirt invalide.')
        profil = D._merge_profil(fid, D.clean_profil(b, partial=True))
        D._put_profil(fid, profil, 'formateur')
        e = _etat(fid, {**acc, 'infos_completes_at': 'x'}, profil)
        patch = {'infos_completes_at': D._now().isoformat() if not e['infos_manque'] else None}
        _patch_accueil(fid, patch)
        D.log('formateur', fid, fid, 'accueil_infos')
        return jsonify({'manque': e['infos_manque']})

    def _doc_pdf(fid, acc, quoi, by):
        f = _formateur(fid)
        profil, _ = D._get_profil(fid)
        if by == 'formateur':
            e = _etat(fid, acc, profil)
            if quoi != 'reglement' and not e['infos_ok']:
                raise D.DossierError("Complète d'abord ta fiche d'informations.", 409)
        nom = f"{f['prenom']}_{f['nom']}".replace(' ', '_')
        if quoi == 'reglement':
            D.log(by, by if by == 'admin' else fid, fid, 'accueil_pdf_reglement')
            return _pdf(reglement_pdf(f"{f['prenom']} {f['nom']}"), f'Reglement_instructeurs_{nom}.pdf')
        if quoi == 'avenant' and not acc.get('apporteur'):
            raise D.DossierError("Pas d'avenant prévu.", 404)
        if not acc.get('contrat_date'):
            acc['contrat_date'] = date.today().isoformat()
            _patch_accueil(fid, {'contrat_date': acc['contrat_date']})
        vals = _valeurs(f, profil, acc)
        docx = fill_docx(_modele(F_CONTRAT if quoi == 'contrat' else F_AVENANT), vals)
        D.log(by, by if by == 'admin' else fid, fid, 'accueil_pdf_' + quoi)
        return _pdf(convert_to_pdf(docx), f"{'Contrat_de_travail' if quoi == 'contrat' else 'Avenant_apport_affaires'}_{nom}.pdf")

    @app.route(P + '/me/pdf/<quoi>', methods=['GET'])
    @guard('formateur')
    def a_pdf(quoi):
        if quoi not in SIGNES:
            raise D.DossierError('Document inconnu.', 404)
        fid, acc = me()
        return _doc_pdf(fid, acc, quoi, 'formateur')

    @app.route(P + '/me/signe/<quoi>', methods=['POST'])
    @guard('formateur')
    def a_signe(quoi):
        fid, acc = me()
        if quoi not in SIGNES or (quoi == 'avenant' and not acc.get('apporteur')):
            raise D.DossierError('Document inconnu.', 404)
        if acc.get('termine_at'):
            raise D.DossierError('Documents déjà transmis. Pour une correction, contacte Swiss ViTa Form.', 403)
        if request.content_length and request.content_length > D.MAX_BYTES + 200 * 1024:
            raise D.DossierError('Fichier trop volumineux (10 Mo maximum).', 413)
        lib = {'contrat': 'Contrat signé par l\'instructeur', 'reglement': 'Règlement signé',
               'avenant': 'Avenant signé par l\'instructeur'}[quoi]
        d = D._save_doc(fid, quoi, request.files.get('fichier'), 'formateur', '', lib, allow_admin_types=True)
        D.log('formateur', fid, fid, 'accueil_depot_' + quoi, d['id'])
        return jsonify(d)

    @app.route(P + '/me/terminer', methods=['POST'])
    @guard('formateur')
    def a_terminer():
        fid, acc = me()
        e = _etat(fid, acc)
        if not (e['lectures_ok'] and e['infos_ok'] and e['signes_ok']):
            raise D.DossierError('Il manque encore des étapes.', 409)
        if not acc.get('termine_at'):
            _patch_accueil(fid, {'termine_at': D._now().isoformat()})
            D.sb('PATCH', 'formateurs', {'id': 'eq.' + fid}, {'accueil_statut': 'a_valider'}, prefer='return=minimal')
            f = _formateur(fid)
            _brevo(brevo_key, ADMIN_MAIL, 'Swiss ViTa Form',
                   f"Accueil terminé : {f['prenom']} {f['nom']} — contrat à contresigner",
                   _mail_cadre(f"<p><b>{f['prenom']} {f['nom']}</b> a terminé son parcours d'accueil : documents lus, "
                               f"fiche d'informations complète, contrat et règlement signés déposés.</p>"
                               f"<p>À faire : contresigner le contrat dans le portail (onglet Formateurs → Accueil), puis valider.</p>"))
            D.log('formateur', fid, fid, 'accueil_termine')
        return jsonify({'status': 'ok'})

    # ======== admin ========
    @app.route(P + '/admin/nouveau', methods=['POST'])
    @guard('admin')
    def a_nouveau():
        b = body()
        prenom, nom = str(b.get('prenom') or '').strip(), str(b.get('nom') or '').strip()
        email, login = str(b.get('email') or '').strip(), str(b.get('login') or '').strip()
        mdp = str(b.get('mot_de_passe') or '').strip()
        if not (prenom and nom and email and login and mdp):
            raise D.DossierError('Prénom, nom, e-mail, identifiant et mot de passe obligatoires.')
        if not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email):
            raise D.DossierError('E-mail invalide.')
        if D.sb('GET', 'formateurs', {'login': 'eq.' + login, 'select': 'id'}):
            raise D.DossierError('Cet identifiant existe déjà.', 409)
        dd = str(b.get('date_debut') or '').strip()
        if dd and not D.DATE_RE.fullmatch(dd):
            raise D.DossierError('Date de début invalide.')
        row = D.sb('POST', 'formateurs', body={'prenom': prenom, 'nom': nom, 'email': email, 'login': login,
                                                'mot_de_passe': mdp, 'accueil_statut': 'en_cours'},
                   prefer='return=representation')[0]
        fid = row['id']
        D.sb('POST', 'formateur_accueil', body={
            'formateur_id': fid, 'date_debut': dd or None,
            'taux_total': float(b.get('taux_total') or 55), 'apporteur': bool(b.get('apporteur', True)),
            'taux_majore': float(b.get('taux_majore') or 65), 'commission_pct': float(b.get('commission_pct') or 8)},
             prefer='return=minimal')
        code, expire = _new_code(fid)
        ok = _brevo(brevo_key, email, f'{prenom} {nom}', 'Bienvenue chez Swiss ViTa Form — ton accès au portail',
                    _mail_cadre(f"""<p>Salut {escape(prenom)},</p>
            <p>Bienvenue dans l'équipe Swiss ViTa Form ! Ton accès au portail des instructeurs est prêt.</p>
            <div style="background:#fff;border-radius:8px;padding:16px 20px;margin:18px 0;border-left:4px solid #c0392b">
              <div style="color:#888;font-size:13px">Identifiant</div><div style="font-weight:bold;margin-bottom:8px">{escape(login)}</div>
              <div style="color:#888;font-size:13px">Mot de passe</div><div style="font-weight:bold">{escape(mdp)}</div></div>
            <p>À ta première connexion, un parcours d'accueil te guide en 3 étapes (environ 20 minutes) :</p>
            <ol><li>lire le règlement, la marche à suivre et ce que nous proposons ;</li>
            <li>remplir ta fiche d'informations et joindre tes diplômes ;</li>
            <li>télécharger ton contrat et le règlement, les signer, les dater et nous les renvoyer depuis le portail.</li></ol>
            <p>Tes informations personnelles sont stockées dans un dossier sécurisé. Pour l'ouvrir, je t'envoie
            séparément (SMS ou WhatsApp) un <b>code d'activation</b>.</p>
            <div style="text-align:center;margin:24px 0"><a href="https://portail.swissvf.ch" style="background:#c0392b;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:bold">Accéder au portail</a></div>
            <p>À très vite,<br>Vincent</p>"""))
        D.log('admin', 'admin', fid, 'accueil_creation')
        return jsonify({'formateur': {'id': fid, 'prenom': prenom, 'nom': nom, 'login': login},
                        'code': code, 'code_expire': expire, 'email_envoye': ok})

    @app.route(P + '/admin/liste', methods=['GET'])
    @guard('admin')
    def a_liste():
        accs = D.sb('GET', 'formateur_accueil', {'order': 'created_at.desc'}) or []
        if request.args.get('tous') != '1':
            accs = [a for a in accs if not a.get('valide_at')]
        out = []
        for a in accs:
            fid = a['formateur_id']
            try:
                f = _formateur(fid)
            except D.DossierError:
                continue
            profil, _ = D._get_profil(fid)
            auth = D._auth_row(fid) or {}
            docs = D._docs(fid)
            e = _etat(fid, a, profil, docs)
            depuis = a.get('created_at') or ''
            signes = [{k: d.get(k) for k in ('id', 'type', 'depose_par', 'created_at', 'nom_original')}
                      for d in docs if d['type'] in SIGNES and (d.get('created_at') or '') >= depuis]
            out.append({'formateur': f, 'accueil': a, 'etat': e, 'dossier_active': bool(auth.get('pwd_hash')),
                        'docs_signes': signes,
                        'contresigne': any(d['type'] == 'contrat' and d['depose_par'] == 'admin' for d in signes),
                        'taille_tshirt': (profil.get('taille_tshirt') or '') + (' ' + profil['coupe_tshirt'] if profil.get('coupe_tshirt') else '')})
        return jsonify({'liste': out})

    @app.route(P + '/admin/<fid>', methods=['PATCH'])
    @guard('admin')
    def a_reglages(fid):
        fid = D._uuid(fid)
        b = body()
        patch = {}
        if 'date_debut' in b:
            dd = str(b.get('date_debut') or '')
            if dd and not D.DATE_RE.fullmatch(dd):
                raise D.DossierError('Date de début invalide.')
            patch['date_debut'] = dd or None
        for k in ('taux_total', 'taux_majore', 'commission_pct'):
            if k in b and b[k] not in ('', None):
                patch[k] = float(b[k])
        if 'apporteur' in b:
            patch['apporteur'] = bool(b['apporteur'])
        if b.get('nouvelle_date_contrat'):
            patch['contrat_date'] = None
        if b.get('rouvrir'):
            patch['termine_at'] = None
            D.sb('PATCH', 'formateurs', {'id': 'eq.' + fid}, {'accueil_statut': 'en_cours'}, prefer='return=minimal')
        if patch:
            _patch_accueil(fid, patch)
        D.log('admin', 'admin', fid, 'accueil_reglages')
        return jsonify({'status': 'ok'})

    @app.route(P + '/admin/<fid>/code', methods=['POST'])
    @guard('admin')
    def a_code(fid):
        fid = D._uuid(fid)
        code, expire = _new_code(fid)
        D.log('admin', 'admin', fid, 'code_activation')
        return jsonify({'code': code, 'code_expire': expire})

    @app.route(P + '/admin/<fid>/pdf/<quoi>', methods=['GET'])
    @guard('admin')
    def a_admin_pdf(fid, quoi):
        fid = D._uuid(fid)
        if quoi not in SIGNES:
            raise D.DossierError('Document inconnu.', 404)
        acc = _get_accueil(fid)
        if not acc:
            raise D.DossierError('Pas de parcours pour cet instructeur.', 404)
        return _doc_pdf(fid, acc, quoi, 'admin')

    @app.route(P + '/admin/<fid>/valider', methods=['POST'])
    @guard('admin')
    def a_valider(fid):
        fid = D._uuid(fid)
        acc = _get_accueil(fid)
        if not acc:
            raise D.DossierError('Pas de parcours pour cet instructeur.', 404)
        docs = D._docs(fid)
        depuis = acc.get('created_at') or ''
        if not any(d['type'] == 'contrat' and d.get('depose_par') == 'admin' and (d.get('created_at') or '') >= depuis for d in docs):
            raise D.DossierError("Dépose d'abord le contrat contresigné.", 409)
        # Les fiches de salaire lisent encore ces champs dans la table formateurs : on les recopie.
        profil, _ = D._get_profil(fid)
        D.sb('PATCH', 'formateurs', {'id': 'eq.' + fid}, {
            'accueil_statut': 'termine',
            'civilite': 'F' if (profil.get('civilite') or '').upper() in ('F', 'MME') else 'M',
            'date_naissance': profil.get('date_naissance') or None, 'adresse': profil.get('adresse') or None,
            'npa_localite': profil.get('npa_localite') or None, 'telephone': profil.get('telephone') or None,
            'avs_no': profil.get('avs_no') or None, 'iban': profil.get('iban') or None}, prefer='return=minimal')
        _patch_accueil(fid, {'valide_at': D._now().isoformat()})
        D.log('admin', 'admin', fid, 'accueil_valide')
        return jsonify({'status': 'ok'})
