"""Signature électronique simple des documents d'accueil (contrat, règlement, avenant).

La signature dessinée à l'écran (PNG transparent) est posée sur le PDF généré, à l'endroit
prévu pour la signature (repéré par le texte du document : « Lieu et date », lignes de points…).
Si les repères sont introuvables (modèle modifié), une page de signatures est ajoutée à la fin.
"""
import base64
import io
import re

from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

ENCRE = (0.10, 0.17, 0.43)          # bleu encre
GRIS = (0.45, 0.45, 0.45)
MAX_PNG = 400 * 1024


class SignatureError(ValueError):
    pass


def png_depuis_dataurl(data_url):
    """Valide et décode la signature envoyée par le navigateur (data:image/png;base64,...)."""
    s = str(data_url or '')
    m = re.match(r'^data:image/png;base64,([A-Za-z0-9+/=\s]+)$', s)
    if not m:
        raise SignatureError('Signature manquante ou invalide.')
    try:
        data = base64.b64decode(m.group(1), validate=False)
    except Exception:
        raise SignatureError('Signature illisible.')
    if not data.startswith(b'\x89PNG\r\n\x1a\n') or len(data) > MAX_PNG:
        raise SignatureError('Signature invalide.')
    w, h = ImageReader(io.BytesIO(data)).getSize()
    if w < 40 or h < 15:
        raise SignatureError('Signature trop petite : signe un peu plus grand.')
    return data


def _runs(page):
    """Morceaux de texte d'une page avec leur position (x, y)."""
    out = []

    def visit(text, cm, tm, fd, fs):
        t = (text or '').strip()
        if not t:
            return
        x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
        y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
        out.append((x, y, t))
    try:
        page.extract_text(visitor_text=visit)
    except Exception:
        pass
    return out


def _cherche(runs, motif):
    rx = re.compile(motif, re.I)
    return [r for r in runs if rx.search(r[2])]


def _dernier_page_avec(reader, motif):
    for i in range(len(reader.pages) - 1, -1, -1):
        runs = _runs(reader.pages[i])
        if _cherche(runs, motif):
            return i, runs
    return None, []


# ---------- repères par document ----------
def reperes(quoi, pdf_bytes):
    """Renvoie {'page': i, 'instructeur': (x, y, w, h), 'employeur': (x, y, w, h) | None,
    'lieu_date': (x, y) | None} ou None si les repères ne sont pas trouvés."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if quoi == 'contrat':
        i, runs = _dernier_page_avec(reader, r'Lieu et date')
        if i is None:
            return None
        lieu = _cherche(runs, r'Lieu et date')[0]
        pts = sorted([r for r in runs if '……' in r[2] and r[1] > lieu[1]], key=lambda r: r[1])
        if not pts:
            return None
        y_ligne = pts[0][1]
        ligne = [r for r in pts if abs(r[1] - y_ligne) < 3]
        gauche = min(ligne, key=lambda r: r[0])
        droite = max(ligne, key=lambda r: r[0]) if len(ligne) > 1 else None
        return {'page': i,
                'instructeur': (gauche[0], y_ligne + 4, 135, 34),
                'employeur': (droite[0], y_ligne + 4, 150, 34) if droite else None,
                'legende_instr': (lieu[0], lieu[1] - 12),
                'legende_empl': (droite[0], y_ligne - 9) if droite else None,
                'lieu_date': (lieu[0] + 64, lieu[1] + 3)}
    if quoi == 'avenant':
        i, runs = _dernier_page_avec(reader, r"Pour SWISS ViTa Form")
        if i is None:
            return None
        svf = _cherche(runs, r'Pour SWISS ViTa Form')[0]
        ins = _cherche(runs, r"L.Instructeur / l.Instructrice\s*:")
        ins = ins[-1] if ins else None
        lignes = [r for r in runs if '____' in r[2] and r[1] < svf[1]]
        if not ins or not lignes:
            return None
        y_ligne = max(r[1] for r in lignes)
        lieu = _cherche(runs, r'Lieu et date')
        return {'page': i,
                'instructeur': (ins[0], y_ligne + 4, 150, 40),
                'employeur': (svf[0], y_ligne + 4, 150, 40),
                'legende_instr': (ins[0], y_ligne - 9),
                'legende_empl': (svf[0], y_ligne - 9),
                'lieu_date': (lieu[-1][0] + 62, lieu[-1][1] + 2) if lieu else None}
    if quoi == 'reglement':
        i, runs = _dernier_page_avec(reader, r'Lu et accept')
        if i is None:
            return None
        lu = _cherche(runs, r'Lu et accept')[0]
        sig = [r for r in _cherche(runs, r'^Signature') if r[1] < lu[1]]
        if not sig:
            return None
        sig = sig[0]
        return {'page': i,
                'instructeur': (sig[0] + 58, sig[1] + 1, 150, 22),
                'employeur': None,
                'legende_instr': (sig[0] + 58, sig[1] - 10),
                'legende_empl': None,
                'lieu_date': (lu[0] + 82, lu[1] + 2)}
    return None


# ---------- dessin ----------
def _image(c, png, box):
    x, y, bw, bh = box
    img = ImageReader(io.BytesIO(png))
    iw, ih = img.getSize()
    k = min(bw / iw, bh / ih)
    w, h = iw * k, ih * k
    c.drawImage(img, x, y, width=w, height=h, mask='auto')


def _texte(c, xy, txt, size=10, color=ENCRE, font='Helvetica'):
    c.setFillColorRGB(*color)
    c.setFont(font, size)
    c.drawString(xy[0], xy[1], txt)


def _fusionner(pdf_bytes, page_index, dessin):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    page = writer.pages[page_index]
    w, h = float(page.mediabox.width), float(page.mediabox.height)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    dessin(c)
    c.save()
    buf.seek(0)
    page.merge_page(PdfReader(buf).pages[0])
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _page_signatures(pdf_bytes, titre, blocs):
    """Repli : ajoute une page « Signatures » à la fin du document."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    w, h = float(reader.pages[0].mediabox.width), float(reader.pages[0].mediabox.height)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    _texte(c, (60, h - 80), 'Signatures — ' + titre, 14, (0, 0, 0), 'Helvetica-Bold')
    y = h - 130
    for b in blocs:
        _texte(c, (60, y), b['qui'], 11, (0, 0, 0), 'Helvetica-Bold')
        _image(c, b['png'], (60, y - 70, 200, 60))
        _texte(c, (60, y - 84), b['legende'], 8, GRIS)
        y -= 140
    c.save()
    buf.seek(0)
    writer.add_page(PdfReader(buf).pages[0])
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def signer_instructeur(quoi, pdf_bytes, png, nom, lieu, quand, ip):
    """Pose la signature de l'instructeur·trice, le lieu et la date."""
    date_txt = quand.strftime('%d.%m.%Y')
    legende = f"Signé électroniquement par {nom} le {date_txt} à {quand.strftime('%H:%M')}" + (f' (IP {ip})' if ip else '')
    r = reperes(quoi, pdf_bytes)
    if not r:
        return _page_signatures(pdf_bytes, quoi.capitalize(), [{'qui': nom, 'png': png, 'legende': f'{lieu}, le {date_txt} — {legende}'}])

    def dessin(c):
        _image(c, png, r['instructeur'])
        if r.get('lieu_date'):
            _texte(c, r['lieu_date'], f'{lieu}, le {date_txt}' if quoi != 'reglement' else f'{date_txt} à {lieu}', 10)
        _texte(c, r['legende_instr'], legende, 5.5, GRIS)
    return _fusionner(pdf_bytes, r['page'], dessin)


def signer_employeur(quoi, pdf_bytes, png, nom, quand):
    """Contresignature de Swiss ViTa Form sur le document déjà signé par l'instructeur·trice."""
    legende = f"Contresigné électroniquement par {nom} le {quand.strftime('%d.%m.%Y')} à {quand.strftime('%H:%M')}"
    r = reperes(quoi, pdf_bytes)
    if not r or not r.get('employeur'):
        return _page_signatures(pdf_bytes, quoi.capitalize(), [{'qui': nom + ' — Swiss ViTa Form', 'png': png, 'legende': legende}])

    def dessin(c):
        _image(c, png, r['employeur'])
        _texte(c, r['legende_empl'], legende, 5.5, GRIS)
    return _fusionner(pdf_bytes, r['page'], dessin)
