"""Attestation de participation (cours non certifiés SRC : pédiatrie, sur mesure, etc.).

Document distinct du certificat BLS-AED-SRC : il confirme seulement la participation,
sans validité de 2 ans ni reconnaissance du Swiss Resuscitation Council.
"""
import io
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas

from reportlab.lib import colors

from documents_pdf import FONT, FONT_BOLD, FONT_TITLE


def _date_fr(v):
    try:
        return datetime.strptime(str(v)[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
    except Exception:
        return str(v or '')


def _duree(debut, fin):
    """Durée en heures, pause de midi (12h-13h) déduite si le cours l'enjambe."""
    try:
        h1, m1 = [int(x) for x in str(debut).split(':')[:2]]
        h2, m2 = [int(x) for x in str(fin).split(':')[:2]]
    except Exception:
        return None
    s, e = h1 + m1 / 60, h2 + m2 / 60
    d = e - s - (1 if s < 12 < e else 0)
    if d <= 0:
        return None
    return int(d) if d == int(d) else round(d, 1)


# Palette du certificat SRC, pour que les deux documents soient de la même famille
ROUGE = colors.HexColor('#B8050F')
TEXTE = colors.HexColor('#3A2C2C')
GRIS = colors.HexColor('#8A7A7A')
TRAIT = colors.HexColor('#D9C3BC')
CADRE = colors.HexColor('#C0392B')
RESPONSABLE = 'Vincent Detta'
PIED = ('SWISS ViTa Form Vincent Detta  ·  Av. Kiener 29  ·  1400 Yverdon-les-Bains  ·  '
        '078 892 02 63  ·  info@swissvf.ch  ·  www.swissvf.ch')


def _espace(c, text, x, y, font, size, color, espacement, centre=True):
    """Texte à lettres espacées (comme « C E R T I F I C A T »)."""
    c.saveState()  # l'espacement des lettres ne doit pas déborder sur le texte suivant
    c.setFont(font, size)
    c.setFillColor(color)
    largeur = c.stringWidth(text, font, size) + espacement * (len(text) - 1)
    t = c.beginText()
    t.setTextOrigin(x - largeur / 2 if centre else x, y)
    t.setFont(font, size)
    t.setCharSpace(espacement)
    t.textOut(text)
    c.drawText(t)
    c.restoreState()


def _centre(c, text, y, font, size, color, gras=False, penche=False):
    w, _ = A4
    c.setFillColor(color)
    c.setFont(font, size)
    if not (gras or penche):
        c.drawCentredString(w / 2, y, text)
        return
    _signature(c, text, w / 2, y, font, size, color, gras, penche)


def _signature(c, text, cx, y, font, size, color, gras=False, penche=False):
    """Lora n'existe ici qu'en version régulière : gras et italique sont simulés."""
    c.saveState()
    c.setFillColor(color)
    c.setStrokeColor(color)
    largeur = c.stringWidth(text, font, size)
    c.translate(cx, y)
    if penche:
        c.skew(0, 12)
    t = c.beginText()
    t.setTextOrigin(-largeur / 2, 0)
    t.setFont(font, size)
    if gras:
        c.setLineWidth(size * 0.035)
        t.setTextRenderMode(2)
    t.textOut(text)
    c.drawText(t)
    c.restoreState()


def build_attestation_pdf(data, logo_bytes=None, feminin=False):
    w, h = A4
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle('Attestation de participation')
    top = lambda v: h - v  # positions mesurées depuis le haut, comme sur le certificat

    c.setStrokeColor(CADRE)
    c.setLineWidth(0.6)
    c.rect(24, 24, w - 48, h - 48)

    if logo_bytes:
        try:
            lw, lh = 150, 76
            c.drawImage(ImageReader(io.BytesIO(logo_bytes)), (w - lw) / 2, top(140), width=lw, height=lh,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    _espace(c, 'SWISS VITA FORM', w / 2, top(160), FONT, 8, ROUGE, 2.6)
    _espace(c, 'ATTESTATION', w / 2, top(199), FONT_TITLE, 26, TEXTE, 7)

    _centre(c, 'atteste que', top(240), FONT, 10, TEXTE)
    _centre(c, 'Madame' if feminin else 'Monsieur', top(261), FONT, 10, TEXTE)
    _centre(c, f"{data.get('prenom', '')} {data.get('nom', '')}".strip(), top(290), FONT_TITLE, 20, ROUGE, gras=True)
    c.setStrokeColor(TRAIT)
    c.setLineWidth(0.7)
    c.line(73, top(301), w - 73, top(301))

    _centre(c, 'A participé au cours', top(345), FONT, 11, TEXTE)
    cours = data.get('cours', '')
    lignes = simpleSplit(cours, FONT_BOLD, 11.5, w - 160)
    y = 366
    for ligne in lignes:
        _centre(c, ligne, top(y), FONT_BOLD, 11.5, TEXTE)
        y += 16
    _centre(c, f"le {_date_fr(data.get('date_cours'))}", top(y + 4), FONT, 11, TEXTE)
    y += 33
    duree = _duree(data.get('heure_debut'), data.get('heure_fin'))
    if duree:
        _espace(c, f"DURÉE : {duree} HEURE{'S' if duree > 1 else ''}", w / 2, top(y), FONT, 7.5, GRIS, 0.6)
    y += 10

    # Thèmes abordés (facultatif)
    themes = data.get('themes') or ''
    if isinstance(themes, str):
        sep = '\n' if '\n' in themes else (';' if ';' in themes else ',')
        themes = [t.strip(' -•\t') for t in themes.split(sep)]
    themes = [t for t in themes if t][:8]
    if themes:
        y += 26
        _espace(c, 'THÈMES ABORDÉS', w / 2, top(y), FONT, 7.5, ROUGE, 1.6)
        y += 18
        bloc = 300
        x0 = (w - bloc) / 2
        for t in themes:
            for j, ligne in enumerate(simpleSplit(t, FONT, 10, bloc - 14)):
                if j == 0:
                    c.setFillColor(ROUGE)
                    c.circle(x0 + 3, top(y) + 3.2, 1.4, stroke=0, fill=1)
                c.setFillColor(TEXTE)
                c.setFont(FONT, 10)
                c.drawString(x0 + 12, top(y), ligne)
                y += 14
        y += 6

    # Signatures, sur deux colonnes comme le certificat
    signataires = [(RESPONSABLE, 'Responsable de formation')]
    if data.get('formateur') and data['formateur'] != RESPONSABLE:
        signataires.append((data['formateur'], 'Direction du cours'))
    for nom in (data.get('instructeurs') or [])[:4]:
        if nom and nom not in (data.get('formateur'), RESPONSABLE):
            signataires.append((nom, 'Instructeur-trice'))
    ys = max(477, y + 40)
    colonnes = (160, 415)
    for k, (nom, role) in enumerate(signataires):
        cx = colonnes[k % 2]
        yy = ys + (k // 2) * 58
        _signature(c, nom, cx, top(yy), FONT_TITLE, 13, ROUGE, penche=True)
        c.setStrokeColor(TRAIT)
        c.setLineWidth(0.7)
        c.line(cx - 87, top(yy + 11), cx + 87, top(yy + 11))
        c.setFillColor(GRIS)
        c.setFont(FONT, 7.5)
        c.drawCentredString(cx, top(yy + 24), role)

    c.setFillColor(GRIS)
    c.setFont(FONT, 6.8)
    pied = simpleSplit(PIED, FONT, 6.8, w - 120)
    for i, ligne in enumerate(pied):
        c.drawCentredString(w / 2, top(796 + i * 11), ligne)

    c.showPage()
    c.save()
    return buf.getvalue()
