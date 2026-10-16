"""E-mail de fiche de salaire : remerciement différent à chaque envoi + petite phrase du mois."""
import zlib
from datetime import datetime
from zoneinfo import ZoneInfo

MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']

# {p} = prénom, {m} = mois de la fiche. Formulations neutres (valables pour formateurs et formatrices), tutoiement.
REMERCIEMENTS = [
    "Merci du fond du cœur pour ton travail en {m}. Chaque cours que tu donnes, c'est un peu plus de monde prêt à agir le jour où ça compte vraiment.",
    "Un grand merci pour ton engagement en {m}. Ton énergie et ton sérieux font la différence, et les participants le sentent.",
    "Merci pour tout ce que tu as apporté à Swiss ViTa Form en {m}. C'est un vrai plaisir de pouvoir compter sur toi.",
    "Bravo et merci pour {m} ! Ta façon de transmettre ces gestes qui sauvent est précieuse, et elle ne passe pas inaperçue.",
    "Merci d'avoir donné de ton temps et de ton énergie en {m}. Grâce à toi, des gens repartent avec la confiance d'oser agir.",
    "Un immense merci pour ton travail en {m}. Savoir que les formations sont entre de si bonnes mains, c'est un vrai confort.",
    "Merci pour ta disponibilité et ta bonne humeur en {m}. Elles comptent autant que le contenu des cours.",
    "Merci, tout simplement, pour {m}. Ce que tu fais sur le terrain avec les participants, c'est le cœur de Swiss ViTa Form.",
    "Merci pour ton professionnalisme en {m}. Former, rassurer, motiver : tu fais tout cela à la fois, et très bien.",
    "Un grand merci pour ton implication en {m}. Les retours des participants parlent d'eux-mêmes, continue comme ça !",
    "Merci pour ta belle énergie en {m}. Tu contribues directement à rendre notre région un peu plus sûre.",
    "Merci d'avoir fait vivre les cours en {m}. Ce n'est pas toujours simple, et tu le fais avec brio.",
    "Merci pour ta fiabilité et ton enthousiasme en {m}. Travailler avec toi, c'est un plaisir et une chance.",
    "Merci pour le temps que tu consacres à transmettre ces gestes essentiels. {m} n'a pas fait exception, et c'était top !",
]

PIECE_JOINTE = [
    "Tu trouveras ta fiche de salaire de {m} {a} en pièce jointe.",
    "Ta fiche de salaire de {m} {a} est jointe à ce message.",
    "Voici, en pièce jointe, ta fiche de salaire pour {m} {a}.",
    "Ci-joint ta fiche de salaire de {m} {a}.",
]

FIN = [
    "Une question ou une remarque sur la fiche ? Réponds simplement à cet e-mail.",
    "Si quelque chose te semble bizarre sur la fiche, écris-moi, on regarde ça ensemble.",
    "Un doute sur un montant ? N'hésite pas à me répondre directement.",
    "Elle est aussi disponible à tout moment dans ton espace du portail.",
]

# Une petite phrase par mois (sans attribution : phrases d'inspiration maison).
PHRASES = {
    1: "Une nouvelle année, c'est un carnet tout blanc : chaque geste utile en écrit déjà la première page.",
    2: "Au cœur de l'hiver, la chaleur humaine reste le meilleur des remèdes.",
    3: "Le printemps arrive : tout ce qu'on sème aujourd'hui sauvera peut-être quelqu'un demain.",
    4: "Un petit geste appris peut tout changer. À chaque cours, un cœur de plus saura quoi faire.",
    5: "Mai, le mois où tout renaît : et si on donnait à quelqu'un l'envie d'apprendre à sauver ?",
    6: "Les journées sont longues, les sourires aussi. Profite de la lumière de juin.",
    7: "Le repos fait partie du travail bien fait. Profite de l'été, tu l'as mérité.",
    8: "Même au soleil, un bon formateur ou une bonne formatrice garde l'œil ouvert… et la trousse de secours à portée de main.",
    9: "La rentrée, c'est l'occasion de repartir avec l'énergie des tout premiers jours.",
    10: "En octobre, les arbres nous montrent qu'on peut tout lâcher et rester magnifique. Prends soin de toi.",
    11: "Quand les jours raccourcissent, ce sont les petites lumières qui comptent. La tienne en est une.",
    12: "Donner de son temps est le plus beau des cadeaux. Merci d'écrire cette aventure avec nous, et de très belles fêtes !",
}


def _mois_num(mois_label):
    try:
        return MOIS.index((mois_label or '').strip().lower()) + 1
    except ValueError:
        return 0


def _idx(seed, n, step):
    return (zlib.crc32(seed.encode('utf-8')) + step) % n


def prenom_de(formateur_nom):
    parts = (formateur_nom or '').strip().split()
    return parts[0] if parts else ''


def build_fiche_salaire_email(formateur_nom, mois_label, annee, email='', now=None):
    """Retourne (sujet, html). Le texte varie d'une personne à l'autre ET d'un mois à l'autre
    (un même destinataire ne reçoit jamais deux mois de suite le même remerciement)."""
    now = now or datetime.now(ZoneInfo('Europe/Zurich'))
    prenom = prenom_de(formateur_nom)
    m = (mois_label or '').strip().lower()
    try:
        k = int(annee) * 12 + _mois_num(m)
    except (TypeError, ValueError):
        k = now.year * 12 + now.month
    seed = (email or formateur_nom or 'x').lower()
    merci = REMERCIEMENTS[_idx(seed, len(REMERCIEMENTS), k)].format(p=prenom, m=m)
    pj = PIECE_JOINTE[_idx(seed + 'pj', len(PIECE_JOINTE), k)].format(m=m, a=annee)
    fin = FIN[_idx(seed + 'fin', len(FIN), k)]
    phrase = PHRASES[now.month]
    mois_phrase = MOIS[now.month - 1]
    de_mois = ("d'" if mois_phrase[0] in 'aeiouéh' else 'de ') + mois_phrase
    salut = f"Bonjour {prenom}," if prenom else "Bonjour,"
    html = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #333;">
        <div style="background: #c0392b; padding: 20px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">SWISS ViTa Form</h1>
            <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0 0;">Fiche de salaire — {m} {annee}</p>
        </div>
        <div style="padding: 30px; background: #f9f9f9; line-height: 1.55;">
            <p>{salut}</p>
            <p>{merci}</p>
            <p>{pj}<br>{fin}</p>
            <div style="margin: 26px 0; padding: 16px 18px; background: #fff; border-left: 4px solid #c0392b; border-radius: 6px;">
                <div style="font-size: 12px; color: #c0392b; font-weight: bold; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 6px;">✨ La petite phrase {de_mois}</div>
                <div style="font-style: italic; font-size: 15px;">{phrase}</div>
            </div>
            <p>À très bientôt,<br><strong>Vincent</strong><br><span style="color:#888;font-size:13px;">Swiss ViTa Form</span></p>
            <div style="text-align: center; margin: 24px 0 4px;">
                <a href="https://portail.swissvf.ch" style="background: #c0392b; color: white; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: bold;">Accéder au portail</a>
            </div>
        </div>
        <div style="background: #f0f0f0; padding: 16px; text-align: center; font-size: 12px; color: #888;">
            Swiss ViTa Form — Av. Kiener 29, 1400 Yverdon-les-Bains — 078 892 02 63
        </div>
    </div>
    """
    subject = f"Ta fiche de salaire de {m} {annee} 💛"
    return subject, html
