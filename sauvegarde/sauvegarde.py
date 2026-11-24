"""Sauvegarde complète SWISS ViTa Form (lancée chaque nuit par GitHub Actions).

Produit dans le dossier donné en argument :
  - donnees_SVF.xlsx  : une feuille par table (lisible dans Excel)
  - csv/<table>.csv   : les mêmes données en CSV
  - fichiers/<bucket>/...  : tous les fichiers du stockage Supabase
    (fiches de salaire, contrats, documents, certificats…)
Le fichier base.dump (restauration complète) est fait par pg_dump dans le workflow.

Les journaux GitHub peuvent être publics : on n'affiche jamais de noms de
fichiers ni de contenu, uniquement des nombres.
"""
import csv
import datetime as dt
import decimal
import json
import os
import sys
import time
import uuid
from pathlib import Path
from urllib.parse import quote

import psycopg
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

DB_URL = os.environ['SUPABASE_DB_URL']
SB_URL = os.environ['SUPABASE_URL'].rstrip('/')
SB_KEY = os.environ['SUPABASE_SERVICE_KEY']
ENC_KEY = os.environ.get('DOSSIER_ENC_KEY', '').strip()
H = {'apikey': SB_KEY, 'Authorization': f'Bearer {SB_KEY}'}

erreurs = []


def texte(v):
    """Valeur lisible pour Excel / CSV."""
    if v is None:
        return None
    if isinstance(v, (dt.datetime, dt.time)):
        return v.replace(tzinfo=None).isoformat(sep=' ') if isinstance(v, dt.datetime) else v.isoformat()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, decimal.Decimal):
        return float(v)
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False, default=str)
    if isinstance(v, (bytes, memoryview)):
        return '<binaire>'
    if isinstance(v, uuid.UUID):
        return str(v)
    if isinstance(v, str) and len(v) > 32000:  # limite d'une cellule Excel
        return v[:32000] + '…'
    return v


def dechiffreur():
    if not ENC_KEY:
        return None
    try:
        from cryptography.fernet import Fernet
        f = Fernet(ENC_KEY.encode())
    except Exception:
        erreurs.append('DOSSIER_ENC_KEY invalide : dossiers laissés chiffrés')
        return None

    def dec(txt):
        try:
            return f.decrypt(txt.encode()).decode()
        except Exception:
            return '<illisible>'
    return dec


def exporter_tables(dest: Path):
    dec = dechiffreur()
    (dest / 'csv').mkdir(parents=True, exist_ok=True)
    wb = Workbook(write_only=True)
    noms = set()
    total = 0
    with psycopg.connect(DB_URL, connect_timeout=30) as cx, cx.cursor() as cur:
        cur.execute("select tablename from pg_tables where schemaname='public' order by tablename")
        tables = [r[0] for r in cur.fetchall()]
        for t in tables:
            cur.execute(f'select * from public."{t}"')
            cols = [d.name for d in cur.description]
            rows = cur.fetchall()
            enc_idx = [i for i, c in enumerate(cols) if c.endswith('_enc')] if dec else []
            entetes = cols + [cols[i][:-4] + '_dechiffre' for i in enc_idx]
            lignes = []
            for r in rows:
                l = [texte(v) for v in r]
                l += [dec(r[i]) if r[i] else None for i in enc_idx]
                lignes.append(l)
            total += len(lignes)
            with open(dest / 'csv' / f'{t}.csv', 'w', newline='', encoding='utf-8-sig') as fh:
                w = csv.writer(fh, delimiter=';')
                w.writerow(entetes)
                w.writerows(lignes)
            nom = t[:31]
            n = 1
            while nom in noms:
                n += 1
                nom = f'{t[:28]}_{n}'
            noms.add(nom)
            ws = wb.create_sheet(nom)
            for i, c in enumerate(entetes, 1):
                ws.column_dimensions[get_column_letter(i)].width = min(max(len(c) + 2, 12), 40)
            ws.append(entetes)
            for l in lignes:
                ws.append(l)
    wb.save(dest / 'donnees_SVF.xlsx')
    print(f'Tables : {len(tables)} — lignes : {total}')
    return len(tables)


def lister(bucket, prefixe=''):
    """Tous les fichiers d'un bucket (récursif)."""
    out, offset = [], 0
    while True:
        r = requests.post(f'{SB_URL}/storage/v1/object/list/{bucket}', headers=H, timeout=60,
                          json={'prefix': prefixe, 'limit': 1000, 'offset': offset,
                                'sortBy': {'column': 'name', 'order': 'asc'}})
        r.raise_for_status()
        items = r.json()
        for it in items:
            chemin = f"{prefixe}{it['name']}"
            if it.get('id') is None:          # sous-dossier
                out += lister(bucket, chemin + '/')
            else:
                out.append(chemin)
        if len(items) < 1000:
            return out
        offset += 1000


def exporter_fichiers(dest: Path):
    r = requests.get(f'{SB_URL}/storage/v1/bucket', headers=H, timeout=60)
    r.raise_for_status()
    buckets = [b['id'] for b in r.json()]
    n = taille = 0
    for b in buckets:
        for chemin in lister(b):
            cible = dest / 'fichiers' / b / chemin
            cible.parent.mkdir(parents=True, exist_ok=True)
            ok = False
            for essai in range(4):
                try:
                    f = requests.get(f'{SB_URL}/storage/v1/object/{b}/{quote(chemin)}',
                                     headers=H, timeout=300)
                    if f.status_code == 200:
                        cible.write_bytes(f.content)
                        n += 1
                        taille += len(f.content)
                        ok = True
                        break
                    if f.status_code < 500:
                        break
                except requests.RequestException:
                    pass
                time.sleep(3 * (essai + 1))
            if not ok:
                erreurs.append(f'fichier non téléchargé dans le bucket {b}')
    print(f'Buckets : {len(buckets)} — fichiers : {n} — {taille / 1e6:.1f} Mo')


if __name__ == '__main__':
    dest = Path(sys.argv[1])
    dest.mkdir(parents=True, exist_ok=True)
    exporter_tables(dest)
    exporter_fichiers(dest)
    (dest / 'LISEZMOI.txt').write_text(
        f"Sauvegarde SWISS ViTa Form du {dt.datetime.now():%d.%m.%Y %H:%M} (UTC)\n\n"
        "donnees_SVF.xlsx  -> toutes les données, une feuille par table (ouvrir avec Excel)\n"
        "csv/              -> les mêmes données en CSV\n"
        "fichiers/         -> tous les fichiers du portail (fiches de salaire, contrats, documents...)\n"
        "base.dump         -> copie complète de la base pour une restauration (voir RESTAURATION.md sur GitHub)\n"
        + ('\nAvertissements :\n- ' + '\n- '.join(sorted(set(erreurs))) + '\n' if erreurs else ''),
        encoding='utf-8')
    if erreurs:
        print('AVERTISSEMENTS :', len(erreurs))
        for e in sorted(set(erreurs)):
            print(' -', e)
        # la sauvegarde part quand même ; le workflow échouera à la fin pour t'avertir
        Path(os.environ.get('GITHUB_WORKSPACE', '.'), 'AVERTISSEMENTS').write_text('\n'.join(erreurs))
