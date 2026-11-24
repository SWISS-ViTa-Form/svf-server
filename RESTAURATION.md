# Sauvegardes SWISS ViTa Form

Chaque nuit, GitHub (`.github/workflows/sauvegarde.yml`) envoie une archive chiffrée sur Google Drive, dans `SVF-Sauvegardes/` :
- `quotidien/` : les 30 derniers jours
- `mensuel/` : celle du 1er de chaque mois, gardée 12 mois

## Ouvrir une sauvegarde
1. Télécharger le fichier `SVF_sauvegarde_AAAA-MM-JJ.7z` depuis Google Drive.
2. L'ouvrir avec **7-Zip** (gratuit, 7-zip.org) et saisir le mot de passe de sauvegarde.
3. Contenu :
   - `donnees_SVF.xlsx` : toutes les données, une feuille par table (clients, cours, participants, devis/factures, compta, formateurs…). La colonne `profil_dechiffre` contient les dossiers des instructeurs en clair (AVS, IBAN…).
   - `fichiers/` : tous les fichiers du portail (fiches de salaire, contrats, documents…)
   - `base.dump` + `structure.sql` : copie complète de la base, pour une restauration

## Restaurer la base dans Supabase (en cas de catastrophe)
À faire avec l'aide d'une personne technique (ou de Claude) :
1. Créer un nouveau projet Supabase (ou vider l'ancien).
2. `pg_restore --no-owner --no-privileges -d "<connexion Session pooler du projet>" base.dump`
3. Recréer les buckets du stockage avec les mêmes noms et y remettre le contenu de `fichiers/<bucket>/`.
4. Mettre la nouvelle URL et les nouvelles clés dans Render et dans `index.html`.

## Ce qu'il faut garder ailleurs que sur l'ordinateur (gestionnaire de mots de passe + papier)
- le mot de passe de sauvegarde (`BACKUP_PASSWORD`) : **sans lui, les archives sont illisibles**
- `DOSSIER_ENC_KEY` (Render) : la clé qui déchiffre les dossiers des instructeurs dans la base
