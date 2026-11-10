# Enquête MSF GROUP

Site statique (aucun build).

| Fichier | Rôle |
|---|---|
| `index.html` | Formulaire (21 questions) |
| `index.full.html` | Version longue (37 questions) |
| `admin.html` | Stats locales (réponses du navigateur courant) |
| `apps-script.gs` | Google Sheets + e-mail de remerciement avec lien du guide |

## Recevoir les réponses dans Excel + envoyer le guide par e-mail
1. Créer un Google Sheet vide > Extensions > Apps Script > coller `apps-script.gs`.
2. Le lien du guide (Google Drive) est déjà renseigné dans `GUIDE_URL`. Renseigner si voulu `ADMIN_EMAIL`. Vérifier que le fichier Drive est partagé en « Toute personne disposant du lien ».
3. Déployer > Nouveau déploiement > Application Web (Exécuter en tant que : moi / Accès : tout le monde). Autoriser Sheets + envoi d'e-mails.
4. Copier l'URL `/exec` dans `const ENDPOINT = "";` de `index.html`.
5. Mise à jour du script : Déployer > Gérer les déploiements > Modifier > Nouvelle version.
6. Export Excel : Fichier > Télécharger > Microsoft Excel (.xlsx).

Limite e-mails : ~100/jour (Gmail gratuit), ~1500 (Workspace).

## Déploiement
- GitHub : pousser ce dossier.
- Vercel : importer le dépôt, Framework = Other, Root Directory = `language-survey`, pas de build.

