# Chatbot Immigration — Assistant interne pour la plateforme d'administration

Assistant conversationnel (agent LLM à function-calling) destiné aux administrateurs d'une plateforme
de gestion de candidatures à l'immigration/mobilité internationale (études, travail, stages). Il répond
aux questions sur les candidats, demandes, entretiens et commandes, et peut exécuter des actions de
création/modification — **toujours ancré dans les données réelles de MongoDB**, jamais sur ce que le
modèle "croit savoir".

> 🔗 **Backend :** [jaz56/admin-backend](https://github.com/jaz56/admin-backend) · **Frontend :** [jaz56/admin-front](https://github.com/jaz56/admin-front)

## Principe de fonctionnement

Le LLM (modèle `gemma4:31b-cloud` via [Ollama Cloud](https://ollama.com)) ne répond jamais directement
à une question factuelle : il doit appeler une fonction Python qui va chercher la donnée en base
MongoDB, puis formuler sa réponse à partir du résultat exact de cette fonction. Le prompt système
interdit explicitement d'inventer un chiffre, un statut ou une date.

Chaîne de traitement d'un message :

```
Front-end ──POST /api/v1/chat──▶ main.py ──▶ auth.py (vérifie le JWT, charge les permissions)
                                     │
                                     ▼
                                  chat.py (agent : prompt système + ~60 outils function-calling)
                                     │
                                     ▼
                                   db.py (accès MongoDB : la seule source de vérité)
```

- **[main.py](main.py)** — API FastAPI (3 routes : chat, reset de session, téléchargement d'export).
- **[auth.py](auth.py)** — vérifie le JWT (émis par le backend Spring Boot existant) et charge les
  permissions de l'admin depuis MongoDB.
- **[chat.py](chat.py)** — le cœur de l'agent : prompt système, liste des outils exposés au modèle,
  boucle de function-calling, contrôle des permissions par outil, flux de confirmation pour les actions
  d'écriture, export Excel, graphiques, anti-hallucination.
- **[db.py](db.py)** — toutes les fonctions d'accès aux données (comptages, répartitions, profils
  complets, graphiques, créations/modifications, recherche floue de candidat par nom).

## Fonctionnalités

- **Lecture** : comptages et répartitions (par statut, progression, pays, rôle...), profils complets
  (utilisateur, demande, entretien, commande), historique complet d'un candidat (demandes + entretiens +
  commandes rattachées), documents (CIN, photo, enregistrements), transactions, abonnements, statut STB.
- **Graphiques** : données prêtes à l'emploi (barres/camembert/courbe) pour le front-end — répartitions
  et tendances dans le temps.
- **Écriture** (avec confirmation "oui/non" obligatoire avant exécution) : création de candidat/demande/
  entretien, changement de statut ou de progression d'une demande (relayé au vrai backend Spring Boot
  pour garder la logique métier), changement de statut d'un entretien, blocage/déblocage de compte,
  gestion des permissions d'un admin, réinitialisation de mot de passe.
- **Export** : filtrage de demandes puis génération d'un fichier Excel avec sélection des colonnes.
- **Fallback générique** (réservé au super admin) : si aucune fonction dédiée ne couvre la question,
  l'agent peut interroger n'importe quel champ (`generic_query`), produire un graphique (`generic_chart`)
  ou une statistique numérique (`generic_stat`) sur les collections `demandes`, `bookings`, `users`,
  `orders` — toujours avant de répondre qu'une information n'existe pas.
- **Contrôle d'accès** : permissions par module (`users`, `demandes`, `entretiens`, `orders`, `admins`) ×
  action (`read`/`create`/`update`), avec un rôle super admin qui outrepasse tout, et des fonctions
  réservées au super admin.
- **Robustesse** : détection et exécution de secours des faux appels d'outils texte que le modèle peut
  halluciner, relance automatique sur réponse vide/tronquée, historique de conversation par admin.

## Stack technique

| Composant       | Rôle                                                          |
|-----------------|----------------------------------------------------------------|
| FastAPI         | API HTTP exposée au front-end                                  |
| Ollama Cloud    | Modèle LLM (`gemma4:31b-cloud`) avec function-calling           |
| MongoDB (pymongo) | Base de données — seule source de vérité pour l'agent         |
| PyJWT           | Vérification des tokens émis par le backend Spring Boot existant |
| bcrypt          | Hash des mots de passe générés (création/réinitialisation)     |
| openpyxl        | Génération des exports Excel                                   |
| httpx           | Appels vers le backend Spring Boot (mise à jour de statut)      |

## Installation

### Prérequis

- Python 3.12+
- Une instance MongoDB accessible (déjà peuplée par la plateforme existante)
- Une clé API [Ollama Cloud](https://ollama.com)
- Le backend Spring Boot de la plateforme (pour les mises à jour de statut de demande et pour émettre
  les JWT utilisés par l'authentification)

### Étapes

```bash
# 1. Cloner et se placer dans le dossier du projet
cd chatbot-immigration

# 2. Créer et activer un environnement virtuel
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # Linux/macOS

# 3. Installer les dépendances
pip install -r requirements.txt

# 4. Configurer les variables d'environnement
copy .env.example .env          # Windows
# cp .env.example .env          # Linux/macOS
# puis éditer .env avec les vraies valeurs (voir ci-dessous)

# 5. Lancer le serveur en développement
uvicorn main:app --reload --port 8000
```

### Variables d'environnement (`.env`)

Voir [.env.example](.env.example) pour le détail. **Ne jamais committer `.env`** (il est dans
`.gitignore`) — il contient notamment les identifiants MongoDB et le secret JWT.

| Variable           | Description                                                              |
|---------------------|---------------------------------------------------------------------------|
| `OLLAMA_API_KEY`    | Clé API pour appeler le modèle sur Ollama Cloud                          |
| `MONGO_URI`         | URI de connexion MongoDB (avec identifiants)                             |
| `MONGO_DB_NAME`     | Nom de la base MongoDB à utiliser                                        |
| `JWT_SECRET`        | Secret de signature JWT en base64, **doit correspondre** à celui du backend Spring Boot |
| `BACKEND_BASE_URL`  | URL de base du backend Spring Boot (défaut : `http://localhost:8083`)    |

### Docker

```bash
docker build -t chatbot-immigration .
docker run -p 8000:8000 --env-file .env chatbot-immigration
```

## API

Toutes les routes nécessitent un header `Authorization: Bearer <token_JWT>` (le token émis par le
backend Spring Boot, avec la claim `sub` = email de l'admin).

| Méthode | Route                              | Description                                               |
|---------|-------------------------------------|-------------------------------------------------------------|
| POST    | `/api/v1/chat`                      | Envoie un message `{ "message": "..." }`, retourne `{ "reponse", "chart", "utilisateur" }` |
| POST    | `/api/v1/chat/reset`                | Réinitialise la session de l'admin (historique, confirmations en attente, etc.) |
| GET     | `/api/v1/download/{filename}`       | Télécharge un export Excel généré précédemment              |

## Modèle de permissions

Le document utilisateur en base (`users`) porte un champ `permissions` structuré par module :

```json
{
  "isSuperAdmin": false,
  "permissions": {
    "demandes": { "read": true, "create": false, "update": true },
    "entretiens": { "read": true, "create": true, "update": true },
    "users": { "read": true, "create": false, "update": false },
    "orders": { "read": true, "create": false, "update": false },
    "admins": { "read": false, "create": false, "update": false }
  }
}
```

Un `isSuperAdmin: true` outrepasse toutes les vérifications de module, mais reste bloqué sur certaines
actions protégées (ex. : impossible de bloquer ou modifier un autre super admin). Certaines fonctions
(`block_admin`, `unblock_admin`, `update_admin_permission`, `generic_query`, `generic_chart`,
`generic_stat`) sont réservées exclusivement au super admin.

## Sécurité et données sensibles

- `.env`, `venv/`, `__pycache__/` et `exports/` (fichiers Excel générés, contenant des données de
  candidats) sont exclus du dépôt via `.gitignore`.
- `.claude/settings.local.json` (paramètres locaux à la machine du développeur) est également exclu.
- Plusieurs scripts de debug ponctuels (`debug_booking.py`, `debug_trabelsi.py`, `inspect_user.py`)
  contenaient des données réelles codées en dur (noms de candidats, adresse email) — ils ont été retirés
  du suivi git et ajoutés au `.gitignore`. **Ils restent présents sur le disque** pour un usage local,
  mais ne seront plus commités. Si ce dépôt a déjà été poussé vers un remote partagé avant ce nettoyage,
  ces données restent présentes dans l'historique Git existant et nécessiteraient une réécriture
  d'historique pour être purgées complètement.
- Ne jamais committer de token JWT, de clé API ou d'URI MongoDB en dur dans le code — toujours passer
  par les variables d'environnement.

## Scripts utilitaires (hors application)

Quelques scripts de développement, non utilisés par l'application elle-même (`main.py` ne les importe
pas), servent à inspecter la base MongoDB à la main : `check_db.py`, `inspect_db.py`,
`inspect_config.py`, `inspect_orders.py`, `test_mongo.py`, `test_ollama.py`. À lancer directement avec
`python nom_du_script.py` après avoir configuré `.env`.

## Structure du projet

```
chatbot-immigration/
├── main.py              # API FastAPI (routes HTTP)
├── auth.py              # Vérification JWT + chargement des permissions
├── chat.py              # Agent LLM : prompt, outils, boucle de function-calling
├── db.py                # Accès MongoDB (toutes les fonctions "outils")
├── requirements.txt     # Dépendances Python
├── Dockerfile
├── .dockerignore
├── .gitignore
├── .env.example          # Modèle de configuration (à copier en .env)
├── check_db.py / inspect_*.py / test_*.py   # Scripts de debug ponctuels (voir ci-dessus)
└── exports/              # Fichiers Excel générés à la demande (ignoré par git)
```


## 🔗 Dépôts du projet

Ce dépôt fait partie de la **plateforme d'administration**, composée de trois dépôts :

| Partie | Rôle | Dépôt |
|---|---|---|
| Backend | API Spring Boot + MongoDB + JWT (port `8083`) | [jaz56/admin-backend](https://github.com/jaz56/admin-backend) |
| Frontend | Interface d'administration Angular 21 (port `4200`) | [jaz56/admin-front](https://github.com/jaz56/admin-front) |
| Chatbot | Assistant IA FastAPI + Ollama (port `8000`) | [jaz56/chatbot-immigration](https://github.com/jaz56/chatbot-immigration) |
