import difflib
import os
from dotenv import load_dotenv
from pymongo import MongoClient
from datetime import datetime, timedelta, timezone
from bson import ObjectId
from bson.errors import InvalidId
import bcrypt
import secrets
import re
import string
import httpx
import unicodedata
import openpyxl
from openpyxl.styles import Font
import uuid
from bson.decimal128 import Decimal128
EXPORT_DIR = "exports"
os.makedirs(EXPORT_DIR, exist_ok=True)

DEMANDE_EXPORT_FIELDS = {
    "nom": "Nom",
    "prenom": "Prénom",
    "email": "Email",
    "telephone": "Téléphone",
    "candidat_id": "ID Candidat",
    "uniqueId": "ID Demande",
    "status": "Statut",
    "progress": "Progression",
    "pays": "Pays résidence",
    "createdAt": "Date création",
}
SENSITIVE_USER_FIELDS = {"password", "_class"}
load_dotenv()
VALID_MODULES = {"users", "demandes", "entretiens", "orders", "admins"}
VALID_ACTIONS = {"read", "create", "update"}
BACKEND_BASE_URL = os.environ.get("BACKEND_BASE_URL", "http://localhost:8083")
client = MongoClient(os.environ['MONGO_URI'])
GENERIC_QUERY_ALLOWED_COLLECTIONS = {"demandes", "bookings", "users", "orders"}
GENERIC_QUERY_EXCLUDED_FIELDS = {"password"}
USER_DOCUMENT_LABELS = {
    "photoDeProfile": "Photo de profil",
    "cinRecto": "CIN recto",
    "cinVerso": "CIN verso",
    "vocal": "Échantillon vocal",
}

BOOKING_DOCUMENT_LABELS = {
    "interviewRecording": "Enregistrement de l'entretien",
    "interviewReport": "Rapport d'entretien",
}
db = client[os.environ['MONGO_DB_NAME']]


def count_users():
    """Retourne le nombre total d'utilisateurs enregistrés."""
    return db.users.count_documents({})


def generate_temp_password() -> str:
    """Génère un mot de passe temporaire aléatoire (12 caractères)."""
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(12))


def _normalize_text(text: str) -> str:
    """Met en minuscule, retire les accents, et uniformise les séparateurs (_, -) en espaces, pour une comparaison robuste."""
    text = text.strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("_", " ").replace("-", " ")
    text = " ".join(text.split())
    return text


def get_valid_values(category: str) -> list[str]:
    """Retourne la liste des valeurs techniques actives autorisées pour une catégorie config_items donnée."""
    items = db.config_items.find({"category": category, "active": True}, {"value": 1, "_id": 0})
    return [item["value"] for item in items]


def resolve_config_value(category: str, user_input: str) -> str | None:
    """
    Traduit ce que l'admin/le modèle a écrit vers la vraie valeur technique stockée en base,
    en cherchant dans value ET label de config_items.
    """
    items = list(db.config_items.find({"category": category, "active": True}, {"value": 1, "label": 1, "_id": 0}))
    normalized_input = _normalize_text(user_input)

    for item in items:
        if item.get("value") == user_input:
            return item["value"]

    for item in items:
        if _normalize_text(item.get("value", "")) == normalized_input:
            return item["value"]
        if _normalize_text(item.get("label", "")) == normalized_input:
            return item["value"]

    return None


def resolve_against_real_values(collection_name: str, field: str, user_input: str) -> str | None:
    """
    Pour les fonctions de LECTURE : résout un mot dit par l'admin contre les valeurs
    RÉELLEMENT présentes dans la collection (via distinct), pas contre config_items.
    Ne sert jamais à valider une ÉCRITURE.
    """
    coll = db[collection_name]
    real_values = coll.distinct(field)
    normalized_input = _normalize_text(user_input)

    for val in real_values:
        if val == user_input:
            return val
    for val in real_values:
        if _normalize_text(str(val)) == normalized_input:
            return val
    return None


def _serialize_value(value):
    """Convertit récursivement n'importe quelle valeur MongoDB (sous-document, liste, Decimal128, ObjectId, date) en JSON-safe."""
    if isinstance(value, dict):
        return {k: _serialize_value(v) for k, v in value.items() if k not in ("_id", "_class")}
    if isinstance(value, list):
        return [_serialize_value(v) for v in value]
    if isinstance(value, Decimal128):
        return str(value)
    if isinstance(value, ObjectId):
        return str(value)
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y %H:%M")
    return value


def _build_file_url(path: str) -> str:
    """Construit l'URL complète d'un fichier stocké, à partir d'un chemin relatif ou absolu."""
    if not path:
        return None
    if path.startswith("http://") or path.startswith("https://"):
        return path
    if not path.startswith("/"):
        path = "/" + path
    return f"{BACKEND_BASE_URL}{path}"


def find_user_by_name_or_email(query: str):
    words = [w for w in query.strip().split() if w]
    if not words:
        return None

    normalized_words = [_normalize_text(w) for w in words]  # 🟢 déplacé ici, avant utilisation

    or_conditions = []
    for word in words:
        regex = {"$regex": word, "$options": "i"}
        or_conditions.extend([{"nom": regex}, {"prenom": regex}, {"email": regex}])

    candidates = list(db.users.find({"$or": or_conditions}).limit(25))
    if not candidates:
        all_users = list(db.users.find({}, {"nom": 1, "prenom": 1, "email": 1, "role": 1, "isBlocked": 1}))
        close_matches = []
        for u in all_users:
            nom = _normalize_text(u.get("nom", ""))
            prenom = _normalize_text(u.get("prenom", ""))
            for w in normalized_words:
                if difflib.SequenceMatcher(None, w, nom).ratio() > 0.8 or difflib.SequenceMatcher(None, w, prenom).ratio() > 0.8:
                    close_matches.append(u)
                    break
        candidates = close_matches
        if not candidates:
            return None

    def to_dict(u):
        return {
            "id": str(u["_id"]), "nom": u.get("nom"), "prenom": u.get("prenom"),
            "email": u.get("email"), "role": u.get("role"), "isBlocked": u.get("isBlocked", False),
        }

    # 🟢 Priorité 1 : correspondance exacte nom+prénom. Si PLUSIEURS candidats matchent
    # exactement (homonymes), c'est une vraie ambiguïté — on la signale au lieu de deviner.
    exact_matches = []
    for u in candidates:
        nom = _normalize_text(u.get("nom", ""))
        prenom = _normalize_text(u.get("prenom", ""))
        full_fields = {nom, prenom}
        if len(normalized_words) >= 2 and all(w in full_fields for w in normalized_words):
            exact_matches.append(u)

    if len(exact_matches) == 1:
        return to_dict(exact_matches[0])
    if len(exact_matches) > 1:
        return {
            "ambiguous": True,
            "candidates": [
                {"nom": u.get("nom"), "prenom": u.get("prenom"), "email": u.get("email")}
                for u in exact_matches
            ],
        }

    def evaluate(u):
        nom = _normalize_text(u.get("nom", ""))
        prenom = _normalize_text(u.get("prenom", ""))
        email = _normalize_text(u.get("email", ""))
        nom_hit = False
        matched_words = 0
        for w in words:
            nw = _normalize_text(w)
            if not nw:
                continue
            hit = False
            if nw in nom:
                nom_hit = True
                hit = True
            if nw in prenom or nw in email:
                hit = True
            if hit:
                matched_words += 1
        return nom_hit, matched_words

    scored = []
    for u in candidates:
        nom_hit, matched_words = evaluate(u)
        if nom_hit or matched_words == len(words):
            scored.append((matched_words, nom_hit, u))

    if not scored:
        return None

    scored.sort(key=lambda x: (x[1], x[0]), reverse=True)
    top = scored[0]

    if len(scored) > 1 and (scored[1][0], scored[1][1]) == (top[0], top[1]):
        return None

    return to_dict(top[2])
def _find_one_by_user_ref(collection, user_id_str: str):
    """
    Cherche un document où le champ 'user' correspond à l'utilisateur donné,
    en tentant à la fois la chaîne et l'ObjectId (les données existantes sont incohérentes
    sur ce typage selon la manière dont elles ont été insérées).
    """
    doc = collection.find_one({"user": user_id_str})
    if doc:
        return doc
    try:
        doc = collection.find_one({"user": ObjectId(user_id_str)})
    except (InvalidId, TypeError):
        doc = None
    return doc


def _find_many_by_user_ref(collection, user_id_str: str, **extra_filters):
    """Version 'find' (plusieurs résultats) : combine TOUJOURS les résultats en chaîne et en ObjectId, pour ne rater aucun document quel que soit le typage utilisé à l'insertion."""
    results = list(collection.find({"user": user_id_str, **extra_filters}))

    seen_ids = {r["_id"] for r in results}
    try:
        oid_results = list(collection.find({"user": ObjectId(user_id_str), **extra_filters}))
        for r in oid_results:
            if r["_id"] not in seen_ids:
                results.append(r)
                seen_ids.add(r["_id"])
    except (InvalidId, TypeError):
        pass

    return results


def _find_user_by_id(user_ref):
    """Cherche un utilisateur par son _id, stocké en ObjectId ou en chaîne selon les documents."""
    try:
        return db.users.find_one({"_id": ObjectId(user_ref)})
    except (InvalidId, TypeError):
        return db.users.find_one({"_id": user_ref})


def _resolve_user_or_error(user_query: str):
    """
    Résout un candidat/admin par nom/prénom/email (find_user_by_name_or_email + gestion ambiguïté).
    Retourne le dict utilisateur (résumé) si trouvé, sinon un message d'erreur (str) à renvoyer tel quel.
    """
    user = find_user_by_name_or_email(user_query)
    if user and user.get("ambiguous"):
        return format_ambiguous_error(user)
    if not user:
        return f"Aucun utilisateur trouvé correspondant à '{user_query}'."
    return user


def _resolve_full_user(user_query: str):
    """
    Comme _resolve_user_or_error, mais récupère aussi le document utilisateur complet.
    Retourne (résumé, document complet) si trouvé, sinon un message d'erreur (str).
    """
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user
    full_doc = db.users.find_one({"_id": ObjectId(user["id"])})
    if not full_doc:
        return f"Aucun utilisateur trouvé correspondant à '{user_query}'."
    return user, full_doc


def _set_user_blocked(user_id: str, blocked: bool):
    """Met à jour le statut isBlocked d'un utilisateur (admin ou candidat)."""
    db.users.update_one(
        {"_id": ObjectId(user_id)},
        {"$set": {"isBlocked": blocked, "blockedUntil": None, "updatedAt": datetime.now(timezone.utc)}}
    )


def _count_with_resolved_value(collection_name: str, field: str, value: str, resolver) -> int:
    """Compte les documents dont `field` correspond à `value`, après résolution via `resolver`."""
    resolved = resolver(value)
    return db[collection_name].count_documents({field: resolved if resolved else value})


def _breakdown_pipeline(collection_name: str, field: str) -> list:
    """Exécute le pipeline d'agrégation de répartition (count par valeur) pour un champ, trié par count décroissant."""
    pipeline = [
        {"$group": {"_id": f"${field}", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
    ]
    return list(db[collection_name].aggregate(pipeline))


def _breakdown(collection_name: str, field: str) -> dict:
    """Retourne la répartition (count par valeur) d'un champ, sous forme de dict {valeur: count}."""
    return {r["_id"]: r["count"] for r in _breakdown_pipeline(collection_name, field)}


def _distribution_chart(collection_name: str, field: str, chart_type: str, title: str, sort_by_id: bool = False) -> dict:
    """Construit les données d'un graphique de répartition (count par valeur) pour un champ d'une collection."""
    sort_stage = {"_id": 1} if sort_by_id else {"count": -1}
    pipeline = [
        {"$group": {"_id": f"${field}", "count": {"$sum": 1}}},
        {"$sort": sort_stage},
    ]
    results = list(db[collection_name].aggregate(pipeline))
    return {
        "type": chart_type,
        "title": title,
        "labels": [r["_id"] or "Non défini" for r in results],
        "values": [r["count"] for r in results],
    }


def _trend_chart(collection_name: str, days: int, title: str) -> dict:
    """Construit les données d'un graphique d'évolution jour par jour sur les N derniers jours."""
    since = datetime.now(timezone.utc) - timedelta(days=days)
    pipeline = [
        {"$match": {"createdAt": {"$gte": since}}},
        {"$group": {"_id": {"$dateToString": {"format": "%Y-%m-%d", "date": "$createdAt"}}, "count": {"$sum": 1}}},
        {"$sort": {"_id": 1}},
    ]
    results = list(db[collection_name].aggregate(pipeline))
    return {
        "type": "line",
        "title": title,
        "labels": [r["_id"] for r in results],
        "values": [r["count"] for r in results],
    }


def count_by_config_category(category: str):
    """Retourne le nombre de valeurs actives pour une catégorie de config_items donnée."""
    return db.config_items.count_documents({"category": category, "active": True})


def count_demandes_by_status(status: str):
    """Compte le nombre de demandes ayant un statut donné, résolu contre les vraies valeurs en base."""
    return _count_with_resolved_value("demandes", "status", status, lambda v: resolve_against_real_values("demandes", "status", v))


def count_demandes_by_progress(progress: str):
    """Compte le nombre de demandes à une étape de progression donnée."""
    return _count_with_resolved_value("demandes", "progress", progress, lambda v: resolve_config_value("demande_progress_step", v))


def count_demandes_by_country(country_label: str):
    """Compte le nombre de demandes pour un pays de résidence donné."""
    return db.demandes.count_documents({"paysResidence.label": country_label})


def count_bookings_by_status(status: str):
    """Compte le nombre d'entretiens ayant un statut donné, résolu contre les vraies valeurs en base."""
    return _count_with_resolved_value("bookings", "status", status, lambda v: resolve_against_real_values("bookings", "status", v))


def count_bookings_by_payment_status(payment_status: str):
    """Compte le nombre de réservations selon leur statut de paiement."""
    return _count_with_resolved_value("bookings", "paymentStatus", payment_status, lambda v: resolve_config_value("statut_paiement", v))


def average_interview_score():
    """Calcule la note moyenne des entretiens complétés."""
    pipeline = [
        {"$match": {"interviewCompleted": True}},
        {"$group": {"_id": None, "avg": {"$avg": {"$toDouble": "$interviewScore"}}}}
    ]
    result = list(db.bookings.aggregate(pipeline))
    return round(result[0]["avg"], 2) if result else None


def get_user_permissions(email: str):
    """Retourne le document utilisateur complet nécessaire aux vérifications de permissions."""
    return db.users.find_one({"email": email})


def count_demandes_total():
    """Retourne le nombre total de demandes, tous statuts confondus."""
    return db.demandes.count_documents({})


def count_bookings_total():
    """Retourne le nombre total d'entretiens/réservations, tous statuts confondus."""
    return db.bookings.count_documents({})


def count_users_by_role(role: str):
    """Compte le nombre d'utilisateurs ayant un rôle donné."""
    return db.users.count_documents({"role": role})


def count_demandes_created_since(days: int):
    """Compte le nombre de demandes créées au cours des N derniers jours."""
    since = datetime.now(timezone.utc) - timedelta(days=days)
    return db.demandes.count_documents({"createdAt": {"$gte": since}})


def count_bookings_created_since(days: int):
    """Compte le nombre d'entretiens/réservations créés au cours des N derniers jours."""
    since = datetime.now(timezone.utc) - timedelta(days=days)
    return db.bookings.count_documents({"createdAt": {"$gte": since}})

def create_candidat(email: str, nom: str, prenom: str):
    """Crée un nouveau candidat avec un mot de passe temporaire."""
    existing = db.users.find_one({"email": email})
    if existing:
        return f"Un utilisateur avec l'email {email} existe déjà."

    temp_password = generate_temp_password()
    hashed = bcrypt.hashpw(temp_password.encode(), bcrypt.gensalt(rounds=10)).decode()

    user_doc = {
        "email": email,
        "nom": nom,
        "prenom": prenom,
        "password": hashed,
        "role": "candidat",
        "acceptedTerms": False,
        "nvRegister": True,
        "verifiedEmail": False,
        "isSuperAdmin": False,
        "isBlocked": False,
        "candidateVerificationStatus": "pending",
        "candidateVerified": False,
        "balance": 0.0,
        "cashback": 0.0,
        "interessesPar": [],
        "favoriteJobs": [],
        "subscriptions": [],
        "providers": [],
        "balanceTransactions": [],
        "cashbackTransactions": [],
        "createdAt": datetime.now(timezone.utc),
        "updatedAt": datetime.now(timezone.utc),
        "_class": "com.yassmine.administration.model.User",
    }
    result = db.users.insert_one(user_doc)

    return {
        "message": "Candidat créé avec succès",
        "id": str(result.inserted_id),
        "email": email,
        "motDePasseTemporaire": temp_password,
    }


def create_demande(user_query: str):
    """Crée une nouvelle demande vide pour un candidat existant, recherché par nom/prénom/email."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    existing = _find_one_by_user_ref(db.demandes, user["id"])
    if existing:
        return f"{user['prenom']} {user['nom']} a déjà une demande existante (ID: {existing.get('uniqueId')})."

    unique_id = f"{datetime.now().strftime('%y-%m')}-T-{secrets.token_hex(4).upper()}"
    demande_doc = {
        "user": user["id"],
        "uniqueId": unique_id,
        "status": "en_attente",
        "progress": "STEP_1",
        "eligibiliteNote": "",
        "ouverteTouteOpportunites": False,
        "besoinVisa": False,
        "typeHebergement": False,
        "maladieContagieuse": False,
        "handicape": False,
        "preinscription": False,
        "existenceDeGarant": False,
        "diplomes": [],
        "connaissanceLinguistique": [],
        "anneesDeEtudeAFinancier": [],
        "createdAt": datetime.now(timezone.utc),
        "updatedAt": datetime.now(timezone.utc),
        "_class": "com.yassmine.administration.model.Demande",
    }
    db.demandes.insert_one(demande_doc)

    return {
        "message": "Demande créée avec succès",
        "candidat": f"{user['prenom']} {user['nom']}",
        "uniqueId": unique_id,
        "status": "en_attente",
    }


def create_booking(user_query: str, appointment_date: str, appointment_time: str, appointment_type: str = "standard"):
    """Crée un nouvel entretien lié à la demande d'un candidat existant. appointment_date au format YYYY-MM-DD."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    demande = _find_one_by_user_ref(db.demandes, user["id"])
    if not demande:
        return f"{user['prenom']} {user['nom']} n'a pas encore de demande — impossible de créer un entretien sans dossier associé."

    try:
        date_obj = datetime.strptime(appointment_date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    except ValueError:
        return f"Format de date invalide : '{appointment_date}'. Utilisez le format AAAA-MM-JJ."

    unique_id = f"BK{secrets.token_hex(3).upper()}"
    booking_doc = {
        "user": user["id"],
        "demande": str(demande["_id"]),
        "uniqueId": unique_id,
        "appointmentType": appointment_type,
        "appointmentDate": date_obj,
        "appointmentTime": appointment_time,
        "price": 0,
        "paymentStatus": "unpaid",
        "status": "pending",
        "interviewCompleted": False,
        "interviewRecording": "",
        "interviewReport": "",
        "documents": {},
        "documentsMetadata": {},
        "createdAt": datetime.now(timezone.utc),
        "updatedAt": datetime.now(timezone.utc),
        "_class": "com.yassmine.administration.model.Booking",
    }
    db.bookings.insert_one(booking_doc)

    return {
        "message": "Entretien créé avec succès",
        "candidat": f"{user['prenom']} {user['nom']}",
        "uniqueId": unique_id,
        "date": appointment_date,
        "heure": appointment_time,
    }

def update_demande_status_unified(new_status: str, token: str, user_query: str = None, unique_id: str = None):
    """Modifie le statut d'une demande, identifiée soit par candidat (user_query) soit par uniqueId, via le vrai backend."""
    valid_values = get_valid_values("statut_demande")
    if new_status not in valid_values:
        return f"Statut invalide : '{new_status}'. Valeurs autorisées : {', '.join(valid_values)}."

    if unique_id:
        demande = db.demandes.find_one({"uniqueId": unique_id})
        if not demande:
            return f"Aucune demande trouvée avec l'identifiant '{unique_id}'."
    elif user_query:
        user = _resolve_user_or_error(user_query)
        if isinstance(user, str):
            return user
        demande = _find_one_by_user_ref(db.demandes, user["id"])
        if not demande:
            return f"Aucune demande trouvée pour {user['prenom']} {user['nom']}."
    else:
        return "Il faut préciser soit le candidat, soit l'identifiant de la demande."

    return _call_backend_update_status(str(demande["_id"]), new_status, token)

def update_demande_progress(user_query: str, new_progress: str, token: str):
    """Modifie l'étape de progression d'une demande d'un candidat existant, via le vrai backend."""
    valid_values = get_valid_values("demande_progress_step")
    if new_progress not in valid_values:
        return f"Étape de progression invalide : '{new_progress}'. Valeurs autorisées : {', '.join(valid_values)}."

    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    demande = _find_one_by_user_ref(db.demandes, user["id"])
    if not demande:
        return f"Aucune demande trouvée pour {user['prenom']} {user['nom']}."

    return _call_backend_update_status(str(demande["_id"]), demande.get("status"), token, new_progress)

def update_booking_status_unified(new_status: str, user_query: str = None, unique_id: str = None):
    """Modifie le statut d'un entretien, identifié soit par candidat soit par uniqueId."""
    resolved_status = resolve_config_value("statut_entretien", new_status)
    if resolved_status is None:
        valid_values = get_valid_values("statut_entretien")
        return f"Statut invalide : '{new_status}'. Valeurs autorisées : {', '.join(valid_values)}."

    if unique_id:
        booking = db.bookings.find_one({"uniqueId": unique_id})
        if not booking:
            return f"Aucun entretien trouvé avec l'identifiant '{unique_id}'."
        candidat_label = unique_id
    elif user_query:
        user = _resolve_user_or_error(user_query)
        if isinstance(user, str):
            return user
        booking = _find_one_by_user_ref(db.bookings, user["id"])
        if not booking:
            return f"Aucun entretien trouvé pour {user['prenom']} {user['nom']}."
        candidat_label = f"{user['prenom']} {user['nom']}"
    else:
        return "Il faut préciser soit le candidat, soit l'identifiant de l'entretien."

    old_status = booking.get("status")
    db.bookings.update_one({"_id": booking["_id"]}, {"$set": {"status": resolved_status, "updatedAt": datetime.now(timezone.utc)}})
    return {"message": f"Statut de l'entretien de {candidat_label} modifié : {old_status} → {resolved_status}"}

def demandes_breakdown_by_status():
    """Retourne la répartition du nombre de demandes par statut, déjà formatée en texte."""
    results = _breakdown_pipeline("demandes", "status")
    if not results:
        return "Aucune demande enregistrée."
    lines = [f"- {r['_id']} : {r['count']}" for r in results]
    return "\n".join(lines)


def demandes_breakdown_by_progress():
    """Retourne la répartition du nombre de demandes par étape de progression."""
    return _breakdown("demandes", "progress")


def demandes_breakdown_by_country():
    """Retourne la répartition du nombre de demandes par pays de résidence."""
    return _breakdown("demandes", "paysResidence.label")


def bookings_breakdown_by_status():
    """Retourne la répartition du nombre d'entretiens/réservations par statut."""
    return _breakdown("bookings", "status")


def bookings_breakdown_by_payment_status():
    """Retourne la répartition du nombre de réservations par statut de paiement."""
    return _breakdown("bookings", "paymentStatus")


def users_breakdown_by_role():
    """Retourne la répartition du nombre d'utilisateurs par rôle."""
    return _breakdown("users", "role")


def block_admin(user_query: str, reason: str = ""):
    """Bloque indéfiniment le compte d'un admin, recherché par nom/prénom/email. Ne bloque jamais un super admin."""
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user, full_doc = resolved

    if full_doc.get("isSuperAdmin"):
        return "Action refusée : un compte Super Administrateur ne peut pas être bloqué."

    _set_user_blocked(user["id"], True)

    return {
        "message": f"Compte de {user['prenom']} {user['nom']} bloqué" + (f" (motif : {reason})" if reason else ""),
        "utilisateur": f"{user['prenom']} {user['nom']}",
    }


def unblock_admin(user_query: str):
    """Débloque le compte d'un admin, recherché par nom/prénom/email."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    _set_user_blocked(user["id"], False)

    return {
        "message": f"Compte de {user['prenom']} {user['nom']} débloqué",
        "utilisateur": f"{user['prenom']} {user['nom']}",
    }


def update_admin_permission(user_query: str, module: str, action: str, value: bool):
    """Modifie un droit précis (read/create/update) sur un module précis pour un admin existant."""
    if module not in VALID_MODULES:
        return f"Module invalide : '{module}'. Modules valides : {', '.join(sorted(VALID_MODULES))}."
    if action not in VALID_ACTIONS:
        return f"Action invalide : '{action}'. Actions valides : {', '.join(sorted(VALID_ACTIONS))}."

    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user, full_doc = resolved

    if full_doc.get("isSuperAdmin"):
        return "Action refusée : les permissions d'un Super Administrateur ne peuvent pas être modifiées."
    if full_doc.get("role") != "admin":
        return f"{user['prenom']} {user['nom']} n'est pas un administrateur — les permissions par module ne s'appliquent qu'aux comptes admin."

    field_path = f"permissions.{module}.{action}"
    db.users.update_one(
        {"_id": ObjectId(user["id"])},
        {"$set": {field_path: value, "updatedAt": datetime.now(timezone.utc)}}
    )

    return {
        "message": f"Droit '{action}' sur le module '{module}' de {user['prenom']} {user['nom']} mis à : {value}",
        "utilisateur": f"{user['prenom']} {user['nom']}",
        "module": module,
        "action": action,
        "valeur": value,
    }


def count_candidats_pending_verification():
    """Compte les candidats dont la vérification est en attente."""
    return db.users.count_documents({"role": "candidat", "candidateVerificationStatus": "pending"})


def count_blocked_users():
    """Compte le nombre total de comptes actuellement bloqués."""
    return db.users.count_documents({"isBlocked": True})


def count_bookings_this_week():
    """Compte le nombre d'entretiens/réservations prévus dans les 7 prochains jours."""
    now = datetime.now(timezone.utc)
    in_a_week = now + timedelta(days=7)
    return db.bookings.count_documents({"appointmentDate": {"$gte": now, "$lte": in_a_week}})


def reset_candidat_password(user_query: str):
    """Réinitialise le mot de passe d'un candidat avec un nouveau mot de passe temporaire généré."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user
    if user.get("role") != "candidat":
        return f"{user['prenom']} {user['nom']} n'est pas un candidat — cette action est réservée aux comptes candidats."

    new_password = generate_temp_password()
    hashed = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt(rounds=10)).decode()

    db.users.update_one(
        {"_id": ObjectId(user["id"])},
        {"$set": {"password": hashed, "updatedAt": datetime.now(timezone.utc)}}
    )

    return {
        "message": f"Mot de passe de {user['prenom']} {user['nom']} réinitialisé",
        "utilisateur": f"{user['prenom']} {user['nom']}",
        "nouveauMotDePasseTemporaire": new_password,
    }


def orders_breakdown_by_status():
    """Retourne la répartition du nombre de commandes/orders par statut."""
    return _breakdown("orders", "status")


def get_demande_by_unique_id(unique_id: str):
    """Retourne les détails d'une demande recherchée par son identifiant unique exact."""
    demande = db.demandes.find_one({"uniqueId": unique_id})
    if not demande:
        return f"Aucune demande trouvée avec l'identifiant '{unique_id}'."

    demande_user = demande.get("user")
    user_doc = _find_user_by_id(demande_user) if demande_user else None
    candidat = f"{user_doc.get('prenom')} {user_doc.get('nom')}" if user_doc else "Inconnu"

    return {
        "candidat": candidat,
        "uniqueId": demande.get("uniqueId"),
        "status": demande.get("status"),
        "progress": demande.get("progress"),
        "pays": demande.get("paysResidence", {}).get("label"),
        "createdAt": str(demande.get("createdAt")),
    }


def get_booking_by_unique_id(unique_id: str):
    """Retourne les détails d'un entretien/booking recherché par son identifiant unique exact."""
    booking = db.bookings.find_one({"uniqueId": unique_id})
    if not booking:
        return f"Aucun entretien trouvé avec l'identifiant '{unique_id}'."

    booking_user = booking.get("user")
    user_doc = _find_user_by_id(booking_user) if booking_user else None
    candidat = f"{user_doc.get('prenom')} {user_doc.get('nom')}" if user_doc else "Inconnu"

    return {
        "candidat": candidat,
        "uniqueId": booking.get("uniqueId"),
        "status": booking.get("status"),
        "appointmentType": booking.get("appointmentType"),
        "appointmentDate": str(booking.get("appointmentDate")),
        "paymentStatus": booking.get("paymentStatus"),
        "interviewCompleted": booking.get("interviewCompleted"),
        "interviewScore": str(booking.get("interviewScore")) if booking.get("interviewScore") else None,
    }


def block_candidat(user_query: str, reason: str = ""):
    """Bloque le compte d'un candidat, recherché par nom/prénom/email."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user
    if user.get("role") != "candidat":
        return f"{user['prenom']} {user['nom']} n'est pas un candidat — utilisez le blocage admin pour un compte staff."

    _set_user_blocked(user["id"], True)

    return {
        "message": f"Compte candidat de {user['prenom']} {user['nom']} bloqué" + (f" (motif : {reason})" if reason else ""),
        "utilisateur": f"{user['prenom']} {user['nom']}",
    }


def unblock_candidat(user_query: str):
    """Débloque le compte d'un candidat, recherché par nom/prénom/email."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user
    if user.get("role") != "candidat":
        return f"{user['prenom']} {user['nom']} n'est pas un candidat — utilisez le déblocage admin pour un compte staff."

    _set_user_blocked(user["id"], False)

    return {
        "message": f"Compte candidat de {user['prenom']} {user['nom']} débloqué",
        "utilisateur": f"{user['prenom']} {user['nom']}",
    }


def count_bookings_tomorrow():
    """Compte le nombre d'entretiens/réservations prévus demain."""
    now = datetime.now(timezone.utc)
    tomorrow_start = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    tomorrow_end = tomorrow_start + timedelta(days=1)
    return db.bookings.count_documents({"appointmentDate": {"$gte": tomorrow_start, "$lt": tomorrow_end}})


def count_users_by_email_verification(verified: bool):
    """Compte le nombre d'utilisateurs ayant vérifié (ou non) leur email."""
    return db.users.count_documents({"verifiedEmail": verified})


def _call_backend_update_status(demande_id: str, new_status: str, token: str, new_progress: str = None):
    """Appelle le vrai endpoint Spring Boot PATCH /api/demandes/{id}/status."""
    url = f"{BACKEND_BASE_URL}/api/demandes/{demande_id}/status"
    payload = {"status": new_status}
    if new_progress:
        payload["progress"] = new_progress
    headers = {"Authorization": f"Bearer {token}"}
    try:
        response = httpx.patch(url, json=payload, headers=headers, timeout=10.0)
        response.raise_for_status()
        return {
            "message": f"Statut mis à jour vers '{new_status}' via le backend (logique métier appliquée : synchronisation automatique incluse)",
        }
    except httpx.HTTPStatusError as e:
        return f"Le backend a refusé la modification : {e.response.status_code} - {e.response.text}"
    except httpx.RequestError as e:
        return f"Impossible de contacter le backend Spring Boot : {e}"


def count_orders_by_status(status: str):
    """Compte le nombre de commandes ayant un statut donné."""
    return db.orders.count_documents({"status": status})


def count_orders_by_type(order_type: str):
    """Compte le nombre de commandes d'un type donné."""
    return db.orders.count_documents({"type": order_type})


def get_recent_orders(limit: int = 5):
    """Retourne les N commandes les plus récemment créées."""
    docs = db.orders.find().sort("createdAt", -1).limit(limit)
    return [
        {
            "orderId": d.get("order_id"),
            "status": d.get("status"),
            "type": d.get("type"),
            "dateRendezVous": d.get("date"),
            "dateCreation": str(d.get("createdAt")),
        }
        for d in docs
    ]


def filter_demandes(status: str = None, progress: str = None, country_label: str = None):
    """Cherche les demandes correspondant aux filtres donnés (tous optionnels), avec les infos du candidat associé."""
    query = {}
    if status:
        resolved = resolve_against_real_values("demandes", "status", status)
        query["status"] = resolved if resolved else status
    if progress:
        resolved = resolve_config_value("demande_progress_step", progress)
        query["progress"] = resolved if resolved else progress
    if country_label:
        query["paysResidence.label"] = {"$regex": country_label, "$options": "i"}

    demandes = list(db.demandes.find(query).limit(200))
    if not demandes:
        return {"count": 0, "rows": [], "summary": "Aucune demande ne correspond à ces critères."}

    rows = []
    for d in demandes:
        d_user = d.get("user")
        user = _find_user_by_id(d_user) if d_user else None
        rows.append({
            "nom": user.get("nom", "") if user else "",
            "prenom": user.get("prenom", "") if user else "",
            "email": user.get("email", "") if user else "",
            "telephone": user.get("telephone", "") if user else "",
            "candidat_id": str(user["_id"]) if user else "",
            "uniqueId": d.get("uniqueId", ""),
            "status": d.get("status", ""),
            "progress": d.get("progress", ""),
            "pays": d.get("paysResidence", {}).get("label", ""),
            "createdAt": str(d.get("createdAt", "")),
        })

    summary_lines = [f"- {r['prenom']} {r['nom']} ({r['email']}) — statut: {r['status']}" for r in rows[:20]]
    summary = f"{len(rows)} demande(s) trouvée(s) :\n" + "\n".join(summary_lines)
    if len(rows) > 20:
        summary += f"\n... et {len(rows) - 20} autre(s)."

    return {"count": len(rows), "rows": rows, "summary": summary}


def generate_demandes_excel(rows: list[dict], selected_fields: list[str]) -> str:
    """Génère un fichier Excel à partir des lignes filtrées et des champs choisis."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Demandes"

    headers = [DEMANDE_EXPORT_FIELDS.get(f, f) for f in selected_fields]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(f, "") for f in selected_fields])

    filename = f"demandes_export_{uuid.uuid4().hex[:8]}.xlsx"
    filepath = os.path.join(EXPORT_DIR, filename)
    wb.save(filepath)
    return filename


def generic_query(collection: str, field: str, value: str):
    """
    Requête de secours : cherche des documents où un champ correspond à une valeur,
    résolue contre les vraies valeurs présentes en base. Tous les champs sont autorisés
    sauf le mot de passe. Réservé au super admin (contrôlé côté chat.py).
    """
    if collection not in GENERIC_QUERY_ALLOWED_COLLECTIONS:
        return f"Collection non autorisée : '{collection}'. Collections disponibles : {', '.join(GENERIC_QUERY_ALLOWED_COLLECTIONS)}."
    if field in GENERIC_QUERY_EXCLUDED_FIELDS:
        return f"Le champ '{field}' n'est pas accessible."

    resolved_value = resolve_against_real_values(collection, field, value)
    query_value = resolved_value if resolved_value else value

    coll = db[collection]
    count = coll.count_documents({field: query_value})
    if count == 0:
        docs = list(coll.find({field: {"$regex": value, "$options": "i"}}).limit(20))
        count = len(docs)
        if count == 0:
            return f"Aucun résultat pour {field} = '{value}' dans {collection}."
    else:
        docs = list(coll.find({field: query_value}).limit(20))

    lines = [f"{count} résultat(s) trouvé(s) pour {field} = '{query_value}' dans {collection} :"]
    for d in docs[:10]:
        summary_fields = {k: v for k, v in d.items() if k not in ("_id", "_class") and k not in GENERIC_QUERY_EXCLUDED_FIELDS}
        lines.append(f"- {str(summary_fields)[:180]}")
    if count > 10:
        lines.append(f"... et {count - 10} autre(s).")
    return "\n".join(lines)

def get_user_full_profile(user_query: str):
    """Retourne le profil complet d'un utilisateur."""
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    _, full_doc = resolved

    profile = {}
    for key, value in full_doc.items():
        if key in SENSITIVE_USER_FIELDS or key == "_id":
            continue
        profile[key] = _serialize_value(value)

    return profile

def get_orders_for_candidat(user_query: str):
    """Cherche les commandes (orders) liées à un candidat, via son user_id, déjà formatées en texte fiable."""
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    orders = list(db.orders.find({"user_id": ObjectId(user["id"])}).limit(10))
    if not orders:
        return f"Aucune commande trouvée pour {user['prenom']} {user['nom']}."

    blocks = []
    for o in orders:
        montant = o.get("amount")
        devise = o.get("currency", "")
        montant_str = f"{montant} {devise}".strip() if montant is not None else "N/A"

        lines = [
            f"Détails de la commande de {user['prenom']} {user['nom']} :",
            f"- **Identifiant unique** : {o.get('order_id', 'N/A')}",
            f"- **Statut** : {o.get('status', 'N/A')}",
            f"- **Type** : {o.get('type', 'N/A')}",
            f"- **Montant** : {montant_str}",
            f"- **Date** : {o.get('date', 'N/A')}",
            f"- **Heure** : {o.get('time', 'N/A')}",
        ]
        if o.get("description"):
            lines.append(f"- **Description** : {o.get('description')}")
        blocks.append("\n".join(lines))

    return "\n\n".join(blocks)

def get_order_full_profile(order_id: str):
    """Retourne le profil complet d'une commande recherchée par son order_id exact."""
    order = db.orders.find_one({"order_id": order_id})
    if not order:
        return f"Aucune commande trouvée avec l'identifiant '{order_id}'."

    profile = {}
    for key, value in order.items():
        if key in ("_id", "_class", "__v"):
            continue
        profile[key] = _serialize_value(value)
    return profile


def get_demande_full_profile(unique_id: str):
    """
    Retourne TOUS les champs d'une demande recherchée par son identifiant unique exact
    (profil complet, non curaté — à utiliser quand get_demande_by_unique_id ou
    get_all_demandes_with_details ne couvrent pas le champ demandé).
    """
    demande = db.demandes.find_one({"uniqueId": unique_id})
    if not demande:
        return f"Aucune demande trouvée avec l'identifiant '{unique_id}'."

    profile = {}
    for key, value in demande.items():
        if key in ("_id", "_class"):
            continue
        profile[key] = _serialize_value(value)
    return profile


def get_booking_full_profile(unique_id: str):
    """
    Retourne TOUS les champs d'un entretien recherché par son identifiant unique exact
    (profil complet, non curaté — à utiliser quand get_booking_by_unique_id ou
    get_all_demandes_with_details ne couvrent pas le champ demandé).
    """
    booking = db.bookings.find_one({"uniqueId": unique_id})
    if not booking:
        return f"Aucun entretien trouvé avec l'identifiant '{unique_id}'."

    profile = {}
    for key, value in booking.items():
        if key in ("_id", "_class"):
            continue
        profile[key] = _serialize_value(value)
    return profile


def get_documents_for_candidat(user_query: str):
    """
    Retourne la liste des documents disponibles pour un candidat (photo, CIN recto/verso,
    échantillon vocal, enregistrement et rapport d'entretien), sous forme de liens cliquables.
    """
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user_summary, full_user = resolved

    links = []
    for field, label in USER_DOCUMENT_LABELS.items():
        url = _build_file_url(full_user.get(field))
        if url:
            links.append(f"- [{label}]({url})")

    booking = _find_one_by_user_ref(db.bookings, user_summary["id"])
    if booking:
        for field, label in BOOKING_DOCUMENT_LABELS.items():
            url = _build_file_url(booking.get(field))
            if url:
                links.append(f"- [{label}]({url})")

        documents = booking.get("documents") or {}
        for doc_key, doc_value in documents.items():
            if isinstance(doc_value, str) and doc_value:
                url = _build_file_url(doc_value)
                if url:
                    links.append(f"- [{doc_key}]({url})")

    if not links:
        return f"Aucun document n'est disponible pour {user_summary['prenom']} {user_summary['nom']}."

    header = f"Documents disponibles pour {user_summary['prenom']} {user_summary['nom']} :"
    return header + "\n" + "\n".join(links)


def get_user_transactions(user_query: str):
    """Retourne les transactions de solde d'un candidat, déjà formatées en texte fiable."""
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user_summary, full_doc = resolved

    transactions = full_doc.get("balanceTransactions") or []
    if not transactions:
        return f"{user_summary['prenom']} {user_summary['nom']} n'a aucune transaction enregistrée."

    lines = [f"Transactions de {user_summary['prenom']} {user_summary['nom']} ({len(transactions)}) :"]
    for i, t in enumerate(transactions, 1):
        amount = _serialize_value(t.get("amount"))
        created = _serialize_value(t.get("createdAt"))
        lines.append(
            f"{i}. Description : {t.get('description', 'N/A')} | "
            f"Type : {t.get('type', 'N/A')} | "
            f"Montant : {amount} | "
            f"Statut : {t.get('status', 'N/A')} | "
            f"Date : {created}"
        )
    return "\n".join(lines)


def get_user_subscriptions(user_query: str):
    """Retourne les abonnements d'un candidat, déjà formatés en texte fiable."""
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user_summary, full_doc = resolved

    subscriptions = full_doc.get("subscriptions") or []
    if not subscriptions:
        return f"{user_summary['prenom']} {user_summary['nom']} n'a aucun abonnement enregistré."

    lines = [f"Abonnements de {user_summary['prenom']} {user_summary['nom']} ({len(subscriptions)}) :"]
    for i, s in enumerate(subscriptions, 1):
        fields = ", ".join(f"{k}={_serialize_value(v)}" for k, v in s.items() if k not in ("_id",))
        lines.append(f"{i}. {fields}")
    return "\n".join(lines)


def get_user_stb_info(user_query: str):
    """Retourne le statut STB (vérification automatisée) d'un candidat, déjà formaté en texte fiable."""
    resolved = _resolve_full_user(user_query)
    if isinstance(resolved, str):
        return resolved
    user_summary, full_doc = resolved

    stb = full_doc.get("stbInfo")
    if not stb:
        return f"Aucune information STB disponible pour {user_summary['prenom']} {user_summary['nom']}."

    lines = [f"Statut STB de {user_summary['prenom']} {user_summary['nom']} :"]
    for k, v in stb.items():
        if k == "_id":
            continue
        if k == "interviewRecording" and v:
            url = _build_file_url(v) if not str(v).startswith("http") else v
            lines.append(f"- {k} : [Écouter l'enregistrement]({url})")
        elif k == "interviewReport" and v:
            url = _build_file_url(v) if not str(v).startswith("http") else v
            lines.append(f"- {k} : [Voir le rapport]({url})")
        else:
            val = _serialize_value(v)
            lines.append(f"- {k} : {val if val not in (None, '', []) else 'Non renseigné'}")
    return "\n".join(lines)


def chart_demandes_by_status():
    """Données de graphique : répartition des demandes par statut (barres)."""
    return _distribution_chart("demandes", "status", "bar", "Demandes par statut")


def chart_demandes_by_progress():
    """Données de graphique : répartition des demandes par étape de progression (barres)."""
    return _distribution_chart("demandes", "progress", "bar", "Demandes par étape de progression", sort_by_id=True)


def chart_bookings_by_status():
    """Données de graphique : répartition des entretiens par statut (camembert)."""
    return _distribution_chart("bookings", "status", "pie", "Entretiens par statut")


def chart_users_by_role():
    """Données de graphique : répartition des utilisateurs par rôle (camembert)."""
    return _distribution_chart("users", "role", "pie", "Utilisateurs par rôle")


def chart_orders_by_status():
    """Données de graphique : répartition des commandes par statut (barres)."""
    return _distribution_chart("orders", "status", "bar", "Commandes par statut")


def chart_demandes_trend(days: int = 30):
    """Données de graphique : évolution du nombre de demandes créées jour par jour (courbe)."""
    return _trend_chart("demandes", days, f"Évolution des demandes créées ({days} derniers jours)")


def chart_bookings_trend(days: int = 30):
    """Données de graphique : évolution du nombre d'entretiens créés jour par jour (courbe)."""
    return _trend_chart("bookings", days, f"Évolution des entretiens créés ({days} derniers jours)")


def generic_chart(collection: str, field: str, chart_type: str = "bar"):
    """
    Graphique de secours : génère un graphique de répartition (count par valeur) sur
    n'importe quel champ (sauf mot de passe) d'une des 4 collections principales.
    Réservé au super admin (contrôlé côté chat.py).
    """
    if collection not in GENERIC_QUERY_ALLOWED_COLLECTIONS:
        return f"Collection non autorisée : '{collection}'. Collections disponibles : {', '.join(GENERIC_QUERY_ALLOWED_COLLECTIONS)}."
    if field in GENERIC_QUERY_EXCLUDED_FIELDS:
        return f"Le champ '{field}' n'est pas accessible."
    if chart_type not in ("bar", "pie", "line"):
        chart_type = "bar"

    coll = db[collection]
    pipeline = [
        {"$group": {"_id": f"${field}", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$limit": 15},
    ]
    results = list(coll.aggregate(pipeline))

    if not results:
        return f"Aucune donnée disponible pour générer un graphique sur {field} dans {collection}."

    return {
        "type": chart_type,
        "title": f"Répartition de {collection} par {field}",
        "labels": [str(r["_id"]) if r["_id"] is not None else "Non défini" for r in results],
        "values": [r["count"] for r in results],
    }


GENERIC_STAT_OPERATIONS = {"count", "avg", "sum", "min", "max"}


def generic_stat(collection: str, field: str, operation: str = "avg"):
    """
    Statistique de secours : calcule une agrégation numérique (moyenne, somme, min, max, ou compte)
    sur n'importe quel champ (sauf mot de passe) d'une des 4 collections principales.
    Réservé au super admin. À utiliser en dernier recours, pour une statistique numérique
    non couverte par une fonction dédiée (average_interview_score, count_..., etc.).
    """
    if collection not in GENERIC_QUERY_ALLOWED_COLLECTIONS:
        return f"Collection non autorisée : '{collection}'. Collections disponibles : {', '.join(GENERIC_QUERY_ALLOWED_COLLECTIONS)}."
    if field in GENERIC_QUERY_EXCLUDED_FIELDS:
        return f"Le champ '{field}' n'est pas accessible."
    if operation not in GENERIC_STAT_OPERATIONS:
        return f"Opération invalide : '{operation}'. Opérations disponibles : {', '.join(sorted(GENERIC_STAT_OPERATIONS))}."

    coll = db[collection]
    if operation == "count":
        return coll.count_documents({field: {"$exists": True, "$ne": None}})

    pipeline = [
        {"$match": {field: {"$exists": True, "$ne": None}}},
        {"$group": {"_id": None, "result": {f"${operation}": {"$toDouble": f"${field}"}}}},
    ]
    result = list(coll.aggregate(pipeline))
    if not result or result[0]["result"] is None:
        return f"Aucune donnée numérique disponible pour {field} dans {collection}."
    return round(result[0]["result"], 2)


def count_superadmins():
    """Compte le nombre de comptes ayant le statut Super Administrateur."""
    return db.users.count_documents({"isSuperAdmin": True})

def count_bookings_for_demande(demande_unique_id: str):
    """Compte le nombre d'entretiens associés à une demande précise, recherchée par son identifiant unique."""
    demande = db.demandes.find_one({"uniqueId": demande_unique_id})
    if not demande:
        return f"Aucune demande trouvée avec l'identifiant '{demande_unique_id}'."
    demande_id_str = str(demande["_id"])
    count = db.bookings.count_documents({"demande": demande_id_str})
    if count == 0:
        count = db.bookings.count_documents({"demande": demande["_id"]})
    return f"{count} entretien(s) associé(s) à la demande {demande_unique_id}."
def count_demandes_for_candidat(user_query: str):
    """Compte le nombre de demandes associées à un candidat précis, recherché par nom/prénom/email."""
    user = find_user_by_name_or_email(user_query)
    if user and user.get("ambiguous"):
        return format_ambiguous_error(user)
    if not user:
        return None
    demandes = _find_many_by_user_ref(db.demandes, user["id"])
    return f"{user['prenom']} {user['nom']} a {len(demandes)} demande(s)."

def chart_orders_by_type():
    """Données de graphique : répartition des commandes par type (barres)."""
    return _distribution_chart("orders", "type", "bar", "Commandes par type")
def format_ambiguous_error(result: dict) -> str:
    """Formate un message clair quand plusieurs candidats homonymes sont trouvés."""
    lines = ["Plusieurs candidats correspondent à ce nom, merci de préciser avec l'email :"]
    for c in result["candidates"]:
        lines.append(f"- {c['prenom']} {c['nom']} ({c['email']})")
    return "\n".join(lines)
def get_all_demandes_with_details(user_query: str):
    """
    Retourne TOUTES les demandes d'un candidat, chacune avec le détail complet de
    ses entretiens (tous les champs), et les commandes rattachées à chaque entretien.
    """
    user = _resolve_user_or_error(user_query)
    if isinstance(user, str):
        return user

    demandes = _find_many_by_user_ref(db.demandes, user["id"])
    if not demandes:
        return f"{user['prenom']} {user['nom']} n'a aucune demande enregistrée."

    all_orders = list(db.orders.find({"user_id": ObjectId(user["id"])}))
    orders_by_booking = {}
    for o in all_orders:
        bid = o.get("booking_id")
        if bid:
            orders_by_booking.setdefault(bid, []).append(o)

    lines = [f"{user['prenom']} {user['nom']} a {len(demandes)} demande(s) :"]
    linked_order_ids = set()

    for i, d in enumerate(demandes, 1):
        date_str = d.get("createdAt")
        date_fmt = date_str.strftime("%d/%m/%Y") if hasattr(date_str, "strftime") else str(date_str)

        lines.append(f"\n{i}. Demande {d.get('uniqueId', 'N/A')}")
        lines.append(f"   - Statut : {d.get('status', 'N/A')}")
        lines.append(f"   - Progression : {d.get('progress', 'N/A')}")
        lines.append(f"   - Date de création : {date_fmt}")

        demande_id_str = str(d["_id"])
        bookings = list(db.bookings.find({"demande": demande_id_str}))
        if not bookings:
            try:
                bookings = list(db.bookings.find({"demande": d["_id"]}))
            except Exception:
                bookings = []

        if not bookings:
            lines.append("   - Entretiens : aucun")
        else:
            lines.append(f"   - Entretiens ({len(bookings)}) :")
            for b in bookings:
                booking_uid = b.get("uniqueId")
                score = b.get("interviewScore")
                score_str = str(score) if score is not None else "Non noté"
                lines.append(f"       • {booking_uid} :")
                lines.append(f"           - Statut : {b.get('status', 'N/A')}")
                lines.append(f"           - Type : {b.get('appointmentType', 'N/A')}")
                lines.append(f"           - Date : {b.get('appointmentDate', 'N/A')}")
                lines.append(f"           - Heure : {b.get('appointmentTime', 'N/A')}")
                lines.append(f"           - Statut paiement : {b.get('paymentStatus', 'N/A')}")
                lines.append(f"           - Entretien réalisé : {'Oui' if b.get('interviewCompleted') else 'Non'}")
                lines.append(f"           - Score : {score_str}")

                related_orders = orders_by_booking.get(booking_uid, [])
                if related_orders:
                    for o in related_orders:
                        montant = o.get("amount")
                        devise = o.get("currency", "")
                        lines.append(f"           ↳ Commande {o.get('order_id', 'N/A')} : statut {o.get('status', 'N/A')}, montant {montant} {devise}".strip())
                        linked_order_ids.add(o.get("order_id"))
                else:
                    lines.append(f"           ↳ Aucune commande associée")

    remaining = [o for o in all_orders if o.get("order_id") not in linked_order_ids]
    if remaining:
        lines.append(f"\nAutres commandes (non rattachées à un entretien précis) :")
        for o in remaining:
            montant = o.get("amount")
            devise = o.get("currency", "")
            lines.append(f"   • {o.get('order_id', 'N/A')} — statut : {o.get('status', 'N/A')}, type : {o.get('type', 'N/A')}, montant : {montant} {devise}".strip())

    return "\n".join(lines)
def get_bookings_by_date(year: int, month: int, day: int = None):
    """
    Retourne les identifiants des entretiens dont la date de rendez-vous tombe
    dans le mois/année donnés, ou à un jour précis si day est fourni.
    """
    if day:
        start = datetime(year, month, day, tzinfo=timezone.utc)
        end = start + timedelta(days=1)
        period_label = f"{day:02d}/{month:02d}/{year}"
    else:
        start = datetime(year, month, 1, tzinfo=timezone.utc)
        end = datetime(year + 1, 1, 1, tzinfo=timezone.utc) if month == 12 else datetime(year, month + 1, 1, tzinfo=timezone.utc)
        period_label = f"{month:02d}/{year}"

    bookings = list(db.bookings.find(
        {"appointmentDate": {"$gte": start, "$lt": end}},
        {"uniqueId": 1, "status": 1, "appointmentDate": 1, "_id": 0}
    ))

    if not bookings:
        return f"Aucun entretien trouvé pour {period_label}."

    lines = [f"{len(bookings)} entretien(s) le {period_label}:" if day else f"{len(bookings)} entretien(s) en {period_label} :"]
    for b in bookings:
        date_fmt = b["appointmentDate"].strftime("%d/%m/%Y")
        lines.append(f"- {b.get('uniqueId', 'N/A')} ({date_fmt}) — statut : {b.get('status', 'N/A')}")
    return "\n".join(lines)