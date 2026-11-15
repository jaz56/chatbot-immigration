import os
import json
import re
import difflib
from dotenv import load_dotenv
from ollama import Client
from collections import deque
from bson import ObjectId
from datetime import datetime
import logging
from db import (
    update_demande_status_unified,
    get_bookings_by_date,
    db,
    _find_one_by_user_ref,
    _find_many_by_user_ref,
    find_user_by_name_or_email,
    format_ambiguous_error,
    count_demandes_for_candidat,
    count_users,
    count_by_config_category,
    count_demandes_by_status,
    count_demandes_by_progress,
    count_demandes_by_country,
    count_bookings_by_status,
    count_bookings_by_payment_status,
    average_interview_score,
    count_bookings_for_demande,
    count_demandes_total,
    count_bookings_total,
    count_users_by_role,
    count_demandes_created_since,
    count_bookings_created_since,
    create_candidat,
    create_demande,
    create_booking,
    update_demande_progress,
    demandes_breakdown_by_status,
    demandes_breakdown_by_progress,
    demandes_breakdown_by_country,
    bookings_breakdown_by_status,
    bookings_breakdown_by_payment_status,
    users_breakdown_by_role,
    block_admin,
    unblock_admin,
    update_admin_permission,
    count_candidats_pending_verification,
    count_blocked_users,
    count_bookings_this_week,
    reset_candidat_password,
    orders_breakdown_by_status,
    get_demande_by_unique_id,
    get_booking_by_unique_id,
    block_candidat,
    unblock_candidat,
    count_bookings_tomorrow,
    count_users_by_email_verification,
    count_orders_by_status,
    count_orders_by_type,
    get_recent_orders,
    filter_demandes,
    generate_demandes_excel,
    DEMANDE_EXPORT_FIELDS,
    generic_query,
    get_user_full_profile,
    get_orders_for_candidat,
    get_order_full_profile,
    get_demande_full_profile,
    get_booking_full_profile,
    get_documents_for_candidat,
    get_user_transactions,
    get_user_subscriptions,
    get_user_stb_info,
    chart_demandes_by_status,
    chart_demandes_by_progress,
    chart_bookings_by_status,
    chart_users_by_role,
    chart_orders_by_status,
    chart_orders_by_type,
    chart_demandes_trend,
    chart_bookings_trend,
    generic_chart,
    generic_stat,
    count_superadmins,
    get_all_demandes_with_details,
    _normalize_text,
    update_booking_status_unified,
)

load_dotenv()

FUNCTIONS_NEEDING_TOKEN = { "update_demande_status_unified", "update_demande_progress", "update_booking_status_unified", "block_admin", "unblock_admin", "update_admin_permission", "reset_candidat_password", "block_candidat", "unblock_candidat" }

client = Client(
    host="https://ollama.com",
    headers={"Authorization": f"Bearer {os.environ['OLLAMA_API_KEY']}"}
)

CHART_FUNCTIONS = {
    "chart_demandes_by_status", "chart_demandes_by_progress", "chart_bookings_by_status",
    "chart_users_by_role", "chart_orders_by_status", "chart_orders_by_type",
    "chart_demandes_trend", "chart_bookings_trend", "generic_chart",
}

def build_system_prompt() -> str:
    today = datetime.now().strftime("%A %d %B %Y")
    return f"""Tu es l'assistant interne de la plateforme d'administration immigration.

DATE ACTUELLE : nous sommes le {today}. Utilise cette date comme référence pour toute question relative dans le temps (cette semaine, ce mois-ci, demain, récemment, etc.) et pour toute date que tu dois écrire toi-même (par exemple lors de la création d'un entretien).

CONTEXTE PLATEFORME :
Cette plateforme gère le parcours de candidats à l'immigration/mobilité internationale (études, travail, stages).
Les entités principales sont :
- Les CANDIDATS (utilisateurs avec un compte, collection "users", rôle "candidat")
- Les ADMINS (staff interne, collection "users", rôle "admin", avec des permissions par module)
- Les DEMANDES : dossier de candidature, aussi appelé "dossier". Statut (ex: pre_selection) et progression en étapes (ex: STEP_2)
- Les ENTRETIENS/RÉSERVATIONS (appelés "bookings" en base, aussi dits "rendez-vous"), liés à une demande

INTERDICTION ABSOLUE #1 : ne modifie, n'arrondis, ne complète et n'invente JAMAIS un chiffre, un statut ou une date qui ne provient pas exactement d'un résultat d'outil. Si un outil retourne une seule transaction, ne réponds jamais qu'il y en a deux. Recopie fidèlement les champs exacts (type, statut, montant) sans les remplacer par des synonymes ou des valeurs plausibles.

INTERDICTION ABSOLUE #2 : pour TOUTE question portant sur une donnée factuelle d'un candidat, d'une demande, d'un entretien ou d'une commande — même une question de suivi courte type "sa fonction" ou "son activité" — tu dois appeler l'outil correspondant DANS CE MÊME message, sans exception, même si un candidat a été mentionné juste avant. Ne réponds JAMAIS à partir de ce que tu "te souviens" de la conversation : chaque champ factuel doit provenir d'un résultat d'outil de CE message précis. Si tu ne sais pas de quel candidat il s'agit, demande explicitement de préciser le nom plutôt que d'inventer.

Attention : il existe deux notions distinctes d'"entretien". L'entretien STB (vérification automatisée du candidat, avec son propre score/rapport/enregistrement) fait partie du profil utilisateur (get_user_full_profile). L'entretien/rendez-vous réel planifié (booking, avec sa date et son heure) est une entité séparée. Si la question mentionne "STB", "vérification", ou "profil complété", utilise le profil utilisateur. Sinon, si elle mentionne une date, une heure, un type de rendez-vous, utilise le booking.

Si l'admin demande un graphique, une courbe, une visualisation ou un diagramme, utilise TOUJOURS l'outil "chart_..." correspondant plutôt que l'équivalent textuel — jamais l'inverse.

Un candidat peut avoir PLUSIEURS demandes, et chaque demande peut avoir plusieurs entretiens. Pour TOUTE question portant sur un champ précis d'une demande ou d'un entretien (statut, score, date, type, prix, etc.), utilise get_all_demandes_with_details — elle couvre tous les cas, qu'il y ait un seul ou plusieurs résultats, et évite de montrer arbitrairement un seul résultat quand il en existe plusieurs.
Il existe deux scores distincts et à ne jamais confondre : le score STB (users.interviewScore, vérification automatisée initiale) et le score de l'entretien réel planifié (bookings.interviewScore). Sauf si l'admin précise explicitement "STB" ou "vérification automatisée", utilise TOUJOURS par défaut le score de l'entretien (booking).
Quand get_all_demandes_with_details retourne un seul entretien pour un seul candidat, et que l'admin a demandé un champ précis (score, statut, date, etc.), réponds directement avec cette seule valeur, sans réafficher tout le détail. En revanche, si plusieurs demandes ou plusieurs entretiens existent, liste le champ demandé pour CHACUN d'eux, en précisant à quel entretien/demande il correspond.

VOCABULAIRE ÉQUIVALENT (les admins utilisent ces mots de façon interchangeable) :
- "dossier" = "demande"
- "candidat" = utilisateur avec le rôle "candidat"
- "rendez-vous" = "entretien" = "booking" = "réservation"
- "commande" = "order"

LISTES DYNAMIQUES :
Les statuts et étapes valides (statut_demande, statut_entretien, demande_progress_step, etc.) sont définis dans une collection de configuration modifiable par les admins. Ne traduis JAMAIS toi-même une valeur de statut vers l'anglais ou vers une autre langue : transmets le mot exact utilisé par l'admin à l'outil concerné. Si l'outil ne trouve aucune correspondance, transmets son refus avec les valeurs qu'il te donne, sans en inventer d'autres.

TON RÔLE :
Tu réponds aux questions des administrateurs sur les données de la plateforme : comptages, filtres, statistiques, détails d'une entité précise, et tu peux proposer de créer ou modifier des candidats, demandes ou entretiens.
Tu utilises TOUJOURS les outils (function calling) disponibles pour obtenir des informations réelles — tu n'inventes jamais de données ni de chiffres.
Quand tu affiches un statut ou une valeur technique issue d'une fonction, restitue-la fidèlement (traduis le terme en français si besoin, mais ne change JAMAIS sa signification : "pending" veut dire "en attente", pas "confirmé").
Si aucun outil ne correspond à la question posée, dis clairement que tu ne peux pas encore répondre à ce type de question plutôt que d'inventer une réponse.
Si un outil retourne un refus d'accès, explique poliment à l'admin qu'il n'a pas les droits nécessaires sur ce module, sans donner plus de détails sur les données concernées.
N'utilise `generic_query` qu'en tout dernier recours, uniquement si aucune autre fonction ne correspond à la demande de l'admin.
Pour un graphique sur un champ non couvert par les fonctions chart_... dédiées, utilise generic_chart en dernier recours plutôt que de refuser directement.
Avant de refuser une demande de graphique, vérifie toujours si generic_chart peut la couvrir (elle accepte demandes, bookings, users et orders comme collections) — ne présume jamais qu'une collection n'est pas supportée sans l'avoir tenté.
Pour une information précise sur un utilisateur utilise get_user_full_profile. Pour une commande déjà identifiée par son order_id, utilise get_order_full_profile.
Pour savoir si un candidat, une demande ou un entretien a des commandes associées, utilise get_orders_for_candidat.
Si l'admin demande un champ précis d'une demande (uniqueId connu) qui n'est couvert ni par get_demande_by_unique_id ni par get_all_demandes_with_details (ex : note d'éligibilité, visa, hébergement, diplômes, garant), utilise get_demande_full_profile avant de conclure que l'information n'existe pas — elle contient TOUS les champs de la demande. De même, pour un entretien (uniqueId connu), utilise get_booking_full_profile si le champ n'est couvert ni par get_booking_by_unique_id ni par get_all_demandes_with_details (ex : prix, documents joints).
Pour une statistique numérique (moyenne, somme, minimum, maximum) sur un champ non couvert par une fonction dédiée (average_interview_score, count_..., etc.), utilise generic_stat en dernier recours, exactement comme generic_chart pour un graphique non couvert. N'annonce jamais qu'une donnée n'existe pas sans avoir d'abord tenté le fallback générique approprié (get_*_full_profile, generic_query, generic_chart ou generic_stat selon le cas).
Pour "aujourd'hui" ou "demain", calcule toi-même l'année/mois/jour exacts à partir de la date actuelle indiquée en tête de ce prompt, puis appelle get_bookings_by_date avec ces valeurs précises.

ACTIONS DE CRÉATION ET MODIFICATION :
Quand l'admin demande de créer ou modifier un candidat, une demande ou un entretien, appelle l'outil correspondant normalement — le système gère automatiquement la demande de confirmation à sa place.
Pour modifier le statut d'une demande ou d'un entretien, utilise update_demande_status_unified ou update_booking_status_unified. Ces outils acceptent soit un nom de candidat (paramètre user_query), soit un identifiant unique de demande/entretien (paramètre unique_id) — utilise celui que l'admin a fourni, jamais les deux à la fois.
Si la question commence par "combien" ou demande un nombre, utilise TOUJOURS une fonction de comptage (count_..., filter_..., etc.), jamais une fonction de détail.

STYLE :
Réponds toujours en français, de façon concise et professionnelle, comme s'adresserait un collègue de confiance à un autre membre du staff.
Ne te présente pas à chaque message — seulement si on te le demande explicitement.
Quand un outil te retourne déjà un texte formaté en liste (avec des tirets), recopie-le tel quel sans le résumer ni le raccourcir — SAUF pour get_all_demandes_with_details avec un seul résultat, où tu dois extraire uniquement le champ précis demandé par l'admin.
N'ajoute jamais d'unité, de devise ou de symbole (€, $, %, etc.) qui n'est pas explicitement présent dans la donnée retournée par un outil."""
FUNCTION_MODULES = {
    "update_demande_status_unified": "demandes",
    "get_bookings_by_date": "entretiens",
    "count_users": "users",
    "count_by_config_category": None,
    "count_demandes_by_status": "demandes",
    "count_demandes_by_progress": "demandes",
    "count_demandes_by_country": "demandes",
    "count_bookings_by_status": "entretiens",
    "count_bookings_by_payment_status": "entretiens",
    "average_interview_score": "entretiens",
    "count_demandes_total": "demandes",
    "count_bookings_total": "entretiens",
    "count_users_by_role": "users",
    "count_demandes_created_since": "demandes",
    "count_bookings_created_since": "entretiens",
    "get_my_permissions": None,
    "create_candidat": "users",
    "create_demande": "demandes",
    "create_booking": "entretiens",
    "update_demande_progress": "demandes",
    "demandes_breakdown_by_status": "demandes",
    "demandes_breakdown_by_progress": "demandes",
    "demandes_breakdown_by_country": "demandes",
    "bookings_breakdown_by_status": "entretiens",
    "bookings_breakdown_by_payment_status": "entretiens",
    "users_breakdown_by_role": "users",
    "block_admin": "admins",
    "unblock_admin": "admins",
    "update_admin_permission": "admins",
    "count_candidats_pending_verification": "users",
    "count_blocked_users": "users",
    "count_bookings_this_week": "entretiens",
    "reset_candidat_password": "users",
    "orders_breakdown_by_status": "orders",
    "get_demande_by_unique_id": "demandes",
    "get_booking_by_unique_id": "entretiens",
    "update_booking_status_unified": "entretiens",
    "block_candidat": "users",
    "unblock_candidat": "users",
    "count_bookings_tomorrow": "entretiens",
    "count_users_by_email_verification": "users",
    "count_orders_by_status": "orders",
    "count_orders_by_type": "orders",
    "get_recent_orders": "orders",
    "filter_demandes": "demandes",
    "generic_query": None,
    "get_user_full_profile": "users",
    "get_orders_for_candidat": "orders",
    "get_order_full_profile": "orders",
    "get_demande_full_profile": "demandes",
    "get_booking_full_profile": "entretiens",
    "get_documents_for_candidat": "users",
    "get_user_transactions": "users",
    "get_user_subscriptions": "users",
    "get_user_stb_info": "users",
    "generic_chart": None,
    "generic_stat": None,
    "count_bookings_for_demande": "entretiens",
    "count_superadmins": "admins",
    "get_all_demandes_with_details": "demandes",
}

DIRECT_RETURN_FUNCTIONS = {
    "get_documents_for_candidat",
    "get_user_transactions", "get_user_subscriptions", "get_user_stb_info",
    "get_orders_for_candidat","get_bookings_by_date",
}

FUNCTION_ACTIONS = {
    "create_candidat": "create",
    "create_demande": "create",
    "create_booking": "create",
    "update_demande_progress": "update",
    "update_booking_status_unified": "update",
    "reset_candidat_password": "update",
    "get_demande_by_unique_id": "read",
    "get_booking_by_unique_id": "read",
    "block_candidat": "update",
    "unblock_candidat": "update",
    "count_bookings_tomorrow": "read",
    "count_users_by_email_verification": "read",
}

WRITE_FUNCTIONS = {
    "create_candidat", "create_demande", "create_booking",
    "update_demande_progress", "update_booking_status_unified",
    "block_admin", "unblock_admin", "update_admin_permission",
    "reset_candidat_password",
    "block_candidat", "unblock_candidat",
}

SUPERADMIN_ONLY_FUNCTIONS = {
    "block_admin", "unblock_admin", "update_admin_permission",
    "generic_query", "generic_chart", "generic_stat",
}

# Stockages en mémoire, par email admin
pending_actions: dict[str, dict] = {}
last_filter_results: dict[str, dict] = {}
pending_export_field_selection: dict[str, dict] = {}
chart_results: dict[str, dict] = {}
conversation_history: dict[str, deque] = {}
last_mentioned_candidate: dict[str, str] = {}

HISTORY_MAX_MESSAGES = 10

EXPORT_TRIGGER_WORDS = {"exporte", "exporter", "export", "télécharge", "telecharge", "génère un excel", "excel"}
ALL_FIELDS_WORDS = {"tous", "tout", "toutes", "toutes les colonnes", "tous les champs"}
AFFIRMATIVE_WORDS = {"oui", "yes", "confirme", "confirmé", "ok", "okay", "oki", "d'accord", "vas-y", "valide", "je confirme", "carrément", "yep", "yup", "affirmatif"}
NEGATIVE_WORDS = {"non", "no", "annule", "annulé", "stop", "abandonne", "nope", "négatif"}

EXPORT_FILTER_EXTRACT_PROMPT = """Un administrateur demande d'exporter des demandes, en précisant parfois un critère de filtre dans la même phrase.

Il existe deux types de critères bien distincts, à ne pas confondre :
- Le STATUT (status) : l'état global du dossier. Exemples typiques : "en attente", "pré-sélection", "acceptée", "rejetée".
- La PROGRESSION (progress) : l'étape technique interne, presque toujours écrite avec le mot "étape" ou au format STEP_X.

Si la phrase ne contient PAS explicitement le mot "étape" ou un format STEP_X, considère TOUJOURS le critère comme un statut (status), jamais une progression.

Réponds selon ce format :
- status=LA_VALEUR ou progress=LA_VALEUR ou country=LA_VALEUR (séparés par virgule si plusieurs)
- Si aucun critère, réponds exactement : AUCUN
Ne réponds rien d'autre, pas d'explication."""

CONFIRMATION_CLASSIFIER_PROMPT = """Tu classes la réponse d'un administrateur à une question de confirmation oui/non.
Réponds strictement par un seul mot, sans ponctuation ni explication : "oui", "non", ou "incertain" si le message ne répond clairement ni oui ni non."""

EXPORT_CLASSIFIER_PROMPT = """Tu détermines si un message demande d'exporter ou de télécharger des données sous forme de fichier Excel.
Réponds strictement par un seul mot : "oui" ou "non", sans explication."""

FIELD_CLASSIFIER_PROMPT_TEMPLATE = """Voici la liste des champs disponibles pour un export, avec leur clé technique et leur nom affiché :
{field_list}

L'administrateur va décrire les champs qu'il veut inclure, en langage naturel, éventuellement de façon imprécise, familière ou avec des synonymes.
Réponds UNIQUEMENT avec les clés techniques correspondantes, séparées par des virgules, sans aucun autre texte.
Si l'administrateur semble vouloir tout inclure, réponds exactement : TOUS
Si tu ne reconnais aucun champ, réponds exactement : AUCUN"""

DEMANDE_FIELD_ALIASES = {
    "fonction": "fonction",
    "dernier poste": "dernierPosteOccupe",
    "poste occupé": "dernierPosteOccupe",
    "activité": "activite",
    "poste souhaité": "posteSouhaite",
    "niveau d'étude": "niveauEtude",
    "niveau d'études": "niveauEtude",
    "années d'étude": "nombreAnneesEtude",
    "années d'études": "nombreAnneesEtude",
    "mois de stage": "nombreMoisStage",
    "statut actuel du candidat": "condidatStatutActuel",
    "années d'expérience": "nombreAnneesExperience",
    "destination": "journeesDestination",
}

FAKE_CALL_PATTERN = re.compile(r"^\s*call\s*:\s*(\w+)\s*\{(.*)\}\s*$", re.DOTALL)

available_functions = {
    "update_demande_status_unified": update_demande_status_unified,
    "get_bookings_by_date": get_bookings_by_date,
    "count_users": count_users,
    "count_by_config_category": count_by_config_category,
    "count_demandes_by_status": count_demandes_by_status,
    "count_demandes_by_progress": count_demandes_by_progress,
    "count_demandes_by_country": count_demandes_by_country,
    "count_bookings_by_status": count_bookings_by_status,
    "count_bookings_by_payment_status": count_bookings_by_payment_status,
    "average_interview_score": average_interview_score,
    "count_demandes_total": count_demandes_total,
    "count_bookings_total": count_bookings_total,
    "count_users_by_role": count_users_by_role,
    "count_demandes_created_since": count_demandes_created_since,
    "count_bookings_created_since": count_bookings_created_since,
    "create_candidat": create_candidat,
    "create_demande": create_demande,
    "create_booking": create_booking,
    "update_demande_progress": update_demande_progress,
    "demandes_breakdown_by_status": demandes_breakdown_by_status,
    "demandes_breakdown_by_progress": demandes_breakdown_by_progress,
    "demandes_breakdown_by_country": demandes_breakdown_by_country,
    "bookings_breakdown_by_status": bookings_breakdown_by_status,
    "bookings_breakdown_by_payment_status": bookings_breakdown_by_payment_status,
    "users_breakdown_by_role": users_breakdown_by_role,
    "block_admin": block_admin,
    "unblock_admin": unblock_admin,
    "update_admin_permission": update_admin_permission,
    "count_candidats_pending_verification": count_candidats_pending_verification,
    "count_blocked_users": count_blocked_users,
    "count_bookings_this_week": count_bookings_this_week,
    "reset_candidat_password": reset_candidat_password,
    "orders_breakdown_by_status": orders_breakdown_by_status,
    "get_demande_by_unique_id": get_demande_by_unique_id,
    "get_booking_by_unique_id": get_booking_by_unique_id,
    "update_booking_status_unified": update_booking_status_unified,
    "block_candidat": block_candidat,
    "unblock_candidat": unblock_candidat,
    "count_bookings_tomorrow": count_bookings_tomorrow,
    "count_users_by_email_verification": count_users_by_email_verification,
    "count_orders_by_status": count_orders_by_status,
    "count_orders_by_type": count_orders_by_type,
    "get_recent_orders": get_recent_orders,
    "filter_demandes": filter_demandes,
    "generic_query": generic_query,
    "get_user_full_profile": get_user_full_profile,
    "get_orders_for_candidat": get_orders_for_candidat,
    "get_order_full_profile": get_order_full_profile,
    "get_demande_full_profile": get_demande_full_profile,
    "get_booking_full_profile": get_booking_full_profile,
    "get_documents_for_candidat": get_documents_for_candidat,
    "get_user_transactions": get_user_transactions,
    "get_user_subscriptions": get_user_subscriptions,
    "get_user_stb_info": get_user_stb_info,
    "chart_demandes_by_status": chart_demandes_by_status,
    "chart_demandes_by_progress": chart_demandes_by_progress,
    "chart_bookings_by_status": chart_bookings_by_status,
    "chart_users_by_role": chart_users_by_role,
    "chart_orders_by_status": chart_orders_by_status,
    "chart_orders_by_type": chart_orders_by_type,
    "chart_demandes_trend": chart_demandes_trend,
    "chart_bookings_trend": chart_bookings_trend,
    "generic_chart": generic_chart,
    "generic_stat": generic_stat,
    "count_bookings_for_demande": count_bookings_for_demande,
    "count_superadmins": count_superadmins,
    "get_all_demandes_with_details": get_all_demandes_with_details,
}

tools = [
    {"type": "function", "function": {"name": "count_users", "description": "Compte le nombre total d'utilisateurs enregistrés dans la plateforme", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_by_config_category", "description": "Compte le nombre de valeurs actives pour une catégorie de configuration donnée", "parameters": {"type": "object", "properties": {"category": {"type": "string"}}, "required": ["category"]}}},
    {"type": "function", "function": {"name": "count_demandes_by_status", "description": "Compte le nombre de demandes ayant un statut donné", "parameters": {"type": "object", "properties": {"status": {"type": "string"}}, "required": ["status"]}}},
    {"type": "function", "function": {"name": "count_demandes_by_progress", "description": "Compte le nombre de demandes à une étape de progression donnée", "parameters": {"type": "object", "properties": {"progress": {"type": "string"}}, "required": ["progress"]}}},
    {"type": "function", "function": {"name": "count_demandes_by_country", "description": "Compte le nombre de demandes pour un pays de résidence donné", "parameters": {"type": "object", "properties": {"country_label": {"type": "string"}}, "required": ["country_label"]}}},
    {"type": "function", "function": {"name": "count_bookings_by_status", "description": "Compte le nombre d'entretiens ayant un statut donné", "parameters": {"type": "object", "properties": {"status": {"type": "string"}}, "required": ["status"]}}},
    {"type": "function", "function": {"name": "count_bookings_by_payment_status", "description": "Compte le nombre de réservations selon leur statut de paiement", "parameters": {"type": "object", "properties": {"payment_status": {"type": "string"}}, "required": ["payment_status"]}}},
    {"type": "function", "function": {"name": "average_interview_score", "description": "Calcule la note moyenne des entretiens complétés", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_demandes_total", "description": "Compte le nombre total de demandes", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_bookings_total", "description": "Compte le nombre total d'entretiens", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_users_by_role", "description": "Compte le nombre d'utilisateurs ayant un rôle donné", "parameters": {"type": "object", "properties": {"role": {"type": "string"}}, "required": ["role"]}}},
    {"type": "function", "function": {"name": "count_demandes_created_since", "description": "Compte le nombre de demandes créées au cours des N derniers jours", "parameters": {"type": "object", "properties": {"days": {"type": "integer"}}, "required": ["days"]}}},
    {"type": "function", "function": {"name": "count_bookings_created_since", "description": "Compte le nombre d'entretiens créés au cours des N derniers jours", "parameters": {"type": "object", "properties": {"days": {"type": "integer"}}, "required": ["days"]}}},
    {"type": "function", "function": {"name": "get_my_permissions", "description": "Retourne les permissions de l'utilisateur actuellement connecté", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "create_candidat", "description": "Crée un nouveau candidat, avec un mot de passe temporaire généré automatiquement", "parameters": {"type": "object", "properties": {"email": {"type": "string"}, "nom": {"type": "string"}, "prenom": {"type": "string"}}, "required": ["email", "nom", "prenom"]}}},
    {"type": "function", "function": {"name": "create_demande", "description": "Crée une nouvelle demande vide pour un candidat existant", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "create_booking", "description": "Crée un nouvel entretien pour un candidat ayant déjà une demande", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}, "appointment_date": {"type": "string"}, "appointment_time": {"type": "string"}, "appointment_type": {"type": "string"}}, "required": ["user_query", "appointment_date", "appointment_time"]}}},
    {"type": "function", "function": {"name": "update_demande_progress", "description": "Modifie l'étape de progression de la demande d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}, "new_progress": {"type": "string"}}, "required": ["user_query", "new_progress"]}}},
    {"type": "function", "function": {"name": "demandes_breakdown_by_status", "description": "Répartition du nombre de demandes par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "demandes_breakdown_by_progress", "description": "Répartition du nombre de demandes par étape de progression", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "demandes_breakdown_by_country", "description": "Répartition du nombre de demandes par pays de résidence", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "bookings_breakdown_by_status", "description": "Répartition du nombre d'entretiens par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "bookings_breakdown_by_payment_status", "description": "Répartition du nombre de réservations par statut de paiement", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "users_breakdown_by_role", "description": "Répartition du nombre d'utilisateurs par rôle", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "block_admin", "description": "Bloque le compte d'un administrateur (réservé au super admin)", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}, "reason": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "unblock_admin", "description": "Débloque le compte d'un administrateur (réservé au super admin)", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "update_admin_permission", "description": "Modifie un droit précis sur un module pour un administrateur (réservé au super admin)", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}, "module": {"type": "string"}, "action": {"type": "string"}, "value": {"type": "boolean"}}, "required": ["user_query", "module", "action", "value"]}}},
    {"type": "function", "function": {"name": "count_candidats_pending_verification", "description": "Compte les candidats en attente de vérification", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_blocked_users", "description": "Compte le nombre total de comptes bloqués", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_bookings_this_week", "description": "Compte les entretiens prévus dans les 7 prochains jours", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "reset_candidat_password", "description": "Réinitialise le mot de passe d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "orders_breakdown_by_status", "description": "Répartition du nombre de commandes par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "get_demande_by_unique_id", "description": "Détails d'une demande par son identifiant unique exact", "parameters": {"type": "object", "properties": {"unique_id": {"type": "string"}}, "required": ["unique_id"]}}},
    {"type": "function", "function": {"name": "get_booking_by_unique_id", "description": "Détails d'un entretien par son identifiant unique exact", "parameters": {"type": "object", "properties": {"unique_id": {"type": "string"}}, "required": ["unique_id"]}}},
    {"type": "function", "function": {"name": "update_booking_status_unified", "description": "Modifie le statut d'un entretien par son identifiant unique exact", "parameters": {"type": "object", "properties": {"unique_id": {"type": "string"}, "new_status": {"type": "string"}}, "required": ["unique_id", "new_status"]}}},
    {"type": "function", "function": {"name": "block_candidat", "description": "Bloque le compte d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}, "reason": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "unblock_candidat", "description": "Débloque le compte d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "count_bookings_tomorrow", "description": "Compte les entretiens prévus demain", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_users_by_email_verification", "description": "Compte les utilisateurs ayant vérifié ou non leur email", "parameters": {"type": "object", "properties": {"verified": {"type": "boolean"}}, "required": ["verified"]}}},
    {"type": "function", "function": {"name": "count_orders_by_status", "description": "Compte les commandes ayant un statut donné", "parameters": {"type": "object", "properties": {"status": {"type": "string"}}, "required": ["status"]}}},
    {"type": "function", "function": {"name": "count_orders_by_type", "description": "Compte les commandes d'un type donné", "parameters": {"type": "object", "properties": {"order_type": {"type": "string"}}, "required": ["order_type"]}}},
    {"type": "function", "function": {"name": "get_recent_orders", "description": "Retourne les commandes les plus récemment créées", "parameters": {"type": "object", "properties": {"limit": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {"name": "filter_demandes", "description": "Cherche des demandes par statut/progression/pays, garde le résultat en mémoire pour export", "parameters": {"type": "object", "properties": {"status": {"type": "string"}, "progress": {"type": "string"}, "country_label": {"type": "string"}}, "required": []}}},
    {"type": "function", "function": {"name": "generic_query", "description": "Requête de secours sur n'importe quel champ (sauf mot de passe) d'une collection (demandes, bookings, users, orders). Réservé au super admin. À utiliser en dernier recours.", "parameters": {"type": "object", "properties": {"collection": {"type": "string"}, "field": {"type": "string"}, "value": {"type": "string"}}, "required": ["collection", "field", "value"]}}},
    {"type": "function", "function": {"name": "get_user_full_profile", "description": "Profil complet d'un utilisateur : téléphone, adresse, solde, statut STB, abonnements, transactions", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "get_orders_for_candidat", "description": "Cherche les commandes liées à un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "get_order_full_profile", "description": "Profil complet d'une commande par son order_id exact", "parameters": {"type": "object", "properties": {"order_id": {"type": "string"}}, "required": ["order_id"]}}},
    {"type": "function", "function": {"name": "get_demande_full_profile", "description": "Profil COMPLET d'une demande par son identifiant unique exact (tous les champs, y compris ceux non couverts par get_demande_by_unique_id ou get_all_demandes_with_details : note d'éligibilité, visa, hébergement, diplômes, garant, etc.)", "parameters": {"type": "object", "properties": {"unique_id": {"type": "string"}}, "required": ["unique_id"]}}},
    {"type": "function", "function": {"name": "get_booking_full_profile", "description": "Profil COMPLET d'un entretien par son identifiant unique exact (tous les champs, y compris ceux non couverts par get_booking_by_unique_id ou get_all_demandes_with_details : prix, documents joints, etc.)", "parameters": {"type": "object", "properties": {"unique_id": {"type": "string"}}, "required": ["unique_id"]}}},
    {"type": "function", "function": {"name": "get_documents_for_candidat", "description": "Tous les documents d'un candidat (photo, CIN, vocal, enregistrement, rapport) en liens cliquables", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "get_user_transactions", "description": "Transactions financières d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "get_user_subscriptions", "description": "Abonnements d'un candidat", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "get_user_stb_info", "description": "Statut STB complet (vérification automatisée) d'un candidat : score STB, statut vérification, rapport, enregistrement STB", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
    {"type": "function", "function": {"name": "chart_demandes_by_status", "description": "Graphique en barres de la répartition des demandes par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_demandes_by_progress", "description": "Graphique en barres de la répartition des demandes par étape de progression", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_bookings_by_status", "description": "Graphique camembert de la répartition des entretiens par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_users_by_role", "description": "Graphique camembert de la répartition des utilisateurs par rôle", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_orders_by_status", "description": "Graphique en barres de la répartition des commandes par statut", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_orders_by_type", "description": "Graphique en barres de la répartition des commandes par type", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "chart_demandes_trend", "description": "Courbe d'évolution du nombre de demandes créées, jour par jour", "parameters": {"type": "object", "properties": {"days": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {"name": "chart_bookings_trend", "description": "Courbe d'évolution du nombre d'entretiens créés, jour par jour", "parameters": {"type": "object", "properties": {"days": {"type": "integer"}}, "required": []}}},
    {"type": "function", "function": {"name": "generic_chart", "description": "Graphique de secours sur n'importe quel champ (sauf mot de passe) de demandes/bookings/users/orders. Réservé au super admin. À utiliser en dernier recours.", "parameters": {"type": "object", "properties": {"collection": {"type": "string"}, "field": {"type": "string"}, "chart_type": {"type": "string"}}, "required": ["collection", "field"]}}},
    {"type": "function", "function": {"name": "generic_stat", "description": "Statistique numérique de secours (moyenne, somme, min, max, ou compte) sur n'importe quel champ numérique (sauf mot de passe) de demandes/bookings/users/orders. Réservé au super admin. À utiliser en dernier recours, pour une statistique non couverte par une fonction dédiée.", "parameters": {"type": "object", "properties": {"collection": {"type": "string"}, "field": {"type": "string"}, "operation": {"type": "string", "description": "count, avg, sum, min ou max"}}, "required": ["collection", "field"]}}},
    {"type": "function", "function": {"name": "count_superadmins", "description": "Compte les comptes ayant le statut Super Administrateur", "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "count_bookings_for_demande", "description": "Compte les entretiens associés à une demande précise, par son identifiant unique", "parameters": {"type": "object", "properties": {"demande_unique_id": {"type": "string"}}, "required": ["demande_unique_id"]}}},
    {"type": "function", "function": {"name": "get_all_demandes_with_details", "description": "Retourne TOUTES les demandes d'un candidat, chacune avec ses entretiens complets et les commandes rattachées. À utiliser en PRIORITÉ pour toute question sur un champ précis d'une demande ou d'un entretien (statut, score, date, type, prix, niveau d'étude, fonction, etc.), car un candidat peut avoir plusieurs demandes/entretiens.", "parameters": {"type": "object", "properties": {"user_query": {"type": "string"}}, "required": ["user_query"]}}},
{"type": "function", "function": {"name": "get_bookings_by_date", "description": "Retourne les identifiants et statuts des entretiens à une date précise (si day fourni) ou sur tout un mois (si day omis). Utilise cet outil pour toute demande d'entretiens 'd'aujourd'hui', 'de demain', 'du 15 septembre', ou 'de tel mois'.", "parameters": {"type": "object", "properties": {"year": {"type": "integer", "description": "L'année, ex: 2026"}, "month": {"type": "integer", "description": "Le mois, de 1 à 12"}, "day": {"type": "integer", "description": "Le jour précis (optionnel) — omets ce paramètre pour obtenir tout le mois"}}, "required": ["year", "month"]}}},    {"type": "function", "function": {"name": "update_demande_status_unified", "description": "Modifie le statut d'une demande, identifiée soit par candidat soit par uniqueId", "parameters": {"type": "object", "properties": {"new_status": {"type": "string"}, "user_query": {"type": "string"}, "unique_id": {"type": "string"}}, "required": ["new_status"]}}},

]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("chatbot")

def get_history(email: str) -> deque:
    if email not in conversation_history:
        conversation_history[email] = deque(maxlen=HISTORY_MAX_MESSAGES)
    return conversation_history[email]


def add_to_history(email: str, role: str, content: str):
    get_history(email).append({"role": role, "content": content})


def llm_classify(instruction: str, user_text: str, history: list[dict] = None) -> str:
    """Appelle le modèle pour classifier un texte selon une instruction stricte, avec contexte optionnel."""
    try:
        messages = [{"role": "system", "content": instruction}]
        if history:
            messages.extend(history)
        messages.append({"role": "user", "content": user_text})
        response = client.chat(model="gemma4:31b-cloud", messages=messages)
        return response.message.content.strip().lower().strip(".!?\"'")
    except Exception:
        return "incertain"


def classify_confirmation(text: str, history: list[dict] = None) -> str:
    normalized = text.strip().lower().strip(".!?")
    if normalized in AFFIRMATIVE_WORDS or any(w in normalized.split() for w in AFFIRMATIVE_WORDS):
        return "oui"
    if normalized in NEGATIVE_WORDS or any(w in normalized.split() for w in NEGATIVE_WORDS):
        return "non"
    raw = llm_classify(CONFIRMATION_CLASSIFIER_PROMPT, text, history)
    has_oui = "oui" in raw
    has_non = "non" in raw
    if has_oui and not has_non:
        return "oui"
    if has_non and not has_oui:
        return "non"
    return "incertain"


def describe_action(func_name: str, func_args: dict) -> str:
    if func_name == "create_candidat":
        return f"créer un nouveau candidat : {func_args.get('prenom')} {func_args.get('nom')} ({func_args.get('email')})"
    if func_name == "create_demande":
        return f"créer une nouvelle demande pour {func_args.get('user_query')}"
    if func_name == "create_booking":
        return f"créer un entretien pour {func_args.get('user_query')} le {func_args.get('appointment_date')} à {func_args.get('appointment_time')}"
    if func_name == "update_demande_status_unified":
        cible = func_args.get('user_query') or func_args.get('unique_id')
        return f"changer le statut de la demande de {cible} vers « {func_args.get('new_status')} »"
    if func_name == "update_demande_progress":
        return f"changer la progression de la demande de {func_args.get('user_query')} vers « {func_args.get('new_progress')} »"
    if func_name == "update_booking_status_unified":
        cible = func_args.get('user_query') or func_args.get('unique_id')
        return f"changer le statut de l'entretien de {cible} vers « {func_args.get('new_status')} »"
    if func_name == "block_admin":
        return f"bloquer le compte de {func_args.get('user_query')}"
    if func_name == "unblock_admin":
        return f"débloquer le compte de {func_args.get('user_query')}"
    if func_name == "update_admin_permission":
        etat = "accorder" if func_args.get('value') else "retirer"
        return f"{etat} le droit « {func_args.get('action')} » sur le module « {func_args.get('module')} » pour {func_args.get('user_query')}"
    if func_name == "reset_candidat_password":
        return f"réinitialiser le mot de passe de {func_args.get('user_query')}"
    if func_name == "block_candidat":
        return f"bloquer le compte candidat de {func_args.get('user_query')}"
    if func_name == "unblock_candidat":
        return f"débloquer le compte candidat de {func_args.get('user_query')}"
    if func_name == "update_order_status":
        return f"changer le statut de la commande {func_args.get('order_id')} vers « {func_args.get('new_status')} »"
    return f"exécuter l'action {func_name}"

def has_permission(user: dict, func_name: str) -> bool:
    if func_name in SUPERADMIN_ONLY_FUNCTIONS:
        return user.get("isSuperAdmin", False)
    if user.get("isSuperAdmin"):
        return True
    module = FUNCTION_MODULES.get(func_name)
    if module is None:
        return True
    action = FUNCTION_ACTIONS.get(func_name, "read")
    perms = user.get("permissions", {})
    module_perms = perms.get(module)
    if module_perms is None:
        return False
    return module_perms.get(action, False)


def looks_like_fake_tool_call(text: str) -> bool:
    """Détecte si le modèle a halluciné un format d'appel d'outil en texte au lieu d'un vrai tool_call structuré."""
    return bool(re.match(r"^\s*call\s*:\s*\w+\s*\{", text.strip()))


def looks_truncated(text: str) -> bool:
    """Heuristique : réponse vide, tronquée, ou hallucination de type 'call:...'."""
    stripped = text.strip()
    if not stripped:
        return True
    if looks_like_fake_tool_call(stripped):
        return True
    if len(stripped) < 15:
        return False
    return stripped[-1] not in ".!?»\"':)"


def try_execute_fake_tool_call(text: str, user: dict):
    """Parse et exécute directement un pseudo-appel d'outil que le modèle aurait écrit en texte brut."""
    match = FAKE_CALL_PATTERN.match(text.strip())
    if not match:
        return None
    func_name, args_str = match.group(1), match.group(2)
    if func_name not in available_functions:
        logger.error(f"❌ Fonction inconnue dans le faux appel : {func_name}")
        return None

    func_args = {}
    for part in args_str.split(","):
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        func_args[key.strip()] = value.strip()

    if not has_permission(user, func_name):
        return "Accès refusé : vous n'avez pas les droits nécessaires pour effectuer cette action."

    try:
        if func_name in FUNCTIONS_NEEDING_TOKEN:
            result = available_functions[func_name](**func_args, token=user.get("token"))
        else:
            result = available_functions[func_name](**func_args)
    except TypeError:
        return None

    if isinstance(result, dict):
        return "\n".join(f"{k} : {v}" for k, v in result.items())
    return str(result)


def is_export_request(text: str, history: list[dict] = None) -> bool:
    normalized = text.strip().lower()
    if any(w in normalized for w in EXPORT_TRIGGER_WORDS):
        return True
    words = normalized.split()
    for word in words:
        if difflib.get_close_matches(word, EXPORT_TRIGGER_WORDS, n=1, cutoff=0.75):
            return True
    return llm_classify(EXPORT_CLASSIFIER_PROMPT, text, history) == "oui"


def extract_filter_from_export_request(text: str) -> dict:
    raw = llm_classify(EXPORT_FILTER_EXTRACT_PROMPT, text)
    if raw == "aucun" or "=" not in raw:
        return {}
    result = {}
    for part in raw.split(","):
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        key, value = key.strip(), value.strip()
        if key == "status":
            result["status"] = value
        elif key == "progress":
            result["progress"] = value
        elif key == "country":
            result["country_label"] = value
    return result


def _direct_match_fields(text: str) -> list[str]:
    normalized = text.strip().lower()
    selected = []
    for key, label in DEMANDE_EXPORT_FIELDS.items():
        if label.lower() in normalized or key.lower() in normalized:
            selected.append(key)
    return selected


def parse_selected_fields(text: str, history: list[dict] = None) -> list[str]:
    normalized = text.strip().lower()
    if any(w in normalized for w in ALL_FIELDS_WORDS):
        return list(DEMANDE_EXPORT_FIELDS.keys())

    field_list = "\n".join(f"- {k} : {v}" for k, v in DEMANDE_EXPORT_FIELDS.items())
    prompt = FIELD_CLASSIFIER_PROMPT_TEMPLATE.format(field_list=field_list)
    raw = llm_classify(prompt, text, history)

    if raw == "tous":
        return list(DEMANDE_EXPORT_FIELDS.keys())
    if raw == "aucun":
        return []
    candidates = [c.strip() for c in raw.split(",")]
    matched = [c for c in candidates if c in DEMANDE_EXPORT_FIELDS]
    return matched if matched else _direct_match_fields(text)


def extract_candidate_name(question: str) -> str | None:
    """Extrait un nom probable en fin de phrase, précédé de 'de'/'à'/'pour', pour mémoriser le dernier candidat évoqué."""
    match = re.search(
        r"(?:de|à|pour)\s+([a-zàâäéèêëïîôöùûüç]+(?:[\s\-][a-zàâäéèêëïîôöùûüç]+)*)\s*\??$",
        question.strip().lower()
    )
    return match.group(1).strip() if match else None

def clear_session(email: str):
    """Nettoie tout l'état en mémoire d'un admin : confirmations en attente, historique, dernier candidat, filtres, graphiques."""
    pending_actions.pop(email, None)
    last_filter_results.pop(email, None)
    pending_export_field_selection.pop(email, None)
    chart_results.pop(email, None)
    conversation_history.pop(email, None)
    last_mentioned_candidate.pop(email, None)
    logger.info(f"🧹 Session nettoyée pour {email}")

def try_direct_field_routing(question: str, user: dict, last_candidate: str | None) -> str | None:
    try:
        if not last_candidate:
            return None
        normalized = _normalize_text(question)
        for phrase, field_key in DEMANDE_FIELD_ALIASES.items():
            if _normalize_text(phrase) in normalized:
                if not has_permission(user, "get_all_demandes_with_details"):
                    return None
                profile = get_all_demandes_with_details(last_candidate)
                if isinstance(profile, str):
                    return profile
                value = profile.get(field_key)
                if value in (None, ""):
                    return f"Le champ '{phrase}' n'est pas renseigné pour ce candidat."
                return f"{phrase.capitalize()} : {value}"
        return None
    except Exception as e:
        logger.error(f" Erreur dans try_direct_field_routing : {e}")
        return None


def _lookup_named_user(name: str):
    """Cherche un utilisateur par nom pour les routages directs. Retourne le dict utilisateur,
    un message d'erreur (str) si le nom est ambigu, ou None si introuvable."""
    found_user = find_user_by_name_or_email(name)
    if not found_user:
        return None
    if found_user.get("ambiguous"):
        return format_ambiguous_error(found_user)
    return found_user


def try_direct_count_routing(question: str, user: dict) -> str | None:
    """Court-circuite le choix d'outil du modèle pour les questions de comptage évidentes liées à un candidat nommé."""
    normalized = question.strip().lower()
    if "combien" not in normalized:
        return None

    match = re.search(
        r"(?:de|à|pour|admet[- ]?(?:t[- ]?il)?)\s+([a-zàâäéèêëïîôöùûüç]+(?:[\s\-][a-zàâäéèêëïîôöùûüç]+)*)\s*\??$",
        normalized
    )
    name = match.group(1).strip() if match else None
    if not name or len(name) < 3:
        return None

    if "entretien" in normalized or "booking" in normalized or "rendez-vous" in normalized:
        found_user = _lookup_named_user(name)
        if found_user is None:
            return None
        if isinstance(found_user, str):
            return found_user

        demandes = _find_many_by_user_ref(db.demandes, found_user["id"])
        if not demandes:
            return f"Aucune demande trouvée pour {found_user['prenom']} {found_user['nom']}, donc aucun entretien associé possible."

        if len(demandes) == 1:
            return count_bookings_for_demande(demandes[0].get("uniqueId"))

        lines = [f"{found_user['prenom']} {found_user['nom']} a {len(demandes)} demandes, avec les entretiens suivants :"]
        for d in demandes:
            result = count_bookings_for_demande(d.get("uniqueId"))
            lines.append(f"- Demande {d.get('uniqueId')} ({d.get('status')}) : {result}")
        return "\n".join(lines)

    if "order" in normalized or "commande" in normalized:
        found_user = _lookup_named_user(name)
        if found_user is None:
            return None
        if isinstance(found_user, str):
            return found_user
        orders = list(db.orders.find({"user_id": ObjectId(found_user["id"])}))
        return f"{found_user['prenom']} {found_user['nom']} a {len(orders)} commande(s)."

    if "demande" in normalized or "dossier" in normalized:
        return count_demandes_for_candidat(name)

    return None


def ask(question: str, user: dict):
    email = user.get("email", "")
    chart_results.pop(email, None)
    history = list(get_history(email))
    answer = _ask_core(question, user, email, history)
    chart = chart_results.pop(email, None)
    add_to_history(email, "user", question)
    add_to_history(email, "assistant", answer)
    return {"text": answer, "chart": chart}


def _ask_core(question: str, user: dict, email: str, history: list[dict]) -> str:
    logger.info(f"🟢 _ask_core démarré pour : {question}")
    name = extract_candidate_name(question)
    if name:
        last_mentioned_candidate[email] = name

    field_direct = try_direct_field_routing(question, user, last_mentioned_candidate.get(email))
    if field_direct is not None:
        logger.info(f" Routage direct (champ) déclenché")
        return field_direct

    direct = try_direct_count_routing(question, user)
    if direct is not None:
        logger.info(f" Routage direct (comptage) déclenché")
        return direct

    # 0a. L'admin est-il en train de choisir ses champs d'export ?
    if email in pending_export_field_selection:
        pending = pending_export_field_selection.pop(email)
        fields = parse_selected_fields(question, history)
        if not fields:
            return f"Je n'ai reconnu aucun champ valide. Champs disponibles : {', '.join(DEMANDE_EXPORT_FIELDS.values())}. Réessayez, ou dites « tous les champs »."
        filename = generate_demandes_excel(pending["rows"], fields)
        return f"Export prêt : [Télécharger le fichier Excel](/api/v1/download/{filename})"

    # 0b. L'admin demande-t-il un export ?
    if is_export_request(question, history):
        if email not in last_filter_results:
            filter_args = extract_filter_from_export_request(question)
            if filter_args:
                result = filter_demandes(**filter_args)
                if result["rows"]:
                    last_filter_results[email] = {"rows": result["rows"]}
                else:
                    return "Aucune demande ne correspond à ces critères, rien à exporter."
            else:
                return "Je n'ai pas encore de résultats à exporter. Demandez-moi d'abord de filtrer des demandes, puis redemandez l'export."

        pending_export_field_selection[email] = last_filter_results[email]

        fields = parse_selected_fields(question, history)
        if fields:
            pending_export_field_selection.pop(email)
            filename = generate_demandes_excel(last_filter_results[email]["rows"], fields)
            return f"Export prêt : [Télécharger le fichier Excel](/api/v1/download/{filename})"

        return f"Quels champs voulez-vous inclure dans l'export ? Champs disponibles : {', '.join(DEMANDE_EXPORT_FIELDS.values())}. Répondez avec les champs souhaités, ou dites « tous les champs »."

    # 1. Une action est-elle en attente de confirmation pour cet admin ?
        # 1. Une action est-elle en attente de confirmation pour cet admin ?
    if email in pending_actions:
        pending = pending_actions[email]
        verdict = classify_confirmation(question, history)

        if verdict == "oui":
            func_name = pending["function"]
            func_args = pending["args"]
            del pending_actions[email]
            if not has_permission(user, func_name):
                return "Accès refusé : vous n'avez pas les droits nécessaires pour effectuer cette action."
            if func_name in FUNCTIONS_NEEDING_TOKEN:
                result = available_functions[func_name](**func_args, token=user.get("token"))
            else:
                result = available_functions[func_name](**func_args)
            message = result.get("message", str(result)) if isinstance(result, dict) else str(result)
            if isinstance(result, str) and any(w in result.lower() for w in ["invalide", "aucun", "existe déjà", "n'est pas", "refusé"]):
                return message
            return f"C'est fait : {message}"
        elif verdict == "non":
            del pending_actions[email]
            return "Action annulée, aucune modification n'a été effectuée."
        else:
            # Ni oui ni non détecté : c'est probablement une nouvelle demande, pas une réponse.
            # On abandonne silencieusement l'ancienne confirmation et on traite le nouveau message normalement.
            del pending_actions[email]
            logger.info(f"Confirmation abandonnée (nouvelle demande détectée) : {question}")
            # ne pas return ici — on continue le flux normal ci-dessous avec ce même message
    # 2. Flux normal
    messages = [{"role": "system", "content": build_system_prompt()}] + history + [{"role": "user", "content": question}]
    response = client.chat(model="gemma4:31b-cloud", messages=messages, tools=tools)

    if response.message.tool_calls:
        for tool_call in response.message.tool_calls:
            func_name = tool_call.function.name
            func_args = tool_call.function.arguments
            logger.info(f" Appel outil : {func_name}({func_args}) — admin={email}")
            if func_name == "get_my_permissions":
                result = "Super administrateur : accès total à tous les modules." if user.get("isSuperAdmin") else user.get("permissions", {})
            elif not has_permission(user, func_name):
                logger.warning(f" Accès refusé : {func_name} pour {email}")
                result = "Accès refusé : vous n'avez pas les droits nécessaires pour consulter cette information."
            elif func_name in WRITE_FUNCTIONS:
                logger.info(f"✍️ Action en attente de confirmation : {func_name}({func_args})")
                pending_actions[email] = {"function": func_name, "args": func_args}
                description = describe_action(func_name, func_args)
                return f"Vous voulez {description}. Confirmez-vous cette action ? (répondez « oui » pour confirmer, « non » pour annuler)"
            else:
                if func_name in FUNCTIONS_NEEDING_TOKEN:
                    result = available_functions[func_name](**func_args, token=user.get("token"))
                else:
                    result = available_functions[func_name](**func_args)
                logger.info(f"Résultat obtenu pour {func_name} : {str(result)[:150]}")
                if func_name == "filter_demandes" and isinstance(result, dict):
                    if result["rows"]:
                        last_filter_results[email] = {"rows": result["rows"]}
                    result = result["summary"]

                if func_name in CHART_FUNCTIONS and isinstance(result, dict):
                    chart_results[email] = result
                    return f"Voici le graphique demandé : {result['title']}."

                if func_name in DIRECT_RETURN_FUNCTIONS and isinstance(result, str):
                    return result

            messages.append(response.message)
            messages.append({
                "role": "tool",
                "content": json.dumps(result),
                "tool_name": func_name,
            })

        final_response = client.chat(model="gemma4:31b-cloud", messages=messages)
        answer = final_response.message.content

        attempts = 0
        while looks_truncated(answer) and attempts < 2:
            if looks_like_fake_tool_call(answer):
                executed = try_execute_fake_tool_call(answer, user)
                if executed is not None:
                    return executed
            final_response = client.chat(model="gemma4:31b-cloud", messages=messages)
            answer = final_response.message.content
            attempts += 1

        if looks_like_fake_tool_call(answer):
            executed = try_execute_fake_tool_call(answer, user)
            if executed is not None:
                return executed
            return "Je n'ai pas pu formuler de réponse claire à cette question. Pouvez-vous la reformuler ?"

        return answer

    # 3. Aucun outil appelé : protection anti-vide/troncature/faux appel
    answer = response.message.content

    if looks_like_fake_tool_call(answer):
        executed = try_execute_fake_tool_call(answer, user)
        if executed is not None:
            return executed

    attempts = 0
    while looks_truncated(answer) and attempts < 2:
        retry_response = client.chat(model="gemma4:31b-cloud", messages=messages, tools=tools)
        if retry_response.message.tool_calls:
            return _ask_core(question, user, email, history)
        answer = retry_response.message.content
        if looks_like_fake_tool_call(answer):
            executed = try_execute_fake_tool_call(answer, user)
            if executed is not None:
                return executed
        attempts += 1

    if not answer.strip() or looks_like_fake_tool_call(answer):
        return "Je n'ai pas pu formuler de réponse claire à cette question. Pouvez-vous la reformuler, par exemple en précisant le candidat concerné ?"

    return answer


if __name__ == "__main__":
    fake_super_admin = {"isSuperAdmin": True, "permissions": {}, "email": "admin@entreprise.com"}
    print(ask("quel est la fonction de mehdi jlassi", fake_super_admin))
    print("---")
    print(ask("son activité", fake_super_admin))
    print("---")
    print(ask("son poste souhaité", fake_super_admin))