import os
import base64
from datetime import datetime, timezone
from dotenv import load_dotenv
from fastapi import Header, HTTPException
import jwt
from db import get_user_permissions

load_dotenv()

def get_signing_key():
    secret_b64 = os.environ['JWT_SECRET']
    return base64.b64decode(secret_b64)

def verify_token(authorization: str = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Token manquant ou mal formaté")

    token = authorization.replace("Bearer ", "")

    try:
        payload = jwt.decode(token, get_signing_key(), algorithms=["HS256"])
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expiré")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Token invalide")

    email = payload.get("sub")
    user = get_user_permissions(email)

    if user is None:
        raise HTTPException(status_code=403, detail="Utilisateur introuvable")

    # Blocage actif (même logique que AdminPermissionFilter)
    if not user.get("isSuperAdmin") and user.get("isBlocked"):
        raise HTTPException(status_code=403, detail="Compte suspendu")

    return {
        "email": email,
        "isSuperAdmin": user.get("isSuperAdmin", False),
        "permissions": user.get("permissions", {}),
        "token": token,  # 🟢 nécessaire pour relayer vers le backend Spring Boot
    }