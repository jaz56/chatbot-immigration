from fastapi import FastAPI, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from chat import ask
from auth import verify_token
from fastapi.responses import FileResponse
import os
import logging
import traceback
from fastapi import FastAPI, Depends, HTTPException
app = FastAPI()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:4200", "http://localhost:4201"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class ChatRequest(BaseModel):
    message: str
@app.post("/api/v1/chat")
def chat_endpoint(request: ChatRequest, user: dict = Depends(verify_token)):
    result = ask(request.message, user)
    return {"reponse": result["text"], "chart": result["chart"], "utilisateur": user["email"]}

@app.get("/api/v1/download/{filename}")
def download_export(filename: str, user: dict = Depends(verify_token)):
    filepath = os.path.join("exports", filename)
    if not os.path.exists(filepath) or ".." in filename:
        raise HTTPException(status_code=404, detail="Fichier introuvable")
    return FileResponse(filepath, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

logger = logging.getLogger("chatbot")

@app.post("/api/v1/chat")
def chat_endpoint(request: ChatRequest, user: dict = Depends(verify_token)):
    try:
        result = ask(request.message, user)
        return {"reponse": result["text"], "chart": result["chart"], "utilisateur": user["email"]}
    except Exception as e:
        logger.error(f"💥 Erreur non gérée : {e}\n{traceback.format_exc()}")
        raise
@app.post("/api/v1/chat/reset")
def reset_chat_session(user: dict = Depends(verify_token)):
    from chat import clear_session
    clear_session(user["email"])
    return {"status": "ok"}