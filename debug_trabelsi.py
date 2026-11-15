from db import db

for u in db.users.find({
    "nom": {"$regex": "trabelsi", "$options": "i"},
    "prenom": {"$regex": "yassine", "$options": "i"},
}):
    print("USER:", u["_id"], u.get("nom"), u.get("prenom"), u.get("email"))
    for d in db.demandes.find({"user": str(u["_id"])}):
        print("   demande (str):", d.get("uniqueId"), d.get("status"))
    for d in db.demandes.find({"user": u["_id"]}):
        print("   demande (ObjectId):", d.get("uniqueId"), d.get("status"))