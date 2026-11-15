from db import db

for name in ['Tmimi', 'Gharbi', 'Bensalah']:
    u = db.users.find_one({
        "$or": [
            {"nom": {"$regex": name, "$options": "i"}},
            {"prenom": {"$regex": name, "$options": "i"}},
        ]
    })
    if not u:
        print(name, "-> INTROUVABLE")
        continue
    print(name, "->", u["_id"], u.get("nom"), u.get("prenom"))

    b1 = db.bookings.find_one({"user": str(u["_id"])})
    print("   booking avec user=str:", b1["_id"] if b1 else None)

    b2 = db.bookings.find_one({"user": u["_id"]})
    print("   booking avec user=ObjectId:", b2["_id"] if b2 else None)