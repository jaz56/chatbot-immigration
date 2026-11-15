import os
from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

categories = db.config_items.distinct("category")
print("Catégories disponibles :", categories)
print()

for cat in categories:
    items = list(db.config_items.find({"category": cat, "active": True}, {"value": 1, "label": 1, "_id": 0}))
    print(f"--- {cat} ---")
    for item in items:
        print(f"  {item.get('value')} → {item.get('label')}")