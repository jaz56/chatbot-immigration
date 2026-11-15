import os
import json
from dotenv import load_dotenv
from pymongo import MongoClient
from bson import json_util

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

print("Toutes les collections :", db.list_collection_names())
print()

doc = db.orders.find_one()
print("--- order (exemple) ---")
print(json_util.dumps(doc, indent=2))