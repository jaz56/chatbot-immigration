import os
import json
from dotenv import load_dotenv
from pymongo import MongoClient
from bson import json_util

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

for name in ['demandes', 'bookings']:
    doc = db[name].find_one()
    print(f"--- {name} ---")
    print(json_util.dumps(doc, indent=2))
    print()