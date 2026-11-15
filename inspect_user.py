import os
import json
from dotenv import load_dotenv
from pymongo import MongoClient
from bson import json_util

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

doc = db.users.find_one({"email": "gabsiy83@gmail.com"})
print(json_util.dumps(doc, indent=2))