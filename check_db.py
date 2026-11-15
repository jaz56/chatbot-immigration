import os
from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

for name in db.list_collection_names():
    count = db[name].count_documents({})
    print(f"{name}: {count} documents")