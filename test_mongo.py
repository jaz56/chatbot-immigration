import os
from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

client = MongoClient(os.environ['MONGO_URI'])
db = client[os.environ['MONGO_DB_NAME']]

print("Collections disponibles :", db.list_collection_names())