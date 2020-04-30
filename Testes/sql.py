import pymongo

client = pymongo.MongoClient("mongodb://admin:admin@plantao")
db = client.test
print(db)

