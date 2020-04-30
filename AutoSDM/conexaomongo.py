import pymongo
import datetime

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client.test
db = client.autosdm
col = db.incidente

hist_col = db.historico
plant_coll = db.plantao



def limpabase():
    x=col.find()
    for a in x:
        col.find_one_and_delete(a)

'''x = 1

query = {'severidade': 2}
h = col.find(query)

for x in h:
    print(x)

def limpabase():
    x=col.find()
    for a in x:
        col.find_one_and_delete(a)
query = '"numero":"0"'

x = col.find()
print (x.len())

for a in x:
    col.find_one_and_delete(a)

def teste():
    x = col.find({},{"_id":0})

    for a in x:
        print(a)

teste()'''

