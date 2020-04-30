import postgresql

db = postgresql.open(user = 'usename', database = 'dataname', port = 5432, password = 'secret')
