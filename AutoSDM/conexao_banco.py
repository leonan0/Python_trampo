import pyodbc 
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=07HW011196\SQLEXPRESS;'
                      'Database=sdm;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
