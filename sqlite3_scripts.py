import sqlite3
import pandas as pd

dbname = 'students_performance'
conn = sqlite3.connect(dbname + '.sqlite')

# SQL COMMANDS
# cur = conn.cursor()
# cur.execute('SELECT * FROM Table1')
# for row in cur:
#     print(row)

chunksize = 10000
for chunk in pd.read_csv('students_performance.csv', chunksize=chunksize):
    # replacing spaces with underscores for column names
    chunk.columns = chunk.columns.str.replace(' ', '_')
    chunk.to_sql(name='Performance', con=conn,)


cur = conn.cursor()
cur.execute('SELECT * FROM Performance')

for row in cur:
    print(row)

cur.close()
