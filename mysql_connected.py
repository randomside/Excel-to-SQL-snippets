import pandas as pd
from sqlalchemy import create_engine

table_name = "data_one"
df = pd.read_excel(r"data/item_master.xlsx", index_col=False, header=1)
sql_engine = create_engine(
    'mysql+pymysql://root:110996@127.0.0.1/unicus', pool_recycle=3600)
dbConnection = sql_engine.connect()

df.to_sql(con=dbConnection, name=table_name, if_exists='replace', index=False)
print('DONE')
