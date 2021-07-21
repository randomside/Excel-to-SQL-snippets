import pandas as pd
from sqlalchemy import create_engine


def import_to_data_csv(file_name):
    table_name = file_name

    df = pd.read_csv(f"Byun/{file_name}.csv", index_col=False,
                     header=0)  # Read data start from row 2

    sql_engine = create_engine(
        'mysql+pymysql://root:110996@127.0.0.1/unicus', pool_recycle=3600
    )
    dbConnection = sql_engine.connect()

    df.to_sql(con=dbConnection, name=table_name,
              if_exists='replace', index=False)

    print('DONE')


import_to_data_csv('BOM')
import_to_data_csv('exim')
import_to_data_csv('invoice')
import_to_data_csv('mvt_item')
# table_name = "data_one"

# df = pd.read_excel(r"data/item_master.xlsx", index_col=False,
#                    header=1)  # Read data start from row 2

# sql_engine = create_engine(
#     'mysql+pymysql://root:110996@127.0.0.1/unicus', pool_recycle=3600
# )
# dbConnection = sql_engine.connect()

# df.to_sql(con=dbConnection, name=table_name, if_exists='replace', index=False)

# print('DONE')
