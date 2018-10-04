import pandas as pd
import sqlite3

writer = pd.ExcelWriter('glfexcel.xlsx', engine='xlsxwriter')

def get_object_ids(table, object):
    """convert binary to text"""
    
    object_ids = connection.execute("""SELECT
                 substr(hguid, 7, 2) || substr(hguid, 5, 2) || 
                 substr(hguid, 3, 2) || substr(hguid, 1, 2) || '-' ||
                 substr(hguid, 11, 2) || substr(hguid, 9, 2) || '-' || 
                 substr(hguid, 15, 2) || substr(hguid, 13, 2) || '-' || 
                 substr(hguid, 17, 4) || '-' || 
                 substr(hguid, 21, 12) AS {1}
                 FROM (SELECT hex({1}) AS hguid FROM {0})""".format(table,
                                    object))
                                    
    object_ids = [i[0] for i in object_ids.fetchall()]
    return object_ids

db_file = '/home/stepper/Downloads/GLF_Barcelona_20170627_0.0.glf'
connection = sqlite3.connect(db_file)

tables = connection.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
table_names = [i[0] for i in tables.fetchall() if i[0].startswith('tb_l')]

for table in table_names:
    sql = 'select * from {}'.format(table)
    df = pd.read_sql(sql, con=connection)
    fguid_cols = [i for i in df.columns if 'fguid' in i]
    for col in fguid_cols:
        col_ids = get_object_ids(table, col)
        df[col] = pd.Series(col_ids)
    df = df.applymap(str)
    df.to_excel(writer, sheet_name=table, index=False)
    
writer.save()
    #print(fguid_cols)
    #break
#print(table_names)
