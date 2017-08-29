import pandas as pd
import sqlite3
import os
import unicodecsv as csv
import xlwt

f = r'C:\Users\lekalap\Desktop\Digsilent MV Training\GLF_Barcelona_20170627_0.0.glf'
f_name = os.path.splitext(f)[0]
xls_file = str(f_name) + '.xls'
#writer = pd.ExcelWriter(xls_file, engine='xlsxwriter')
currdir = os.getcwd()
csv_dir = os.path.join(currdir, 'csv')

db = sqlite3.connect(f)
cursor = db.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
for table_name in tables:
    table_name = table_name[0]
    csv_file = os.path.join(csv_dir, '{}.csv'.format(table_name))
    SQL = 'SELECT * FROM [{}];'.format(table_name)
    rows = cursor.execute(SQL).fetchall()
    #table = pd.read_sql_query("SELECT * from %s" % table_name, db)
    print 'adding table ...{}'.format(table_name)
    field_names = [i[0] for i in cursor.description]
    #table.to_csv(csv_file, encoding='utf-8', index_label='index')
    with open(csv_file, 'wb') as fou:
        #csv_writer = csv.writer(fou) # default field-delimiter is ","
        csv_writer = csv.writer(fou, dialect='excel', encoding='utf-8')
        csv_writer.writerow(field_names)
        csv_writer.writerows(rows)
    
#print table_name
#SQL = 'SELECT * FROM [{}];'.format(table)
