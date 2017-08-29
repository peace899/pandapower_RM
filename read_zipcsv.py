from __future__ import print_function

import csv
import fnmatch
import logging
import numbers
import numpy as np
import os
import pandas as pd
import re
import zipfile

from calendar import monthrange
from cStringIO import StringIO
from datetime import datetime as dt
from pandas.io.common import CParserError
from tqdm import tqdm, trange
from time import sleep

fmt = '%(filename)s:%(lineno)s %(asctime)s %(levelname)s %(name)s %(message)s'
logging.basicConfig(filename='stats_merge_errors.log', level=logging.DEBUG, 
                    format=fmt)
logger=logging.getLogger(__name__)

# Time the script
startTime = dt.now()

# Files and directories needed
rootPath = r"C:\Users\Peace\Documents\Work\GOU STATS\Jorbeg - Vaal"
dir_stats = os.path.join(rootPath, 'MergedCSVs')
changedcsv_dir = os.path.join(rootPath, 'non compliant csvs')
pattern = '*.zip'
merged_list_file = os.path.join(rootPath, "merged_csvs.txt")

# TODO: unzip found zips inside zip folder
extra_zips = []

# Containers/lists to keep track
zip_files = []
line_data = []

# Check if file already processed
f = open(merged_list_file)
lines = f.readlines()
merged_list = [line.strip() for line in lines]
f.close()

# Create stats directory if not exist
try:
    os.mkdir(dir_stats)
except WindowsError:
    pass

# Update log file    
merged_file = open("merged_csvs.txt", "a")
f_backup = open(merged_list_file, 'a')



def find_files(directory, pattern):
    """ Find all zips in unprocessed data folder."""
    for root, dirs, files in os.walk(directory):
        for basename in files:
            if fnmatch.fnmatch(basename, pattern):
                filename = os.path.join(root, basename)
                zip_files.append(filename)
    
    
    
def digitise(df_col):
    """ Unnecessary but it converts string to number."""
    df = df_col.apply(pd.to_numeric, errors='coerce').fillna(0)
    return df
    
    
def get_zip_csv(filename):
    """ Get csv file from zip and read to pandas dataframe."""    
    with zipfile.ZipFile(filename) as zf:
        for name in zf.namelist():
            combo = (filename, name)
            #print ('Analysing...{}'.format(name))
            if name.endswith('.zip'):
                print ('Another zip named {} in {}'.format(name, filename))
                extra_zips.append(str(combo))
            elif name.endswith(('.MDE', '.log', '.HHF')):
                pass            
            elif name.endswith('/'):
                pass
            else:
                csv_text = StringIO(zf.read(name))
                try:
                    dfa = pd.read_csv(csv_text)
                    prepare_df(dfa, name)
                except CParserError as err:
                    print('Fixing ...{}'.format(name))
                    fix_csv(filename, name, err)
                    

def fix_csv(filename, name, err):
    """ Trick to fix pandas parser tokenizer error and produce dataframe. """
    # Find number of expected columns and use as base
    required_cols = str(err).split()[6]
    n_cols = int(required_cols) - 1
    col_indeces = list(range(n_cols))
    fh = open(filename, 'rb')
    zf = zipfile.ZipFile(fh)
    for f in zf.namelist():
        if name in f:
            csv_text = StringIO(zf.read(f))
            dfa = pd.read_csv(csv_text, usecols=col_indeces)
            dfa = dfa.groupby(lambda x:x, axis=1).sum()
            dfa.reset_index()
                        
            recorder_ids = [i for i in list(set(dfa['RECORDER ID'].values))
                            if 'RECORDER' not in i]
            dframes = []
            for i in recorder_ids:
                df = dfa[dfa['RECORDER ID'] == i]
                            
                kvar_cols = [col for col in df.columns 
                             if 'kvar' in str(col).strip().lower()]
                df['kvar_total'] = 0
                for j in kvar_cols:
                    df['kvar_total'] += df[j].apply(pd.to_numeric,
                                                    errors='coerce').fillna(0)
                df['kvar'] = df['kvar_total']
                df.reset_index()
                            
                df.columns = [str(i).strip().upper() for i in df.columns]
                new_header = ['RECORDER ID', 'DATE', 'HOUR', 'IN', 'UN',
                              'KW', 'KVAR_TOTAL']   
                df2 = df[new_header]
                df2 = df2.reset_index()
                df2 = df[new_header]
                df2.columns = ['RECORDER ID', 'DATE', 'HOUR', 'IN', 'UN',
                               'KW', 'KVAR']
                new_name = '{}_{}'.format(name, i)
                
                prepare_df(df2, new_name)
            """
            for frame in dframes:
                prepare_df(frame, name)
            """    
                
def prepare_df(dfa, name):
    """ Prepare csv dataframe for merging. """
    header = dfa.columns
    date = dfa.iloc[0][1]
    log_tag = '{}_{}'.format(str(name), str(date))
                
    if log_tag in merged_list:
        pass
    else: 
        try:
            if 'RECORDER' in header[0]:
                dfa.columns = [i.strip().upper() for i in dfa.columns]
                    
                if 'RESIDENT1' in name:
                    dfa.rename(columns = {'KW.2':'KVAR'}, inplace = True)
                else:
                    dfa.columns = dfa.columns
                dfa.groupby(lambda x:x, axis=1).sum()
                            
                headers = [i for i in dfa.columns if '.' not in i]
                
                dfa = dfa[headers]
               
                meter_id = dfa.iloc[0][0]
                outfile = os.path.join(dir_stats, str(meter_id) + '.csv')
                        
                get_csv_dataframe(dfa, outfile, log_tag)
                #print('Merged {} to {}.csv'.format(name, meter_id))
                
            else:
                print('transforming ...{}'.format(name))
                transform_file(filename, name)
            """        
            # Log merged csv file
            merged_list.append(log_tag)
            merged_file.write('%s\n' % log_tag)
            """        
        except Exception as err:
            logger.exception('{} on {}'.format(err, log_tag))
            pass


def transform_file(filename, name):
    """ Transform malformed csv...TF 42.csv from cross border files. """
    fh = open(filename, 'rb')
    z = zipfile.ZipFile(fh)
    for f in z.namelist():
        if name in f:
            outpath = changedcsv_dir
            csv_text = StringIO(z.read(f))
    
            dfa = pd.read_csv(csv_text,  header=None)
            lines = dfa.values.tolist()
    
            if len(lines)  in range(100, 1000):
                n = 6
            else:
                n = 3
   
            kw = []
            kvar = []
            date = []
            rows = []
            header = ['RECORDER ID', 'DATE', 'HOUR', 'IN', 'UN', 'kW', 'kvar']
       
            try:
                for i in range(len(lines)):
                    kw_data = lines[(n * i)][8:]
                    kw_data = [int(f) for f in kw_data]
                    kw.extend(kw_data)
                    kvar_data = lines[(n * i) + (n / 3)][8:]
                    kvar_data = [int(f) for f in kvar_data]
                    kvar.extend(kvar_data)
                    date_data = [str(lines[(n * i)][1])] * 48
                    date.extend(date_data)
            except IndexError:
                pass
        
            hr = [30, 100, 130, 200, 230, 300, 330, 400, 430, 500, 530, 600, \
                  630, 700, 730, 800, 830, 900, 930, 1000, 1030, 1100, 1130, \
                  1200, 1230, 1300, 1330, 1400, 1430, 1500, 1530, 1600, 1630, \
                  1700, 1730, 1800, 1830, 1900, 1930, 2000, 2030, 2100, 2130, \
                  2200, 2230, 2300, 2330, 2400] 

            hour = [hr] * len(set(date))
            hour = sum(hour, [])
            recorder = lines[0][0].replace('"', '')
            recorder = [recorder] * len(date)
            un = ['KW'] * len(date)
            interval = ['30'] * len(date)
    
            for i in range(len(date)):
                row = (recorder[i], date[i], hour[i], interval[i], un[i],
                       kw[i], kvar[i])
                rows.append(row)
    
            data = [list(r) for r in rows]
            dfa = pd.DataFrame(data)
            dfa.columns = [i.strip().upper() for i in header]
            meter_id = dfa.iloc[0][0]
            outfile = os.path.join(dir_stats, str(meter_id).strip() + '.csv')
            date = dfa.iloc[0][1]
            log_tag = '{}_{}'.format(str(name), str(date))
            get_csv_dataframe(dfa, outfile, log_tag)                  

    
def time_fix(time):
    """ Format the date and time strings. """
    if ':' in time:
        time_fixed = time
    else:
        time = str(time).zfill(4)
        time_fixed = ':'.join(time[i:i + 2] for i in range(0, len(time), 2))
    return time_fixed.replace('24', '00')

    
def get_csv_dataframe(dfa, outfile, log_tag):
    """ Produce final dataframe in required format. """
    dfx = dfa
    header = [i.strip().upper() for i in dfx.columns if i is not '']
    dfx = dfx[header]
    cut_off_index = [i for i, j in enumerate(dfx['DATE'].values)
                     if 'date' in str(j).strip().lower()]
    
    if len(cut_off_index) > 0:
        n = len(dfx) - cut_off_index[0] - 1
    else:
        n = len(dfx)
        
    df = dfx.head(n)
    
    df['KVAR'] = df['KVAR'].apply(pd.to_numeric, errors='coerce').fillna(0)
    df['KW'] = df['KW'].apply(pd.to_numeric, errors='coerce').fillna(0)
    kvar_squares = df['KVAR']**2
    kw_squares = df['KW']**2
    df['KVA'] = (kvar_squares + kw_squares)**(1./2)
    df['KVA'] = df['KVA'].astype(int)
    
    df['POWERFACTOR'] = np.where(df['KVA'] < 1, df['KVA'],
                                 df['KW']/df['KVA']).round(2)
                          
    date = [('20' + str(i))[:8] for i in df['DATE'].tolist()]
    date = [dt.strptime(i, '%Y%m%d').strftime('%Y/%m/%d') for i in date]
    month = [dt.strptime(i, '%Y/%m/%d').strftime('%b') for i in date]
    tm_list = [time_fix(str(i)) for i in df['HOUR'].tolist()]
    date_time = ['{0} {1}'.format(i, j) for i, j in zip(date, tm_list)]
    df['DATE'] = pd.Series(date).values
    df['MONTH'] = pd.Series(month).values
    df['DATE_TIME'] = pd.Series(date_time).values
    df = df[['RECORDER ID', 'DATE_TIME', 'DATE', 'MONTH',
             'KW', 'KVAR', 'KVA', 'POWERFACTOR']]
    
    appendDFToCSV_void(df, outfile, log_tag, sep=",")
    
    
def appendDFToCSV_void(df, csvFilePath, log_tag, sep=","):
    """ Create csv from dataframe or append to existing csv."""
    if not os.path.isfile(csvFilePath):
        print ('Creating... {}'.format(csvFilePath))
        df.to_csv(csvFilePath, mode='a', index=False, sep=sep)
        f_backup.write('%s\n' % log_tag)
        merged_list.append(log_tag)
        merged_file.write('%s\n' % log_tag)
        
    elif len(df.columns) != len(pd.read_csv(csvFilePath,
                                            nrows=1, sep=sep).columns):
        raise Exception("Columns do not match!! Dataframe has " +
                        str(len(df.columns)) + " columns. CSV file has " +
                        str(len(pd.read_csv(
                        csvFilePath, nrows=1, sep=sep).columns)) +
                        " columns.")
    elif not (df.columns == pd.read_csv(csvFilePath, nrows=1,
                                        sep=sep).columns).all():
        raise Exception("Columns and column order of "
                        "dataframe and csv file do not match!!")
    else:
        #print ('Appending to... {}'.format(csvFilePath))
        df.to_csv(csvFilePath, mode='a', index=False, sep=sep, header=False)
        f_backup.write('%s\n' % log_tag)
        merged_list.append(log_tag)
        merged_file.write('%s\n' % log_tag)
    

# TODO: extract zips in zipfolder
def extract_zipfile(filename, other_zip):
   with zipfile.ZipFile(filename, 'r') as zf:
    for name in zf.namelist():
        if name in other_zip:
             get_zip_csv(os.path.dirname(name))
              
          
# Look for zip files
find_files(rootPath, pattern)

# Set tqdm scale
bar = trange(len(zip_files))
for i in bar:
   
    filename = zip_files[i]
    sleep(0.1)
    get_zip_csv(filename)
    tqdm.write("Processed %s" % filename)
   

# Finalize write to merged list file
merged_file.close()
f_backup.close()
# Report on script run
f = open("merged_csvs.txt")
lines = f.readlines()
files_t = len(list(set(lines)))

total_time = dt.now() - startTime
print('{0} files succesfully merged in {1}'.format(files_t, total_time))  
