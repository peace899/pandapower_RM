from __future__ import absolute_import, division, print_function
import os
import re
from glob import glob
import csv
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import sys
from datetime import datetime
from math import sqrt

feeder = 'Witpoortjie'
directory = '/home/stepper/scripts/python/eskom/MIdrand/CROWTHORNE'
outfile = feeder + '.csv'

os.chdir(directory)

def cleanup():
    try:
        os.remove(outfile)
    except OSError:
        pass


def fix_dates(date, time):
    date_fixed = datetime.strptime(
        str('20' + date), '%Y%m%d').strftime('%Y/%m/%d')
    if ':' in time:
        time_fixed = time
    else:
        time = str(time).zfill(4)
        time_fixed = ':'.join(time[i:i + 2] for i in range(0, len(time), 2))
    time_fixed = time_fixed.replace('24', '00')
    month = datetime.strptime(date_fixed, '%Y/%m/%d').strftime('%b')
    day_hour = '{0} {1}'.format(date_fixed, time_fixed)
    return date_fixed, time_fixed, day_hour, month


def power(P, Q):
    P = float(P)
    Q = float(Q)
    # apparent_power (S), power factor (pf)
    S = sqrt(P**2 + Q**2)
    try:
        pf = P / S
    except ZeroDivisionError:
        pf = 0
    return S, pf


def merge():
    with open(outfile, 'a') as singleFile:
        for csvfile in glob('*.csv'):
            if csvfile == outfile:
                pass
            else:
                for line in open(csvfile, 'r'):
                    if 'RECORDER' in line:
                        pass
                    else:
                        singleFile.write(line)


def fix_data():
    with open(outfile, "r") as f:
        r = csv.reader(f)
        data = [line for line in r]
    with open(outfile, "w") as f:
        writer = csv.writer(f)
        for row in data:
            if 'RECORDER' in row[0]:
                pass
            else:
                date_time = fix_dates(row[1], row[2])
                power_data = power(row[5], row[6])
                row[1] = date_time[2]  # day hour
                row[2] = date_time[0]  # day
                row[3] = date_time[3]  # month
                row[7] = int(power_data[0])
                row[8] = "{0:.2f}".format(power_data[1])
                row = (row[0], row[1], row[2], row[3], row[5], row[6], row[7],
                       row[8])
                writer.writerow(row)


def insert_header():
    with open(outfile, "r") as f:
        r = csv.reader(f)
        data = [line for line in r]
    with open(outfile, "w") as f:
        w = csv.writer(f)
        w.writerow(['RECORDER ID', 'DATE_TIME', 'DATE',
                    'MONTH', 'kW', 'kVAR', 'kVA', 'PowerFactor'])
        w.writerows(data)


def sort_data():
    df = pd.read_csv(outfile)
    df = df.sort_values(by=['DATE_TIME'])
    df.to_csv(outfile, index=False)


def plot(*args):
    df = pd.read_csv(outfile)
    period = args[0]
    if len(args) > 1:
        try:
            select = args[1].capitalize()
        except:
            select = args[1]

    if period == 'year':
        n = 24
        ax = df.kVA.plot(xticks=df.index, rot=90)
        month = list(set(df.MONTH))
        month.sort(key=lambda date: datetime.strptime(date, "%b"))
        month = [v for elt in month for v in (' ', elt)]
        ax.set_xticklabels(month)
    elif period == 'month':
        n = 62
        df_mnth = df[df['MONTH'].str.contains(select)]
        ax = df_mnth.kVA.plot(xticks=df_mnth.index, rot=90)
        day = list(set(df_mnth.DATE))
        day.sort(key=lambda date: datetime.strptime(date, "%Y/%m/%d"))
        day = [v for elt in day for v in (' ', elt)]
        ax.set_xticklabels(day)
    elif period == 'day':
        n = 48
        df_day = df[df['DATE_TIME'].str.contains(select)]
        ax = df_day.kVA.plot(xticks=df_day.index, rot=90)
        hour = list(set(df_day.DATE_TIME))
        hour = [time[11:] for time in hour]
        hour.sort(key=lambda date: datetime.strptime(date, "%H:%M"))
        ax.set_xticklabels(hour)

    plt.locator_params(axis='x', nbins=n)
    plt.show()


def table():
    csv_input = pd.read_csv(outfile)
    a = list(set(csv_input.MONTH))
    a.sort(key=lambda date: datetime.strptime(date, "%b"))
    a = [v for elt in a for v in (' ', elt)]
    df = pd.read_csv(outfile)
    df_mnth = df[df['MONTH'].str.contains("Jul")]
    day = list(set(df_mnth.DATE))
    day.sort(key=lambda date: datetime.strptime(date, "%Y/%m/%d"))
    df_day = df[df['DATE'].str.contains('2016/06/16')]
    hour = list(set(df_day.DATE_TIME))
    hour = [time[11:] for time in hour]
    hour.sort(key=lambda date: datetime.strptime(date, "%H:%M"))
    hour = [v for elt in hour for v in (' ', elt)]
    print(df_mnth)
    


if os.path.isfile(outfile):
    pass
else:
    merge()
    fix_data()
    insert_header()
    sort_data()


if sys.argv[1].lower == 'table':
    table()
    
elif sys.argv[1].lower() == 'plot':
    plot('month', 'mar')
    
elif sys.argv[1].lower == 'refresh':
    cleanup()
    merge()
    fix_data()
    insert_header()
    sort_data()

