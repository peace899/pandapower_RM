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

feeder = ''
directory = ''
outfile = feeder + '.csv'

os.chdir(directory)

csv_input = pd.read_csv(outfile)
a = list(set(csv_input.MONTH))
a.sort(key=lambda date: datetime.strptime(date, "%b"))
a = [v for elt in a for v in (' ', elt)]
df = pd.read_csv(outfile)
month = list(set(df.MONTH))
month.sort(key=lambda date: datetime.strptime(date, "%b"))

for mnth in month:
    df_mnth = df[df['MONTH'].str.contains(mnth)]
    df_mnth_max = df_mnth['kVA'].max()
    print(mnth, df_mnth_max)
day = list(set(df_mnth.DATE))
day.sort(key=lambda date: datetime.strptime(date, "%Y/%m/%d"))
df_day = df[df['DATE'].str.contains('2016/06/16')]
hour = list(set(df_day.DATE_TIME))
hour = [time[11:] for time in hour]
hour.sort(key=lambda date: datetime.strptime(date, "%H:%M"))
hour = [v for elt in hour for v in (' ', elt)]

df_mnth_max = df_mnth['kVA'].max()
idx = df.groupby(['kVA', 'MONTH', 'DATE']).kVA.agg(lambda x: x.idxmax())

