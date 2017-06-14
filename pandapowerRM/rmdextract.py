#Extract data from .rmd file
# Data in form of spans or sections of line/cable, transformers/minisubs

from __future__ import print_function
import os, sys
import re
import fnmatch

criteria_span = 'Conductor'
criteria_trfr = 'Trfr'

#input directory where .rmd files reside
directory = ''
#input name of feeder
fdr = ''

def rmd_file(fdr):
    for f in os.listdir(directory):
        if fnmatch.fnmatch(f, fdr):
            _rmd_file = os.path.join(directory, f)
        return _rmd_file

def spans(fdr):
    spans = []
    bus_poles = []
    lines = open(rmd_file(fdr), "r").readlines()
    for line in lines:
        if re.search(criteria_span, line):
            span = line.split
            span_length = line.split
            span_conductor = line.split        


def transformers(fdr):
    lines = open(rmd_file(fdr), "r").readlines()
    for line in lines:
        if re.search(criteria_trfr, line):
        print(line, end="")
		
