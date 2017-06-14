#Extract data from .rmd file
# Data in form of spans or sections of line/cable, transformers/minisubs

import os, sys
import re
import fnmatch

criteria_span = 'Conductor'
criteria_trfr = 'Trfr'

#select or input path to where .rmd files reside
directory = ''
#input feeder name
fdr = ''

def rmd_file(fdr):
    """Find feeder's .rmd file"""
    for f in os.listdir(directory):
        if fnmatch.fnmatch(f, fdr):
            _rmd_file = os.path.join(directory, f)
        return _rmd_file

def spans(fdr):
    """Get sections of the feeder"""
        


def transformers(fdr)
    """Get feeder transformers"""
    lines = open(rmd_file(fdr), "r" ).readlines()
    transformers = []
    for line in lines:
        if re.search( criteria_trfr, line ):
            print(line, end="")
