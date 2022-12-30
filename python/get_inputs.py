# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 13:25:18 2020

@author: akeller
"""
import json


# Read input files
def get_inputs():
    
    # Get constants from JSON
    try:
        with open('python/inputs/constants.json') as cfile:
            constants = json.load(cfile)
        for j in ['routes', 'garages', 'schools', 'schedulers']:
            with open('python/inputs/{}.json'.format(j)) as jfile:
                #print(j)
                constants[j] = json.load(jfile)
    except:
        print('Error reading JSON input files.')

    # Get list of bulletins from CSV
#    try:
#        with open('python/inputs/bulletins.csv') as bfile:
#            breader = csv.DictReader(bfile)
#            bulletins = [line for line in breader]
#    except:
#        print('Error reading inputs/bulletins.csv.')

    # Get list of duties from CSV
#    try:
#        with open('python/inputs/duties.csv') as dfile:
#            dreader = csv.DictReader(dfile)
#            duties = [line for line in dreader]
#    except:
#        print('Error reading inputs/duties.csv.')
        
    return constants