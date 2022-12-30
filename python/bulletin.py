# -*- coding: utf-8 -*-
"""
Created on Fri Oct  2 09:58:49 2020

@author: akeller
"""

import json, csv#, os
  
from python.read_duties import read_duties
from python.fill_bulletin import fill_bulletin
from python.create_docs import create_docs


def main():

#    bull_1 = 'SB20-0366'
#    bull_2 = 'SB20-0372'
#    send_date = '11/25/2020'
#    
#    get_bulletins(bull_1, bull_2, send_date)
        
    # Read HASTUS exports to testduties CSV file 
    #return 0
    print('Reading duties from HASTUS files...')
    read_duties()
    
    # Get inputs from files
    print('Getting inputs...')
    constants, bulletins, duties = get_inputs() 
    
    # Iterate through bulletins, fill in fields, and write cover sheets
    print('Writing bulletins...')
    for b in bulletins:
        b = fill_bulletin(b, constants, duties)
        create_docs(b, constants)


    # Convert files to PDF
    print('Converting to PDFs...')


    print('Mail merge complete.')


# Read input files
def get_inputs():
    
    # Get constants from JSON
    try:
        with open('python/inputs/constants.json') as cfile:
            constants = json.load(cfile)
        for j in ['routes', 'garages']:
            with open('python/inputs/{}.json'.format(j)) as jfile:
                constants[j] = json.load(jfile)
    except:
        print('Error reading JSON input files.')

    # Get list of bulletins from CSV
    try:
        with open('python/inputs/bulletins.csv') as bfile:
            breader = csv.DictReader(bfile)
            bulletins = [line for line in breader]
    except:
        print('Error reading inputs/bulletins.csv.')

    # Get list of duties from CSV
    try:
        with open('python/inputs/duties.csv') as dfile:
            dreader = csv.DictReader(dfile)
            duties = [line for line in dreader]
    except:
        print('Error reading inputs/duties.csv.')
        
    return constants, bulletins, duties


if __name__ == '__main__':
    main()