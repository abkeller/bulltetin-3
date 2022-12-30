# -*- coding: utf-8 -*-
"""
Created on Thu Oct  1 10:36:57 2020

@author: akeller
"""

import csv, os
import win32com.client
import datetime as dt

import pythoncom
os.chdir('M:\\akeller\\Python\\Bulletin Creation\\Bulletins-m\\Bulletins')
#bull = 'python\\inputs\\pieces\\bulletin'
orig = 'python\\inputs\\pieces\\orig'
bull = "C:\\Users\\AKeller\Desktop\\bulletin creator\\pieces"

def read_duties():
    # changes excel files to csv
    xl_to_csv(bull)
    xl_to_csv(orig)
    # creates pieces for bulletins and orignals
    bull_pieces = get_pieces(bull)
    orig_pieces = get_pieces(orig)
    # converts and combines pieces in to duties
    bull_duties = create_duties(bull_pieces)
    
    
    orig_list = []
    
    for p in bull_pieces:
        for o in orig_pieces:
            if p['Garage'] + p['Duty'] == o['Garage'] + o['Duty'] and p['Op Day'] in o['Op Day']:
                orig_list.append(o)
                
    orig_duties = create_duties(orig_list)
    # merges bulletin and orginal duties, marks them as extras, modified, or revised
    duties = merge_duties(bull_duties, orig_duties)
    # outputs duty info to a csv for bulletin creator
    output = 'python\\inputs\\duties.csv'
    write_csv(duties, output)
    
# Converts all .xls(x) files in a directory to .csv and deletes .xls(x)
def xl_to_csv(direc):
    excel_files = [x for x in os.listdir(os.path.abspath(direc)) if os.path.splitext(x)[1] in
                   ['.xls', '.xlsx']]

    if len(excel_files) == 0:
        return False

    #Open Word
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    
    for x in excel_files:
        absdirec = '{}\\{}'.format(os.path.abspath(direc), x)
        wb = excel.Workbooks.Open(absdirec)      
        wb.SaveAs(absdirec.replace('.xlsx', '.csv'), FileFormat=24)
        wb.Close()
        os.remove(absdirec)
        
    excel.DisplayAlerts = True
    excel.Quit()

    return True
#    except:
#        print('Error converting Excel files to CSV.')

# Gets pieces from CSVs
def get_pieces(direc):
    try:
        csv_files = [x for x in os.listdir(direc) if os.path.splitext(x)[1] == '.csv']

        pieces = []
        for c in csv_files:
            with open('{}\\{}'.format(direc, c)) as csv_file:
                creader = csv.DictReader(csv_file)
                new_pieces = [line for line in creader]
                    
            pieces.extend(new_pieces)
        return pieces
    except:
        print('Error getting pieces from CSV files.')

# Generates a dictionary for a new duty from piece fields
def create_duties(pieces):
    duties = []
    
    for p in pieces:

        if p['Sequence'] == '1':
            d = {}
            # create full duty name including garage (F230)
            d['duty'] = str(p['Garage']) + str(p['Duty'])
            d['route'] = str(p['Route'])
            
            # create a column for day type, if not 'a' (Saturday) or 's' (Sunday) it's 'w' (weekday)
            if p['Op Day'] not in ['a', 's']:
                d['op_days'] = 'Weekday'
            elif p['Op Day'] == 'a':
                d['op_days'] = 'Saturday'
            else:
                d['op_days'] = 'Sunday'
                
            # create columns for report and pull time
            d['rept'] = str(p['Start Time']).replace(':', '').zfill(4)
            if '+' in p['Start']:
                d['pull'] = p['Start'].replace(':', '').zfill(5)
            else:
                d['pull'] = p['Start'].replace(':', '').zfill(4)
            # create column for start place
            if p['Pce has pullout'] == 'TRUE':
                d['start_pl'] = '#{} @ {}'.format(
                    str(p['Route']),
                    p['First Start']
                )
            else:
                if p['First Start'] == p['From'] or not p['Direction']:
                    direction = ''
                else:
                    direction = p['Direction'][0] 
                    
                d['start_pl'] = 'R{} {} @ {}'.format(
                    direction,
                    p['Garage'][0] + p['Duty1'],
                    p['First Start']
                )
            # if platform time is over 7 hours, offer extra for full-timers

            d['full_pay'] = p['Paid']
#            else:
#                d['full_pay'] = None
            d['plat_pay'] = p['Work Time']
            
            # get next duty number for different reliefs from original (DST OWLS only)
            d['relief_duty'] = p['Next DtyRunNumber']
            # look at second part of two piece duty to get the original next duty for a relief
            for q in pieces:
                if q['Sequence'] == '2' and q['Garage'] + q['Duty'] == d['duty']:
                    d['relief_duty'] == q['Next DtyRunNumber']
                    
            # get duty type
            d['dtype'] = p['Type']
            #print(duties)
            duties.append(d)
    sorted(duties, key = lambda i: i['duty'])
    return duties


def merge_duties(bull_duties, orig_duties):
    od_list = []
    duties = []
    # create a list of all original duties numbers with their operating days appended
    for o in orig_duties:
        od_list.append(o['duty'] + o['op_days'])
    # if bulletin duty not in list of original duties, mark as an 'extra' duty    
    for d in bull_duties:
        #print(d)
        if d['duty'] + d['op_days'] not in od_list:
            d['dtype'] = 'extra'
            d['rev_pay'] = None
            d['orig_pay'] = None
            d['pay_diff'] = pay_diff(d['plat_pay'], "0h00")
            
        else:
            for o in orig_duties:
                if o['duty'] + o['op_days'] == d['duty'] + d['op_days']:
                    # if ostart place different display in bold
                    if d['start_pl'] != o['start_pl']:
                        d['st_plB'] = d['start_pl']
                        d['start_pl'] = ""
                    # if report times different display is bold 
                    if d['rept'] != o['rept']:
                        d['reptB'] = d['rept']
                        d['rept'] = ""
                    # if report times different display is bold 
                    if d['pull'] != o['pull']:
                        d['pullB'] = d['pull']
                        d['pull'] = ""
                    
                        
            # if duty is in orig duties, and it's a part time duty mark as 'revised'    
            if d['dtype'][:3] == 'PTO':
                # if revised duty matches duty in list ...
                for o in orig_duties:
                    if o['duty'] + o['op_days'] == d['duty'] + d['op_days']:
                        # revised pay and orignal pay for part time duties will be the same as platform time
                        d['rev_pay'] = d['plat_pay']
                        d['orig_pay'] = o['plat_pay']
                        # duty for or    
                        d['orig_relief_duty'] = o['relief_duty']
                        d['pay_diff'] = pay_diff(d['rev_pay'], d['orig_pay'])
    
    
                d['dtype'] = 'revised'
                
            else:
                for o in orig_duties:
                    if o['duty'] + o['op_days'] == d['duty'] + d['op_days']:
                        d['rev_pay'] = d['full_pay']
                        d['orig_pay'] = o['full_pay']
                        d['orig_relief_duty'] = o['relief_duty']
                        d['pay_diff'] = pay_diff(d['full_pay'], d['orig_pay'])
                        
                        r_pay = dt.timedelta(hours=int(d['rev_pay'].split('h')[0]), minutes=int(d['rev_pay'].split('h')[1]))
                        o_pay = dt.timedelta(hours=int(d['orig_pay'].split('h')[0]), minutes=int(d['orig_pay'].split('h')[1]))
                        if r_pay < o_pay:
                            d['asterisk'] = True
                            
                d['dtype'] = 'adjusted'
            
        duties.append(d)
    #print(duties)
    return (duties)


def pay_diff(revised, orig):
    r_mins = (int(revised.split('h')[0])) * 60 + int(revised.split('h')[1])
    o_mins = (int(orig.split('h')[0])) * 60 + int(orig.split('h')[1])
    #print(r_mins)
    diff = r_mins - o_mins
    return(diff)
#    r_pay = dt.timedelta(hours=int(revised.split('h')[0]), minutes=int(revised.split('h')[1]))
#    o_pay = dt.timedelta(hours=int(orig.split('h')[0]), minutes=int(orig.split('h')[1]))
#    diff = r_pay - o_pay
#    print(diff)
#    min_diff = diff.total_seconds / 60
#    return(min_diff)

# Writes CSV of duties
def write_csv(duties, output):
    fields = ['duty', 'op_days', 'rept', 'pull', 'start_pl', 'full_pay',
              'plat_pay', 'rev_pay', 'orig_pay', 'relief_duty', 'dtype','orig_relief_duty', 'st_plB', 'reptB', 'pullB', 'route', 'pay_diff', 'asterisk']

#    try:
    with open(output, 'w') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fields,
                                lineterminator='\n')
        writer.writeheader()
        for duty in duties:
            writer.writerow(duty)
#    except:
#        print('Error writing duty CSV file.')
        