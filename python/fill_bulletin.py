# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 15:37:19 2020

@author: akeller
"""

import datetime as dt
from num2words import num2words
from python.school_mods import create_school_mods


# Calculates merge fields for bulletin
def fill_bulletin(b, constants, duties):
 
    # pulls data from bulletin to set the type and subtype of bulletin
    create_bull_types(b)
    
    # creates data for the header and info box in top right corner
    create_header_data(b, constants)

#    if b['bull_type'] == "Modification_School":
#        schools = create_school_mods()
        
    # creates, counts and make list of duties based on whether extra, part-time revision, full-time modfication    
    d = create_duties(b, constants, duties)

    # create various formats for date fields
    create_dates(b, constants, d)
    
    # creates fields for the body of the bulletin
    create_body(b, constants)
    
    # creates footers for document and tables
    create_footers(b, constants)
    
#    # creates fields for mail items
#    create_mail(b, constants)

    #print (b)
    return b



def create_bull_types(b):
    bull_types = ['Owl', 'Rail', 'Covid', 'Seasonal', 'South Shops', 'Hold', 'School']
    sub_types = {
            'Owl': ['Fall', 'Thanksgiving', 'Christmas', 'Spring', 'Memorial', 'July_4', 'Labor_Day'],
            'Rail': [],
            'Covid': ['Short'],
            'Seasonal': ['Beach'],
            'South Shops': ['Thanksgiving', 'Christmas Eve', 'Good Friday'],
            'Hold': ['President', 'Holiday', 'Thanksgiving'],
            'School': ['Modification']
            }
    for t in bull_types:
        if t.lower() in b['Event'].lower():
            b['type'] = t
            
            for s in sub_types[t]:
                if s in b['Description']:
                    b['bull_type'] = '{}_{}'.format(s, t)

def create_header_data(b, constants):
    # Garage info, activity code, send date
#    try:
    for c in constants['garages'][b['gar_full']]:
        b[c] = constants['garages'][b['gar_full']][c]
    #print("TYPE TYPE TYPE")
    # print(b['type'])
    b['act_code'] = constants[b['type']]['act_code']
    b['send_date'] = dt.datetime.strptime(b['send_date'],'%Y-%m-%d')
    b['send_date'] = dt.datetime.strftime(b['send_date'], '%B %#d, %Y')
 #    except:
#        bulletin_error(b, 'garage_info')


def create_duties(b, constants, duties):
    
    # Parse routes from bulletins
    if len(b['routes']) > 0:
        parse_routes = b['routes'].split(';')
    
        b['re_routes'] = ''
        for r in parse_routes:
            # create string for re_line ie ("#20 Madison/ #53 Pulaski)
            r_string = r + '-' + constants['routes'][r]['name']
            # first route - no '/' in front of route
            if r == parse_routes[0]:
                b['re_routes'] = '#' + r_string
            else:
                b['re_routes'] += ' / #' + r_string
    try:
        b['pick'] = constants['pick']
    except:
        bulletin_error(b, 'duties')
        
    # Duties (add day exceptions and adjust count)
#    try:        
    #create a list of duties for each type
    x_list = []
    a_list = []
    r_list = []
    
    # set all weekday effective dates to 'Weekday'
    if b['eff_day'] in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        b['eff_day'] = 'Weekday'

    gar = constants['garages'][b['gar_full']]['gar_symbol']
    
    for d in duties:
        
        #print(d)
        if b['type'] == 'Covid':
#            print(d['route'])
#            print(b['routes'].split(";"))
            if d['duty'][0] == gar and d['dtype'] == 'extra' and b['eff_day'] == d['op_days'] and d['route'] in b['routes'].split(";"):
                d['bulletin_no'] = b['bulletin_no']
                d['ex_duty'] = d['duty']
                d['bulletin_no'] = b['bulletin_no']
                if int(d['full_pay'].split('h')[0]) < 7:
                    d['full_pay'] == None
                x_list.append(d)

        else: 
            if d['duty'][0] == gar and d['dtype'] == 'extra' and b['eff_day'] == d['op_days']:
                d['bulletin_no'] = b['bulletin_no']
                d['ex_duty'] = d['duty']
                d['bulletin_no'] = b['bulletin_no']
                if int(d['full_pay'].split('h')[0]) < 7:
                    d['full_pay'] == None
                x_list.append(d)
            
        # add all adjusted full- time duties to list
        if d['duty'][0] == gar and d['dtype'] == 'adjusted' and b['eff_day'] == d['op_days']:
            #print(d)
            d['bulletin_no'] = b['bulletin_no']
            
            if b['bull_type'] == 'Spring_Owl':
                pay_split = d['full_pay'].split('h')
                h = int(pay_split[0]) - 1
                m = pay_split[1]
                d['full_pay'] = '{}h{}'.format(h, m)
                d['pay_diff'] = int(d['pay_diff']) - 60
                
            # if adjusted pay is equal to original pay, not bold in cover
            if d['full_pay'] == d['orig_pay']:
                d['sp'] = d['full_pay']
            # if adjusted pay is not equal to original pay, bold in cover
            else:
                d['rp'] = d['full_pay']
            a_list.append(d)

        elif d['duty'][0] == gar and d['dtype'] == 'revised' and b['eff_day'] == d['op_days']:
            d['bulletin_no'] = b['bulletin_no']
            # if spring owl, subtranct 1 hour from duty paid
            if b['bull_type'] == 'Spring_Owl':
                pay_split = d['plat_pay'].split('h')
                h = int(pay_split[0]) - 1
                m = pay_split[1]
                d['plat_pay'] = '{}h{}'.format(h, m)
                d['pay_diff'] = int(d['pay_diff']) - 60
            
            d['rev_duty'] = d['duty']
            # if revised pay is equal to original pay, not bold in cover
            if d['plat_pay'] == d['orig_pay']:
                d['sp'] = d['plat_pay']
            # if revised pay is not equal to original pay, bold in cover
            else:
                d['rp'] = d['plat_pay']
            r_list.append(d)
    
    # count the numbmer of extras, adjustments and revisions
    b['x_count'] = len(x_list)
    b['a_count'] = len(a_list)
    b['r_count'] = len(r_list)
   
    
    # add duties counts for extras
    if b['x_count'] == 1:
        extra = 'extra'
    else: 
        extra = 'extras'
    b['ex_no'] = "{} {}".format(str(b['x_count']), extra)
    b['ex_no_body'] = "{} ({}) {}".format(num2words(b['x_count']), b['x_count'], extra)
    
    # add run counts for adjusted runs
    if b['a_count'] == 1:
        run = 'full_time run'
    else:
        run = 'full-time runs'
    b['adj_no'] = "{} {}".format(str(b['a_count']), run)
    b['adj_no_body'] = "{} ({})".format(num2words(b['a_count']), b['a_count'])
    # add duty counts for revised duties
    if b['r_count'] == 1:
        dty = 'duty'
    else:
        dty = 'duties'
    b['rev_no'] = "{} {}".format(str(b['r_count']), dty)
    b['rev_no_body'] = "{} ({}) part-time {}".format(num2words(b['r_count']), b['r_count'], dty)
    
    # add duties lists to b and sort by duty if more than one
    b['extras'] = x_list
    if b['x_count'] > 1:
        b['extras'] = sorted(b['extras'], key=lambda i: i['duty'])
    b['adjusted'] = a_list
    if b['a_count'] > 1:
        b['adjusted'] = sorted(b['adjusted'], key=lambda i: i['duty'])
        #print (b['adjusted'])
    b['revised'] = r_list
    if b['r_count'] > 1:
        b['revised'] = sorted(b['revised'], key=lambda i: i['duty'])
     
#    except:
#        bulletin_error(b, 'duties')
        
    # check to see if duties have changes to reliefs
    b['reliefs'] = []
    for d in a_list:
        if d['relief_duty'] != d['orig_relief_duty'] and d['relief_duty'] != '':
            for a in a_list:
                if d['relief_duty'] == a['orig_relief_duty']:                    
                    d['rel_duty'] = a['orig_relief_duty']
                    d['sat_duty'] = a['duty']
                    d['night'] = d['op_days']
            b['reliefs'].append(d)
    b['reliefs'] = sorted(b['reliefs'], key = lambda i: i['duty'])
    # count changed reliefs for Time Change owls (no other bulletins shold need this)
    b['rel_count'] = 0
    if b['bull_type'] in ['Fall_Owl', 'Spring_Owl']:
        b['rel_count'] = len(b['reliefs'])
        b['rel_no'] = "{} ({})".format(num2words(b['rel_count']), str(b['rel_count']))
        
            # Main route number and name
    try:
        b['route_no'] = parse_routes[0]
        b['route_name'] = constants['routes'][b['route_no']]['name']
    except:
        bulletin_error(b, 'route info')
    return d


def create_dates(b, constants, d):
        # Formatted effective date
#    try:
    b['eff_date'] = dt.datetime.strptime(b['eff_date'],'%m/%d/%Y')
    b['eff_box'] = dt.datetime.strftime(b['eff_date'], '%B %#d, %Y')
    b['eff_body'] = dt.datetime.strftime(b['eff_date'], '%A, %B %#d, %Y')
    b['day_sh'] = constants['day_types'][d['op_days']]['day_sh']
    b['day_after'] = dt.datetime.strftime((b['eff_date'] + dt.timedelta(days=1)), '%A, %B %#d, %Y')
    b['week_after'] = dt.datetime.strftime((b['eff_date'] + dt.timedelta(days=7)), '%A, %B %#d, %Y')
    b['next_week'] = dt.datetime.strftime((b['eff_date'] + dt.timedelta(days=7)), '%B %#d')
    b['short_day'] = dt.datetime.strftime(b['eff_date'], '%B %#d')
    b['day_of_week'] = dt.datetime.strftime(b['eff_date'], '%A')
    
    #    except:
#        bulletin_error(b, 'effective date')
    if b['type'] in ['Covid']:
        if b['eff_day'] == 'Weekday':
            b['dates'] = constants[b['bull_type']]['dates'].format(b['eff_body'], b['eff_day'], " (except holidays)", b['pick'])
            
        elif b['eff_day'] == 'Sunday':
            b['dates'] = constants[b['bull_type']]['dates'].format(b['eff_body'], b['eff_day'], "/Holiday", b['pick'])
        else:
            b['dates'] = constants[b['bull_type']]['dates'].format(b['eff_body'], b['eff_day'], "", b['pick'])

def create_body(b, constants):
    ## PURPOSE
#    try:
    b['purpose'] = constants[b['bull_type']]['purpose']
    if b['bull_type'] in ['Holiday_Owl', 'July_4_Owl']:
        b['purpose'] = constants[b['bull_type']]['purpose'].format(b['day_sh'])   
    elif b['type'] in ['Owl', 'South Shops'] and b['bull_type'] not in ['Spring_Owl', 'Fall_Owl']:
        b['purpose2'] = constants[b['bull_type']]['purpose2']
        b['purpose3'] = constants[b['bull_type']]['purpose3']
        
    elif b['bull_type'] in ['Beach_Seasonal']:
        b['purpose'] = constants[b['bull_type']]['purpose'].format(b['eff_day'], b['re_routes'])
    
    # add DST note for Daylight Savings Owl
    if b['bull_type'] in ['Fall_Owl', 'Spring_Owl']:
        b['dst1'] = constants[b['bull_type']]['ext1']
        b['dst2'] = constants[b['bull_type']]['ext2']
        b['dst3'] = constants[b['bull_type']]['ext3'].format(b['day_after'])
        
    if b['bull_type'] == 'Fall_Owl' and constants['garages'][b['gar_full']]['gar_symbol'] == '7':
        b['77th'] = constants[b['bull_type']]['77th']
#    except:
#        bulletin_error(b, 'purpose statement')

    # re_line
#    try:
    b['re_line'] = constants[b['bull_type']]['re_line']
    if b['bull_type'] in ['Modification_School']:
        b['file_re'] = constants[b['bull_type']]['file_re'].format(constants['pick'])
    else:
        b['file_re'] = constants[b['bull_type']]['file_re']
#    except:
#        bulletin_error(b, 'route info')
    if b['bull_type'] in ['Beach_Seasonal']:
        b['location'] = constants['{}_{}'.format(b['routes'], b['bull_type'])]['location']
        b['re_line'] = constants[b['bull_type']]['re_line'].format(b['location'])
        b['file_re'] = constants[b['bull_type']]['file_re'].format(b['location'])


def create_footers(b, constants):
    # Footer date and time, and 'Continued'
    now = dt.datetime.now()
    b['initial'] = b['initials']
    b['f_date'] = dt.datetime.strftime(now, '%m/%d/%y')
    b['f_time'] = dt.datetime.strftime(now, '%I:%M %p')
    if b['rel_count'] > 0 or b['a_count'] > 12:
        b['continued'] = '\n- CONTINUED -\n'

    # convert counts to strings
    b['x_count'] = str(b['x_count'])
    b['a_count'] = str(b['a_count'])
    b['r_count'] = str(b['r_count']) 

    # Asterisk explanation (modified full-time service)
    for d in b['adjusted']:
        if b['adjusted'] == True:
            for ast in ['ast1', 'ast2', 'ast3', 'ast4', 'ast5', 'ast6', 'ast7']:
                b[ast] = constants['asterisks'][ast]
            break

#def create_mail(b, constants):
    b['mail_re'] = constants[b['bull_type']]['mail_re']


# Bulletin error handler
def bulletin_error(b, text):
    print('Error getting {} for bulletin {}.'.format(text, b['bulletin_no']))