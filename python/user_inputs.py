# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 10:35:41 2020

@author: akeller
"""
import datetime as dt
import pandas as pd



def get_bulletins():
    # grab the excel file for bulletin numbers
    #'I:/Bus Scheduling/2020 Bulletin Numbers.xls'
    
    #df = pd.read_excel(r'python/inputs/2021 Bulletin Numbers - test.xls',
    df = pd.read_excel(r'I:/Bus Scheduling/2021 Bulletin Numbers.xls',
#    df.to_csv(r'python/inputs/bulletins.csv', index= None, header=0)
#    df = pd.read_csv('python/inputs/bulletins.csv',
        skiprows = 11,
        usecols = [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
        #na_values = None,
        dtype = "str",
        names = [
            'bulletin_no',
            'route_1',
            'route_2',
            'route_3',
            'route_4',
            'route_5',
            'route_6',
            'gar_full',
            'Event',
            'Description',
            'Effective',
            'initials'
        ])

    
    # set bulletin numbers as index of dataframe for later reference
    df = df.drop_duplicates(subset=['bulletin_no'], keep='last')
    df.set_index('bulletin_no', inplace=True)
    # convert all null values to blanks
    df = df.fillna("")


    # create columns 'routes' with the first values set as the first route column from spreadsheet
    df['routes'] = df['route_1'].astype('str')

    # iterate through dataframe rows based on bulletin numbers
    for row in df.index:
        # iterate through all columns that may contain route numbers and convert to strings
        for i in range(2,6):
            route = str(df.loc[row]['route_{}'.format(i)])
            # if route is present in column
            if route != '':
                # if route was taken in as a number (not ideal solution, but tried lots of other methods) cut after the decimal point
                if '.' in route:
                    route = route.split('.')[0]
                # add route to routes column with ';' in front to be parsed by bulletin
                df.loc[row]['routes'] = df.loc[row]['routes'] + ';' + route
                
    df['nan'] = df['routes'] == 'nan'
    df.loc[df['nan'] == True, 'routes'] = df.loc[df['nan'] == True, 'routes'] = ""
    #merge.loc[merge['negative'] == True, 'plus-minus'] = merge.loc[merge['negative'] == True, 'plus-minus'] = "-"
                
    # convert effective date to 'mm/dd/yyyy' format

    df.Effective = pd.to_datetime(df.Effective)
    # keep these columns and return dataframe
    #df = df[df.Effective >= dt.datetime.now()]
    # convert effective date to 'mm/dd/yyyy' format
    df['eff_date'] = df.Effective.dt.strftime('%m/%d/%Y')
    df['eff_day'] = df.Effective.dt.strftime('%A')
    df = df[['routes', 'gar_full', 'Event', 'Description', 'eff_day', 'eff_date', 'initials']]
    #print(df.Event)
    
    df.to_csv('python/inputs/bulletins.csv')
    return df     
      

def select_bulletins(bull1, bull2, send_date):
    print(bull1)
    print(bull2)
    print(send_date)
    df = pd.read_csv('python/inputs/bulletins.csv')
    #print(df)
    # create range of bulletins to grab from spreadsheet
    if bull2 == "":
        df = df.loc[bull1]
    else:
        df = df.loc[bull1:bull2]
        
    #print(df)
    
    df['send_date'] = send_date
    
    df.to_csv('python/inputs/bulletins.csv')
select_bulletins('SB20-0367', 'SB20-0372', '11/19/2020')   