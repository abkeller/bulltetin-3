# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 16:34:07 2020

@author: akeller
"""

from mailmerge import MailMerge
import win32com.client
import os, shutil

import pythoncom
import pandas as pd

from python.create_mail import create_mail_items
#import datetime as dt

#files = []

def create_docs(selected_bulletins, constants):
    
    # create word document from template
    write_docx(selected_bulletins, constants)
    
    # create pds from word doc
    write_pdf(selected_bulletins)
    
    move_files(selected_bulletins)
    
    create_mail_items(selected_bulletins, constants)
    
    return zip_files()

def write_docx(selected_bulletins, constants):

#    try:
    for b in selected_bulletins:
        #print(b)
        # Check if garage directories exist; if not, create them
        b['t_file'] = 'python\\bulletins\\{}-{}'.format(b['bull_type'], b['initials']).replace(" ", "_")
#        if os.path.exists(b['t_file']):
#            shutil.rmtree(os.path.abspath(b['t_file']))
        if not os.path.exists(b['t_file']):            
            os.mkdir(b['t_file'])
        
        c_file = "{}\\cover_sheets".format(b['t_file'])
        if not os.path.exists(c_file):
            os.mkdir("{}\\cover_sheets".format(b['t_file']))
   
        b['gar_file'] = '{}\\{}'.format(c_file, b['gar_full'])
        if not os.path.exists(b['gar_file']):
            os.mkdir(b['gar_file'])
    
        b['b_file'] = '{}\\{}'.format(b['gar_file'], b['eff_day'])
        if not os.path.exists(b['b_file']):
            os.mkdir(b['b_file'])

        # if not routes, omit hypen from doc name         
        if len(b['routes']) > 0:
            hyphen = "-"
        else:
            hyphen = " "
    #    try:
        b['filename'] = '{}\\{} {}{}{}-{} {}.docx'.format(
            b['b_file'],
            b['bulletin_no'],
            b['routes'].replace(';', '-'),
            hyphen,              
            b['gar_symbol'],
            constants['day_types'][b['eff_day']]['paper'],
            b['file_re']
        )        
    
    
        # Choose template for Owls
        if b['type'] == 'Owl':
            path = 'python/templates/owls/'
            # template for owl with modifications and reliefs
            if b['bull_type'] == 'Holiday_Owl':
                template = '{}{}'.format(path, 'holiday_adjusted.docx')
            elif int(b['r_count']) != 0 and int(b['rel_count']) != 0:
                template = '{}{}'.format(path, 'reliefs_revised.docx')
            # template for owl with modifications and reliefs
            elif int(b['r_count']) == 0 and int(b['rel_count']) != 0:
                template = '{}{}'.format(path, 'reliefs.docx')
            # template for owl with modifications, revisions, and reliefs
            elif int(b['r_count']) != 0 and int(b['rel_count']) == 0:
                template = '{}{}'.format(path, 'revised.docx')  
            # template for owl with modifications only
            else:
                template = '{}{}'.format(path, 'adjusted.docx')
        # choose templates for Covid
        elif b['type'] == 'Covid':
            path = 'python/templates/extras/'
            template = '{}{}'.format(path, 'covid.docx')
        # choose templates for Seasonal
        elif b['type'] in ['Seasonal'] and b['eff_day'] in ['Saturday', 'Sunday']:
            path = 'python/templates/extras/'
            template = '{}{}'.format(path, 'beach_shuttle_weekend.docx')
        # choose template for standard revised bulletins
        elif b['type'] in ['South Shops', 'Hold']:
            path = 'python/templates/standard/'
            template = '{}{}'.format(path, 'revised.docx')
        elif b['bull_type'] == "Modification_School":
            path = 'python/templates/school/'
            template = '{}{}'.format(path, 'modifications.docx')
        else:
            print("Unable to located bulletin type")
    
        #files.append(b['b_file'])
    
        # Merge field values
        with MailMerge(template) as document:
            document.merge(**b)
            #document.merge('M:/akeller/Python/Bulletin Creation/bulletin_creator - v3/templates/extras/63_beach.docx')
            document.merge_rows('duty', b['adjusted'])
            document.merge_rows('rev_duty', b['revised'])
            document.merge_rows('rel_duty', b['reliefs'])
            document.merge_rows('ex_duty', b['extras'])
    
            # Write output and close 
            document.write(b['filename'])


def write_pdf(selected_bulletins):

    try:
        #Open Word
        pythoncom.CoInitialize()    
        word = win32com.client.Dispatch('Word.Application')
        # use the filenmames created for word docs, and translate to pdfs 
        for b in selected_bulletins:
            print(b['bulletin_no'])
            # create absolute path to word docuements
            path = os.path.abspath(b['filename'])
            doc = word.Documents.Open(path)
            doc.SaveAs(path.replace('.docx', '.pdf'), FileFormat=17)
            doc.Close()

    except:
        word.Quit()
        print('Error converting .docx files to PDFs.')
        
dirs = []        
# move bulletins and duties to bulletin type files for later reference
def move_files(selected_bulletins):
    # create list of bulletin types that have been created
    for b in selected_bulletins:
        if b['t_file'] not in dirs:
            dirs.append(b['t_file'])

    df = pd.DataFrame(selected_bulletins)#, columns=columns)
    for b in dirs:
        df1 = df[df.t_file == b]
        duties = []
        for s in selected_bulletins:
            if s['t_file'] == b:
                if len(s['revised']) > 0:
                    for i in s['revised']:
                        duties.append(i)
                if len(s['adjusted']) > 0:
                    for i in s['adjusted']:
                        duties.append(i)
                if len(s['extras']) > 0:
                    for i in s['extras']:
                        duties.append(i)
                       
        ##### create data frame of all duites, and export as csv ######
        # convert duties to dataframe
        duties = pd.DataFrame(duties)
        print(duties.pay_diff)
        duties.pay_diff = duties.pay_diff.astype('int')

        # sum all pay differences by bulletin and merge with bulletins based on bulletin number
        hours = duties.groupby('bulletin_no', as_index=False).agg({'pay_diff': "sum"})
        merge = pd.merge(df1, hours, on=['bulletin_no'])
        print(merge)

        # add a plus to bulletins with a positive paid change
        merge['plus'] = '+'
        merge['plus_minus'] = merge.loc[merge['pay_diff'] > 0, 'plus']
        merge['pd_abs'] = merge.pay_diff * -1
        merge['pay_diff'] = merge.loc[merge['pay_diff'] < 0, 'pd_abs']
        print(merge['pay_diff'])
        
        # convert cost to time format
        merge['Cost'] = (merge.pay_diff / 60).astype('int').astype('str') + ':' +  (duties.pay_diff % 60).astype('str').str.zfill(2)

        # determine which columns to  sve and convert to csv
        merge = merge[['bulletin_no', 'routes', 'gar_full', 'Event', 'Description', 'eff_day', 'eff_date', 'initials', 'bull_type', 'plus_minus', 'Cost']]
        merge.to_csv("{}\\bulletins.csv".format(b))
        duties.to_csv("{}\\duties.csv".format(b))


# helper function to create a zip file containing all output
def make_archive(source, destination):
        base = os.path.basename(destination)
        name = base.split('.')[0]
        format = base.split('.')[1]
        archive_from = os.path.dirname(source)
        archive_to = os.path.basename(source.strip(os.sep))
        #print(source, destination, archive_from, archive_to)
        shutil.make_archive(name, format, archive_from, archive_to)
        shutil.move('%s.%s'%(name,format), destination)



#def zip_files():
#    links = []
#    for d in dirs:
#        link = '{}/zip.zip'.format(d)
#        make_archive(os.path.abspath(d), link)
#        links.append([os.path.abspath(link), d.split("\\")[-1]])
#    return links

def zip_files():
    print(dirs)
    links = []
    for d in dirs:
        source = d
        destination = 'C:\\Users\\AKeller\Desktop\\bulletin creator'
        link = '{}\\bulletins.zip'.format(destination)
        make_archive(source, link)
        links.append(link)
        return links       
        
