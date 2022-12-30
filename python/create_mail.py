# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 17:25:53 2020

@author: akeller
"""

import win32com.client
import os
import pythoncom
from num2words import num2words


def create_mail_items(selected_bulletins, constants):
        

    garages = {}
    for b in selected_bulletins:
        garage = b['gar_full']
        
        # if garage has no dictonary key, create key and add values
        if garage not in garages.keys():
            
            garages[garage] = {
                    "name" : garage, 
                    "re" : b['file_re'],
                    "routes" : [], 
                    "file": b['t_file'],
                    "mail_re": b['mail_re'],
                    "bulletins": [],
                    "bull_type": b['bull_type'],
                    "pick": b['pick']
                    }
        garages[garage]['bulletins'].append(b)
        
        #add route if not in garage level list
        for r in b['routes'].split(";"):
            if r not in garages[garage]['routes']:
                garages[garage]['routes'].append(r)

        
    for g in garages:
        bull_count = len(garages[g]['bulletins'])
        bull_count_str = "{} ({})".format(num2words(bull_count),str(bull_count))
        
        if bull_count > 1:
            sched_bull1 = "Attached are"
            sched_bull2 = "Schedule Bulletins"
        else:
            sched_bull1 = "Attached is"
            sched_bull2 = "Schedule Bulletin"
        # create subject line for mmail, will also be used to name bulletin
        route_str = ""
        if garages[g]['routes'] != [""]:
            for r in garages[g]['routes']:
                route_str = "#{} - ".format(r) + route_str
          
        subject = "Schedule Bulletin - {}{} - {} Garage".format(
            route_str, garages[g]['re'], garages[g]['name']
            )
                
        to = "SB-BusScheduling; SB-AllBulletins; SB-BSM; SB-{}".format(garages[g]['name'].replace(" ", ""))
        
        

        if  garages[g]['bull_type'] == "Modification_School":   
            emailParagraph = """
            Attched is a Schedule Bulletin for <b>{} Garage</b> for <b>School Modifications</b> for the <u>{}</u>.
            <br><br>
            This bulletin contains paddles that should be used when a school is not operating its regular bell schedule. These dates are included on the <b>Master School Calendar</b>, which will be sent out <u>{}</u> and posted on the Planning web page.""".format(garages[g]['name'], garages[g]['pick'], garages[g]['bulletins'][0]['eff_body'])
        
        else:
            emailParagraph = "{} {} {} for <b>{} Garage</b> {}. ".format(sched_bull1, bull_count_str, sched_bull2, garages[g]['name'], garages[g]['mail_re'])


        # create a list of bulletins included for each garage        
        emailBulletins = ''

        for b in garages[g]['bulletins']:     
            bulletin_no = '<br><b>{}:</b>'.format(b['bulletin_no'])
            
            if bull_count == 1:
                eff_date = ' The effective date of this bulletin is <u>{}</u>.<br>'.format(b['eff_body'])
            else:
                eff_date = '<br>Effective Date: <u>{}</u>.<br>'.format(b['eff_body'])
            
            emailBulletins = emailBulletins + bulletin_no + eff_date


        scheduler = constants['schedulers'][b['initials']]
        emailFooter = """
                {}
            <br>
                {}
            <br>
                Chicago Transit Authority
            <br>
                <a href='mailto:{}'>{}</a>
            <br>
                {}
        """.format(scheduler['name'], scheduler['position'], scheduler['email'], scheduler['email'], scheduler['phone'])


        
        htmlBody = """
        <HTML>
            <BODY style='font-family:calibri; font-size:11pt'>
                {}
                <br>
                {}
                <br>
                {}
            </BODY>
        </HTML>
        """.format(emailParagraph, emailBulletins, emailFooter)
        
        olMailItem = 0x0
        pythoncom.CoInitialize()
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail  = obj.CreateItem(olMailItem)
        newMail.subject = subject
        newMail.To = to
        newMail.HTMLBody = htmlBody
        newMail.BodyFormat = 2
        newMail.HTMLBody
        path = os.path.abspath('{}\\{}.msg'.format(b['gar_file'], subject))
        print(path)
        newMail.saveas(path)
