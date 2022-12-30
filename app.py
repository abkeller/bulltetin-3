# -*- coding: utf-8 -*-
"""
Created on Thu Nov 12 11:01:54 2020

@author: akeller
"""

from cs50 import SQL
from flask import Flask, flash, jsonify, redirect, render_template, request, session
from flask_session import Session
from tempfile import mkdtemp

import datetime as dt
import csv, os
from python.get_inputs import get_inputs
from python.user_inputs import get_bulletins, select_bulletins
from python.create_docs import create_docs
from python.read_duties import read_duties
from python.fill_bulletin import fill_bulletin



app = Flask(__name__)
app.run(debug=True)

# Ensure responses aren't cached
@app.after_request
def after_request(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response

# Configure session to use filesystem (instead of signed cookies)
app.config["SESSION_FILE_DIR"] = mkdtemp()
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)


path = "python\\inputs\\"
duties = []
previous_bulletins = []
previous_bulletin = []


@app.route("/")
def index():
    return render_template("index.html")

    
@app.route("/new", methods=["GET", "POST"])
def new():

    session['update_spreadsheet'] = True
    session['update_duties'] = True
    session['update_send_date'] = True
    session['use_initials'] = False

    update_bulletins()

    #print(bulletins)
    if request.method == "GET":
        today = dt.datetime.now()
        min_date = today
        max_date = today + dt.timedelta(days=100)

        return render_template("new.html", today=today, min_date=min_date, max_date=max_date, select_bulletins=select_bulletins)
    
    else:

        initials = request.form.get("initials") 
        if initials in session['initials']:
            session['use_initials'] = True
            session['initial'] = initials
            
        update_selected_bulletins()
        update_duties()
        session['return'] = "/new"

        return redirect("/preview")


@app.route("/preview")
def preview():
    # display table of based on selected bulletins
    return render_template("preview.html")


@app.route("/create")
def create():
    flash('Reading duties from HASTUS files...')
    #read_duties()
    
    # Get inputs from files
    flash('Getting inputs...')
    constants = get_inputs()

    # Iterate through bulletins, fill in fields, and write cover sheets
    flash('Writing bulletins...')
    for b in session['selected_bulletins']:
        b = fill_bulletin(b, constants, session['duties'])
        

    flash('Converting to PDFs...')    
    zip_links = create_docs(session['selected_bulletins'], constants)
    print(zip_links)


    # Convert files to PDF
    flash('Mail merge complete.')
    
    return render_template("create.html", zip_links=zip_links)


@app.route("/update", methods=["GET", "POST"])
def update():
    
    if request.method == "GET":
        # create a list of folders for bulletins previously created
        bull_path = "python\\bulletins\\"
        previous_bulletins.clear()
        for f in os.listdir(bull_path):
            value = f
            text = f.replace("_", " ")
            previous_bulletins.append([value, text])
        
        return render_template("update.html", previous_bulletins=previous_bulletins)

    # if method is POST   
    else:        
        # update checkbox fields to true to determine what will be updated to session        
        session['update_spreadsheet'] = False
        session['update_duties'] = False
        session['update_send_date'] = False
        session['use_initials'] = False
        
        prev_bull = request.form.get("prev_bull")
        session['previous_bulletin'] = [prev_bull, prev_bull.replace("_", " ")]   
        
        if request.form.get("bulls") != None:
            session['update_spreadsheet'] = True
            update_bulletins()

        if request.form.get("duties") != None:
            session['update_duties'] = True
        if request.form.get("sendDate") != None:
            session['update_send_date'] = True

        return render_template("update_options.html")


@app.route("/update-options", methods=["GET", "POST"])
def update_options():
    
    if request.method == "GET":
        # if spreadsheet marked true, get updated version of bulletin spreadsheet values and rewrite bulletins list        
        if session['update_spreadsheet'] == True:
            update_bulletins()
            
        return render_template("update_options.html")
    
    else:
        update_selected_bulletins()
        update_duties() 
        session['return'] = "/update-options"
        
        return redirect("/preview")

@app.route("/file-upload")
def file_upload():
    
    return render_template("file_upload.html")

## HELPER FUNCIONS


# strip bulletin number to integer
def bn(bulletin_no): 
    return int(bulletin_no.split("-")[1])
    
# creates paths to access bulletin or duty data
def prev_path():
    # create pathway if using previous bulletin or duty info
    prev_path = "python\\bulletins\\{}\\".format(session['previous_bulletin'][0])
    
    return(prev_path)
    
    
#updates bulletin list
def update_bulletins():
    # create blank list for bulletins
    session['bulletins'] = []
    
    # create a list of initials
    session['initials'] = []
    
    # get bulletins from spreadsheet and move to inputs files as a csv file
    get_bulletins()

    # get lines from csv and place into bullletins list
    with open("{}bulletins.csv".format(path)) as bfile:
        breader = csv.DictReader(bfile)
        for line in breader:
            session['bulletins'].append(line)
            if line['initials'] not in session['initials'] and line['initials'] != "":
                session['initials'].append(line['initials'])


# updates selected bulletins list
def update_selected_bulletins():
    # create blank list for session selected bulletins    
    session['selected_bulletins'] = []
    
    # if creating new bulletins, or updating date for previos bulletins, set session['date'] to true
    if session['update_send_date'] == True:
        send_date = request.form.get("send_date")
    
    # get values from form
    if session['update_spreadsheet'] == True:
        bull1 = request.form.get("bull1")
        bull2 = request.form.get("bull2")
        ## ADD BULLETINS to SELECTED LIST
        # determine whether bulletin creation will be for a range or single
        if bull2 == "Choose Second Bulletin Number ...":
            # if single list bulletin number
            for b in session['bulletins']:
                if b['bulletin_no'] == bull1:
                    b['send_date'] = send_date
                    session['selected_bulletins'].append(b)
        else:
            if session['use_initials'] == True:
                for b in session['bulletins']:
                    if bn(b['bulletin_no']) in range(bn(bull1), bn(bull2) + 1) and session['initial'] == b['initials']:
                        b['send_date'] = send_date
                        session['selected_bulletins'].append(b)

            else:                
                # if range, select bulletin numbers for range
                for b in session['bulletins']:
                    # create range by using digits from bulletin number (i.e. SB20-0123)
                    if bn(b['bulletin_no']) in range(bn(bull1), bn(bull2) + 1):
                        b['send_date'] = send_date
                        session['selected_bulletins'].append(b)
                    
    else:
        # get name of previous bulletin to create path to bulletins        
        with open("{}bulletins.csv".format(prev_path())) as bfile:
            breader = csv.DictReader(bfile)
            for line in breader:
                line['eff_date'] = dt.datetime.strptime(line['eff_date'],'%Y-%m-%d')
                line['eff_date'] = dt.datetime.strftime(line['eff_date'], '%m/%d/%Y')
                session['selected_bulletins'].append(line)
        
        if session['update_send_date'] == True:
            for b in session['selected_bulletins']:
                b['send_date'] = send_date
                
def update_duties():
    # create blank list for session duties
    session['duties'] = []

    if session['update_duties'] == True:
        read_duties()
        with open("{}duties.csv".format(path)) as dfile:
            dreader = csv.DictReader(dfile)
            for line in dreader:
                session['duties'].append(line)
                
    else:
        with open("{}duties.csv".format(prev_path())) as dfile:
            dreader = csv.DictReader(dfile)
            for line in dreader:
                session['duties'].append(line)
                
if __name__=="__main__":
    app.run()