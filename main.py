
# -*- coding:utf-8 -*-
import os
import sys
from datetime import date
from flask import request, render_template, session
from flaskwebgui import FlaskUI #get the FlaskUI class

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__),os.pardir))
sys.path.append(PROJECT_ROOT)

from app import create_app
from app.forms import SelectionsForm
from main_aux import flash_mess, excecute_main

app = create_app()


# Feed it the flask app instance 
ui = FlaskUI(app,maximized=True, start_server='flask')


# do your logic as usual in Flask
@app.route('/', methods=["GET", 'POST'])
def index():
    return render_template('index.html')

# Open the .xlsx file
@app.route('/xlsx_file', methods=["GET", 'POST'])
def xlsx_file():
    xlsx_path = os.getcwd()
    xlsx_path = os.path.join(xlsx_path + "\Activity_Guide_Changes\GaN_cons_and_dates.xlsx")
    os.startfile(xlsx_path)
    return render_template('index.html')

# Open the folder whith the .docx files     
@app.route('/docx_files', methods=["GET", 'POST'])
def docx_files():
    docx_path = os.getcwd()
    docx_path = os.path.join(docx_path + "\docs_to_change")
    os.startfile(docx_path)
    return render_template('index.html')

# Open the folder with the .pdf files
@app.route('/pdf_files', methods=["GET", 'POST'])
def pdf_files():
    pdf_path = os.getcwd()
    pdf_path = os.path.join(pdf_path + "\pdf_files")
    os.startfile(pdf_path)
    return render_template('index.html')   

# Render the selections form
@app.route('/selections2', methods=["GET", 'POST'])
def selections2():
    selections_form = SelectionsForm()
    selections_form.year.default = date.today().year
    selections_form.process()

    north_consts = session.get('north_consts')
    context = {
        'selections_form': selections_form
        }

    
    # When Submit get the Lists for each item
    if request.method == 'POST':
        year=request.form.get('year')
        north_consts=request.form.getlist('north_consts')
        north_langs=request.form.getlist('north_langs')
        north_lats=request.form.getlist('north_lats')
        south_consts=request.form.getlist('south_consts')
        south_langs=request.form.getlist('south_langs')
        south_lats=request.form.getlist('south_lats')
        download_check = request.form.get('download_charts')

        # Get an error message when there are wrong selections or permit to do all the operations
        message = flash_mess(north_consts,north_langs, north_lats, south_consts, south_langs, south_lats)
        
        # Continue if the user makes the right selections
        if message == 'ok':
            excecute_main(year, north_consts, north_langs, north_lats, south_consts, south_langs, south_lats, download_check)
            
    return render_template('selections2.html', **context)


# error handler 400
@app.errorhandler(404)
def not_found(error):

    return render_template('404.html', error=error)

# Error Handler 500
@app.errorhandler(500)
def not_found(error):

    return render_template('500.html', error=error)



if __name__ =='__main__':

    #run the app
    ui.run()
    

    
