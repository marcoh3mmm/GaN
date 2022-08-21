import sys, os
if sys.executable.endswith('pythonw.exe'):
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.path.join(os.getenv('TEMP'), 'stderr-{}'.format(os.path.basename(sys.argv[0]))), "w")
    
# -*- coding:utf-8 -*-
import os
import sys
from datetime import date
from flask import request, make_response, redirect, render_template,session,flash, url_for
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
    #generate_activity_guides = make_response(redirect('/selections2'))
    return render_template('index.html')

@app.route('/xlsx_file', methods=["GET", 'POST'])
def xlsx_file():
    os.chdir('../')
    xlsx_path = os.getcwd()
    xlsx_path = os.path.join(xlsx_path + "\Activity_Guide_Changes\GaN_cons_and_dates.xlsx")
    os.startfile(xlsx_path)
    os.chdir('main')
    return render_template('index.html')

@app.route('/docx_files', methods=["GET", 'POST'])
def docx_files():
    os.chdir('../')
    docx_path = os.getcwd()
    docx_path = os.path.join(docx_path + "\docs_to_change")
    os.startfile(docx_path)
    os.chdir('main')
    return render_template('index.html')

@app.route('/pdf_files', methods=["GET", 'POST'])
def pdf_files():
    os.chdir('../')
    pdf_path = os.getcwd()
    pdf_path = os.path.join(pdf_path + "\pdf_files")
    os.startfile(pdf_path)
    os.chdir('main')
    return render_template('index.html')   


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
        
        if message == 'ok':
            excecute_main(year, north_consts, north_langs, north_lats, south_consts, south_langs, south_lats, download_check)
            
    return render_template('selections2.html', **context)



@app.errorhandler(404)
def not_found(error):

    return render_template('404.html', error=error)



if __name__ =='__main__':

    ui.run()
    

    
