# -*- coding:utf-8 -*-
import time
import os
import sys
import multiprocessing
from flask import Flask, request, make_response, redirect, render_template
#from flaskwebgui import FlaskUI #get the FlaskUI class
from flask_wtf import FlaskForm
from wtforms import BooleanField,SubmitField, SelectMultipleField,SelectField,widgets
from wtforms.validators import DataRequired



PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__),os.pardir))
sys.path.append(PROJECT_ROOT)

from app import create_app

#import Activity_Guide_Changes as agc

app = create_app()


north_constellations = ["Select All", "Perseus", "Leo", "Bootes", "Cygnus", "Pegasus", "Orion", "Hercules"]
north_languages= ["Select All", "Catalan", "Chinese", "Czech", "English", "Finnish", "French", "Galician", "German", "Greek", "Indonesian", "Japanese", "Polish", "Portuguese", "Romanian", "Serbian", "Slovak", "Slovenian", "Spanish", "Swedish", "Thai"]
latitudes_north = ["Select All", "0", "10N", "20N", "30N", "40N", "50N"]

south_constellations = ["Select All", "Orion","Canis Major", "Crux", "Leo", "Bootes", "Scorpius", "Hercules", "Sagittarius", "Grus", "Pegasus"]
south_languages = ["Select All", "English", "French", "Indonesian", "Portuguese", "Spanish"]
latitudes_south = ["Select All", "0", "10S", "20S", "30S", "40S"]


# Feed it the flask app instance 
#ui = FlaskUI(app,width=1500,height=800)


def select_multi_checkbox(field, ul_class='', **kwargs):
    kwargs.setdefault('type', 'checkbox')
    field_id = kwargs.pop('id', field.id)
    html = ['<ul %s>' % widgets.html_params(id=field_id, class_=ul_class)]
    for value, label, checked in field.iter_choices():
        choice_id = '%s-%s' % (field_id, value)
        options = dict(kwargs, name=field.name, value=value, id=choice_id)
        if checked:
            options['checked'] = 'checked'
        html.append('<li><input %s /> ' % widgets.html_params(**options))
        html.append('<label for="%s">%s</label></li>' % (field_id, label))
    html.append('</ul>')
    return ''.join(html)



class SelectionsForm(FlaskForm):
    years=[2020,2021,2022,2023,2024,2025]
    year=SelectField('Select Year',choices=(years))
    north_constellations = SelectMultipleField('Constellations',choices=(north_constellations), widget=select_multi_checkbox)
    north_languages = SelectMultipleField('Languagues',choices=(north_languages), widget=select_multi_checkbox)
    north_latitudes = SelectMultipleField('Latitudes',choices=(latitudes_north), widget=select_multi_checkbox)

    south_constellations = SelectMultipleField('Constellations',choices=(south_constellations), widget=select_multi_checkbox)
    south_languages = SelectMultipleField('Languagues',choices=(south_languages), widget=select_multi_checkbox)
    south_latitudes = SelectMultipleField('Latitudes',choices=(latitudes_south), widget=select_multi_checkbox)
    
    download_charts=BooleanField('Download Charts')
    select_everything=BooleanField('Select All')
    submit =SubmitField('Submit')



# do your logic as usual in Flask
@app.route('/', methods=["GET", 'POST'])
def index():
    home=make_response(redirect('/index'))
    generate_activity_guides = make_response(redirect('/selections'))
    return render_template('index.html')

@app.route('/selections', methods=["GET", 'POST'])
def selections():
    context = {
        'north_constellations' : north_constellations,
        'north_languages' : north_languages,
        'latitudes_north' : latitudes_north,
        'south_constellations' :south_constellations,
        'south_languages' :south_languages,
        'latitudes_south' : latitudes_south,
        }

    return render_template('selections.html', **context)

@app.route('/selections2', methods=["GET", 'POST'])
def selections2():
    selections_form = SelectionsForm()
    context = {
        'selections_form': selections_form
        }

    return render_template('selections2.html', **context)



@app.errorhandler(404)
def not_found(error):

    return render_template('404.html', error=error)



if __name__ =='__main__':

    app.run()

'''

    # Start time counter
    start = time.time()


    num_process = 8
    images_downloaded = False
    # Get the data from the User for north constellations
    north_year = 2022
    north_constellations = ["Perseus", "Leo", "Bootes", "Cygnus", "Pegasus", "Orion", "Hercules"]
    north_languages = ["Catalan", "Chinese", "Czech", "English", "Finnish", "French", "Galician", "German", "Greek", "Indonesian", "Japanese", "Polish", "Portuguese", "Romanian", "Serbian", "Slovak", "Slovenian", "Spanish", "Swedish", "Thai"]
    latitudes_north = ["50N", "40N", "30N", "20N", "10N", "0"]

    # Get the data from the User for south constellations
    south_year = north_year
    south_constellations = ["Orion","Canis Major", "Crux", "Leo", "Bootes", "Scorpius", "Hercules", "Sagittarius", "Grus", "Pegasus"]
    south_languages = ["English", "French", "Indonesian", "Portuguese", "Spanish"]
    latitudes_south = ["0", "10S", "20S", "30S", "40S"]
 
    
    # Create the folders for the pdf files
    pdf_folder = agc.create_pdf_folder()
    pdf_north_folders = agc.create_pdf_dir("North", north_year, north_constellations,pdf_folder)
    pdf_south_folders = agc.create_pdf_dir("South", south_year, south_constellations,pdf_folder)

    #create the paths for the word files
    north_paths = agc.create_north_paths(pdf_north_folders, north_languages, latitudes_north)
    south_paths = agc.create_south_paths(pdf_south_folders, south_languages, latitudes_south)

    
    ############################## translations ##########################################
     
    if len(north_constellations) == 0:
        print("There are not constellations selected for the north hemisphere.")
        pass
    else:
        # Create a list from the new doc Paths for a leter use in the Images printing
        north_list_paths = []
        #Call de translation for north constellations function, requiring multiprocessing with Pool
        pool_1 = multiprocessing.Pool(processes = num_process)
        for path in north_paths:
            north_list_paths.append(pool_1.apply_async(agc.north_translations, args = (path, )).get())
        pool_1.close()
        pool_1.join()
    
    
    # Create a list from the new doc Paths for a leter use in the Images printing
    if len(south_constellations) == 0:
        print("There are not constellations selected for the south hemisphere.")
        pass
    else:
        #Call de translation for north constellations function, requiring multiprocessing with Pool
        south_list_paths = []
        pool_2 = multiprocessing.Pool(processes = num_process)
        for path in south_paths:
            south_list_paths.append(pool_2.apply_async(agc.south_translations, args = (path, )).get())
        pool_2.close()
        pool_2.join()

    ####################################### Print Charts ##########################################
    if images_downloaded == True:
    
    ################################## Download the Images #########################################
        
        # Activate the Scrapper
        agc.create_image_dir()
        links_north = agc.images_links(north_constellations,latitudes_north)
        links_south = agc.images_links(south_constellations,latitudes_south)
        
        if len(north_constellations) == 0:
            print("There are not constellations selected for the north hemisphere.")
            pass
        else: 
            pool3 = multiprocessing.Pool(processes = num_process)
            for link in links_north:
                pool3.apply_async(agc.images_download, args = (link, ))
            pool3.close()
            pool3.join()

        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            pass
        else:
            pool4 = multiprocessing.Pool(processes = num_process)
            for link in links_south:
                pool4.apply_async(agc.images_download, args = (link, ))
            pool4.close()
            pool4.join()

        ############################# Print Charts Downlaoded in the Word files #####################################

        if len(north_constellations) == 0:
            print("There are not constellations selected for the north hemisphere.")
            north_docx_paths = []
            pass
        else:
            pool5 = multiprocessing.Pool(processes = num_process)
            north_docx_paths = []
            for path in north_list_paths:
                north_docx_paths.append(pool5.apply_async(agc.print_download_image, args = (path, )).get())
            pool5.close()
            pool5.join()
        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            south_docx_paths = []
            pass
        else:
            pool6 = multiprocessing.Pool(processes = num_process)
            south_docx_paths = []
            for path in south_list_paths:
                south_docx_paths.append(pool6.apply_async(agc.print_download_image, args = (path, )).get())
            pool6.close()
            pool6.join()


    ################################ work with local images ##########################################
    ############################# Print local Charts on the Word files #####################################
    else:
   
        if len(north_constellations) == 0:
            print("There are not constellations selected for the north hemisphere.")
            north_docx_paths = []
            pass
        else:
            pool5 = multiprocessing.Pool(processes = num_process)
            north_docx_paths = []
            for path in north_list_paths:
                north_docx_paths.append(pool5.apply_async(agc.print_local_image, args = (path, )).get())
            pool5.close()
            pool5.join()
        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            south_docx_paths = []
            pass
        else:
            pool6 = multiprocessing.Pool(processes = num_process)
            south_docx_paths = []
            for path in south_list_paths:
                south_docx_paths.append(pool6.apply_async(agc.print_local_image, args = (path, )).get())
            pool6.close()
            pool6.join()

    ############################# convert .docx to pdf #####################################
    # Creating a List with the word Paths
    if len(north_docx_paths) == 0: 
        total_docx_paths = south_docx_paths
    elif len(south_docx_paths) == 0: 
        total_docx_paths = north_docx_paths
    else:
        total_docx_paths = north_docx_paths + south_docx_paths 

    # converting .docx to pdf
    pool9 = multiprocessing.Pool(processes = num_process)
    results = pool9.map_async(agc.print_pdf, total_docx_paths)
    results.get()

    # Delete the word files
    agc.remove_docs(total_docx_paths)
    
    
    # Finishing time counter and getting time of execution
    finish = time.time() - start
    print('Execution time: ', time.strftime("%H:%M:%S", time.gmtime(finish)))
    '''