import os 
from datetime import datetime

def create_pdf_folder():

    current_day_hour = datetime.now()

    savePath = os.getcwd()
    savePath = os.path.join(savePath + "\GaN\pdf_files\Activity_Guides_" + current_day_hour.strftime('%Y%m%d_%H%M%S'))
    if not os.path.exists(savePath):
        os.makedirs(savePath)
    return savePath

def create_pdf_dir(cardinal, year, constellations, pdf_folder):
    
    paths = []
    for con in constellations:
        savePath = os.getcwd() 
        savePath = os.path.join(pdf_folder + "\GaN_{cardinal}_{year}_ActivityGuide_{con}".format(cardinal = cardinal, year = year, con = con))        
        print(savePath)
        os.mkdir(savePath)
        paths.append(savePath)

    return paths




