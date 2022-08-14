import os
from docx2pdf import convert
'''
northYear = 2022
north_constellations = ["Perseus", "Leo", "Bootes", "Cygnus", "Pegasus", "Orion", "Hercules"]
north_languages = ["Catalan", "Chinese", "Czech", "English", "Finnish", "French", "Galician", "German", "Greek", "Indonesian", "Japanese", "Polish", "Portuguese", "Romanian", "Serbian", "Slovak", "Slovenian", "Spanish", "Swedish", "Thai"]
latitudes_north = ["50N", "40N", "30N", "20N", "10N", "0"]
# Creating the directories and the Paths for North Constellations
northDirectories= agc.createNorthDir(northYear, northConstellations)
northPaths = agc.createNorthPaths(northDirectories, northLanguages)


# Get the data from the User for south constellations
southYear = northYear
south_constellations = ["Orion","Canis Major", "Crux", "Leo", "Bootes", "Scorpius", "Hercules", "Sagittarius", "Grus", "Pegasus"]
south_languages = ["English", "French", "Indonesian", "Portuguese", "Spanish"]
latitudes_south = ["0", "10S", "20S", "30S", "40S"]
'''

north_word_paths = ['C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_50N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_30N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_50N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_30N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_50N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_30N_chinese (traditional).docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_50N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_30N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_50N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_30N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_50N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_30N_Czech.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_50N_English.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_30N_English.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_50N_English.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Leo\\GaN_2022_ActivityGuide_Leo_lat_30N_English.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_50N_English.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_30N_English.docx'] 
south_word_paths = ['C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_0_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_40S_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Canis Major\\GaN_2022_ActivityGuide_Canis Major_lat_0_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Canis Major\\GaN_2022_ActivityGuide_Canis Major_lat_40S_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Bootes\\GaN_2022_ActivityGuide_Bootes_lat_0_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Bootes\\GaN_2022_ActivityGuide_Bootes_lat_40S_Portuguese.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_0_Spanish.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Orion\\GaN_2022_ActivityGuide_Orion_lat_40S_Spanish.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Canis Major\\GaN_2022_ActivityGuide_Canis Major_lat_0_Spanish.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Canis Major\\GaN_2022_ActivityGuide_Canis Major_lat_40S_Spanish.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Bootes\\GaN_2022_ActivityGuide_Bootes_lat_0_Spanish.docx', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_South_2022_ActivityGuide_Bootes\\GaN_2022_ActivityGuide_Bootes_lat_40S_Spanish.docx']

north_pdf_folders = ['C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_North_2022_ActivityGuide_Perseus', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_North_2022_ActivityGuide_Leo', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_North_2022_ActivityGuide_Orion']
south_pdf_folders =['C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_South_2022_ActivityGuide_Orion', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_South_2022_ActivityGuide_Canis Major', 'C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220719_150959\\GaN_South_2022_ActivityGuide_Bootes']


#path = "C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_50N_chinese (traditional).docx"
#path2 = "C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220718_220633\\GaN_North_2022_ActivityGuide_Perseus"
north_pdf_paths = []
for pdf_folder in north_pdf_folders:
    for word_path in north_word_paths:
        constellation = word_path.split('_')[-4]
        if constellation in pdf_folder:
            word_path1 = word_path[:-5]
            word_path1 = word_path1.split('\\')[-1]

            pdf_path = os.path.join(pdf_folder + "\\" + word_path1 +".pdf" )

            north_pdf_paths.append(pdf_path)

print (north_pdf_paths)









import sys
import os
import comtypes.client

path = "C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\docs_changed\\GaN_North_2022_ActivityGuide_Perseus\\GaN_2022_ActivityGuide_Perseus_lat_50N_chinese (traditional).docx"
path2 = "C:\\Users\\Marco Moreno\\OneDrive\\Documentos\\Enciso Systems\\GaN\\GaN\\pdf_files\\Activity_Guides_20220718_220633\\GaN_North_2022_ActivityGuide_Perseus"

path1 = path[:-5]
path1 = path1.split('\\')[-1]

path2 = os.path.join(path2 + "\\" + path1 +".pdf" )

print (path2)

if path.endswith(".docx"):
    wdFormatPDF = 17

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(path)
    doc.SaveAs(path2, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def docx2pdf(input_file):
    """docx to pdf"""
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(input_file.replace(".docx", ".pdf"), FileFormat=17)
    doc.Close()
    word.Quit()

                const_folder = folder.split('_')[-1]
            cardinal_folder = folder.split('_')[-4]

            const_path =  paths.split('_')[-4]
            cardinal_path = path.split('_')[-4]

            print(const_folder, cardinal_folder)
            print(const_path, cardinal_path)

        time.sleep(10)
        new_path = path.replace(".docx", "1.docx")
        shutil.copyfile(path, new_path)
        pdf_path = new_path.replace(".docx", ".pdf")
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(path)
        doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        print("renamed and printed:  " + pdf_path)




        @app.route('/selections', methods=["GET", 'POST'])
def selections():
    selections_form = SelectionsForm()
    context = {
        'north_constellations' : north_consts,
        'north_languages' : north_langs,
        'latitudes_north' : north_lats,
        'south_constellations' :south_consts,
        'south_languages' :south_langs,
        'latitudes_south' : south_lats,
        'selections_form': selections_form
        }

    if request.method == 'POST':
        year= selections_form.year.data
        print(year)

    return render_template('selections.html', **context)
