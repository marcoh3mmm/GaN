# -*- coding:utf-8 -*-
import time
import os
import sys
import multiprocessing

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__),os.pardir))
sys.path.append(PROJECT_ROOT)
import Activity_Guide_Changes as agc



if __name__ =='__main__':

    # Start time counter
    start = time.time()

    numProcess = 8
    images_downloaded = False
    # Get the data from the User for north constellations
    north_year = 2022
    north_constellations = ["Perseus", "Leo", "Bootes"]
    north_languages = ["Catalan", "Chinese", "Czech","French"]
    latitudes_north = ["50N", "40N", "30N", "20N", "10N"]

    # Get the data from the User for south constellations
    south_year = north_year
    south_constellations = ["Canis Major", "Crux"]
    south_languages = ["French", "Indonesian", "Portuguese", "Spanish"]
    latitudes_south = ["0", "10S", "20S", "30S", "40S"]
 
    
    # Create the folders for the pdf files
    pdf_folder = agc.create_pdf_folder()
    pdf_north_folders = agc.create_pdf_dir("North", north_year, north_constellations,pdf_folder)
    pdf_south_folders = agc.create_pdf_dir("South", south_year, south_constellations,pdf_folder)

    #create the paths for the pdf files
    north_paths = agc.createNorthPaths(pdf_north_folders, north_languages, latitudes_north)
    south_paths = agc.createSouthPaths(pdf_south_folders, south_languages, latitudes_south)

    
    ############################## translations ##########################################
     
    if len(north_constellations) == 0:
        print("There are not constellations selected for the north hemisphere.")
        pass
    else:
        # Create a list from the new doc Paths for a leter use in the Images printing
        north_list_paths = []
        #Call de translation for north constellations function, requiring multiprocessing with Pool
        pool_1 = multiprocessing.Pool(processes = numProcess)
        for path in north_paths:
            north_list_paths.append(pool_1.apply_async(agc.northTranslation, args = (path, )).get())
        pool_1.close()
        pool_1.join()
    
    
    # Create a list from the new doc Paths for a leter use in the Images printing
    if len(south_constellations) == 0:
        print("There are not constellations selected for the south hemisphere.")
        pass
    else:
        #Call de translation for north constellations function, requiring multiprocessing with Pool
        south_list_paths = []
        pool_2 = multiprocessing.Pool(processes = numProcess)
        for path in south_paths:
            south_list_paths.append(pool_2.apply_async(agc.southTranslation, args = (path, )).get())
        pool_2.close()
        pool_2.join()

    ################################ Print Charts ##########################################
    if images_downloaded == True:
    
    ################################ Download the Imagees ##########################################
        
        # Activate the Scrapper
        agc.createImageDir()
        linksNorth = agc.imagesLinks(north_constellations,latitudes_north)
        linksSouth = agc.imagesLinks(south_constellations,latitudes_south)
        
        if len(north_constellations) == 0:
            print("There are not constellations selected for the north hemisphere.")
            pass
        else: 
            pool3 = multiprocessing.Pool(processes = numProcess)
            for link in linksNorth:
                pool3.apply_async(agc.imageDownload, args = (link, ))
            pool3.close()
            pool3.join()

        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            pass
        else:
            pool4 = multiprocessing.Pool(processes = numProcess)
            for link in linksSouth:
                pool4.apply_async(agc.imageDownload, args = (link, ))
            pool4.close()
            pool4.join()

        ############################# Print Charts Downlaoded in the Word files #####################################

        if len(north_constellations) == 0:
            print("There are not constellations selected for the north hemisphere.")
            north_docx_paths = []
            pass
        else:
            pool5 = multiprocessing.Pool(processes = numProcess)
            north_docx_paths = []
            for path in north_list_paths:
                north_docx_paths.append(pool5.apply_async(agc.printDownloadImage, args = (path, )).get())
            pool5.close()
            pool5.join()
        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            south_docx_paths = []
            pass
        else:
            pool6 = multiprocessing.Pool(processes = numProcess)
            south_docx_paths = []
            for path in south_list_paths:
                south_docx_paths.append(pool6.apply_async(agc.printDownloadImage, args = (path, )).get())
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
            pool5 = multiprocessing.Pool(processes = numProcess)
            north_docx_paths = []
            for path in north_list_paths:
                north_docx_paths.append(pool5.apply_async(agc.printLocalImage, args = (path, )).get())
            pool5.close()
            pool5.join()
        
        if len(south_constellations) == 0:
            print("There are not constellations selected for the south hemisphere.")
            south_docx_paths = []
            pass
        else:
            pool6 = multiprocessing.Pool(processes = numProcess)
            south_docx_paths = []
            for path in south_list_paths:
                south_docx_paths.append(pool6.apply_async(agc.printLocalImage, args = (path, )).get())
            pool6.close()
            pool6.join()


    '''
    # Get the saving pdf paths
    # Create the dictionaries with the docx paths and the pdf paths
    if len(north_constellations) == 0:
        print("There are not constellations selected for the north hemisphere.")
        pass
    else:
        north_pdf_paths = agc.create_pdf_paths(north_docx_paths, pdf_north_folders)
        #north_path_dict = dict(zip(north_docx_paths, north_pdf_paths))
    
    if len(south_constellations) == 0:
        print("There are not constellations selected for the south hemisphere.")
        pass
    else:
        south_pdf_paths = agc.create_pdf_paths(south_docx_Paths, pdf_south_folders)
        #south_path_dict = dict(zip(south_docx_Paths, south_pdf_paths))
    '''

    
    '''
    ################### Convert .docx into p.pdf and save it in a new folder ####################
    if len(north_constellations) == 0:
        print("There are not constellations selected for the north hemisphere.")
        pass
    else:
        north_pdf_paths = []
        for path in north_docx_paths:
            north_pdf_paths.append(agc.print_pdf(path))
        



    if len(south_constellations) == 0:
        print("There are not constellations selected for the north hemisphere.")
        pass
    else:
        south_pdf_paths = []
        for path in south_docx_paths:
            south_pdf_paths.append(agc.print_pdf(path))
    '''
    

    if len(north_docx_paths) == 0: 
        total_docx_paths = south_docx_paths
    elif len(south_docx_paths) == 0: 
        total_docx_paths = north_docx_paths
    else:
        total_docx_paths = north_docx_paths + south_docx_paths 


    pool9 = multiprocessing.Pool(processes = numProcess)
    results = pool9.map_async(agc.print_pdf, total_docx_paths)
    results.get()

    agc.remove_docs(total_docx_paths)
    
    
    # Finishing time counter and getting time of execution
    finish = time.time() - start
    print('Execution time: ', time.strftime("%H:%M:%S", time.gmtime(finish)))