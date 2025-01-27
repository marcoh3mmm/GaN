import os
import time
import multiprocessing
import Activity_Guide_Changes as agc
from flask import redirect, flash, url_for

def flash_mess(north_consts, north_langs, north_lats, south_consts, south_langs, south_lats):
    if len(north_consts) > 0 and len(north_langs) > 0 and len(north_lats) > 0 and len(south_consts) > 0 and len(south_langs) > 0 and len(south_lats) > 0:
        return 'ok'

    elif len(north_consts) > 0 and len(north_langs) > 0 and len(north_lats) > 0 and len(south_consts) == 0 and len(south_langs) == 0 and len(south_lats) == 0:
        return 'ok'

    elif len(north_consts) == 0 and len(north_langs) == 0 and len(north_lats) == 0 and len(south_consts) > 0 and len(south_langs) > 0 and len(south_lats) > 0:
        return 'ok'
    
    elif len(north_consts) == 0 and len(north_langs) == 0 and len(north_lats)==0 and len(south_consts) == 0 and len(south_langs) == 0 and len(south_lats)==0:
        flash('Please select a North Constellation or a South Constellation.')
        return redirect(url_for('selections2'))
    
    else:
        flash('To get an Activity guide, you must select at least one Constellation, one Language, and one Latitude simultaneously, in the North or the South constellations')
        return redirect(url_for('selections2'))

    

def excecute_main(year, north_consts, north_langs, north_lats, south_consts, south_langs, south_lats, download_check):
    # Start time counter
            start = time.time()

            #Stablish number of process for multiprocessing
            num_process = 8

            # Activate the Scrapper
            if download_check == None:
                images_downloaded = False
            else:
                images_downloaded = True

            # Get the data from the User for north constellations
            
            north_year = year
            north_constellations = north_consts
            north_languages = north_langs
            latitudes_north = north_lats

            # Get the data from the User for south constellations
            south_year = north_year
            south_constellations = south_consts
            south_languages = south_langs
            latitudes_south = south_lats

            
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

            #os.chdir('main')
            
            # Finishing time counter and getting time of execution
            finish = time.time() - start
            print('Execution time: ', time.strftime("%H:%M:%S", time.gmtime(finish)))

            flash("The activity guides are available for use.")
            os.startfile(pdf_folder)
            