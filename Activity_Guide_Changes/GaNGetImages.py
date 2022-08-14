from bs4 import BeautifulSoup
import requests
import os
from shutil import rmtree

# Create a new folder to download the images
def create_image_dir():
    save_path = os.getcwd() 
    save_path = os.path.join(save_path + "\images")
    rmtree(save_path)
    os.mkdir(save_path)
    return save_path


# Get the latitudes according to the images names in the GaN webpage (north ="10", south ="10s")
def transform_latitude(lat):
    if "N" in lat:
        lat = str(lat.rstrip(lat[-1]))
    else:
        lat = str(lat.lower())
    return lat

def images_links(constellations, latitudes):
        # Get the url in GaN website where the images are located
    url = 'https://www.globeatnight.org/magcharts'
    gan = requests.get(url)
    
    # get the soup to pass the content
    soup = BeautifulSoup(gan.text, 'html.parser')
    # Searching the "div with id = finder" where are the images links to get the first Link
    image = soup.find('div' , attrs= {"id" : "finder"}).find('img')
    #Get the link from the image
    image_first_link = str(image['src'])

    images_links = []
    magnitudes = ["05", "15", "25", "35", "45", "55", "65", "75"]
    # Replace the Constellation names in the North and the latitudes in the imageLink
    for const in constellations:
        for lat in latitudes:
            for mag in magnitudes:
                new_link = image_first_link.replace("hercules", const.lower().replace(" ", "-")).replace("10", transform_latitude(lat).replace("05", mag)).replace("05", mag)
                images_links.append(new_link)
    return images_links       


# Download the images and save them in the images folder
def images_download(link):
    #Change the path to the new folder
    new_path = os.getcwd()
    new_path = os.path.join(new_path + "\images\\")

    # Verify the status code of the link
    if requests.get(link).status_code == 200:
    # Save the images in a local folder for a later use with an easier name
        link_string = str(link.replace('https://www.globeatnight.org/img/2021/', '').replace('/day/600/','-'))
        new_link_string = os.path.join(new_path + link_string)
        with open(new_link_string, 'wb') as f:
            img = requests.get(link)
            f.write(img.content)
                        
            return print(link_string + " has been downloaded.\n____________________________________________________________________________________________\n")






