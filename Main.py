import os
from tkinter import filedialog
from tkinter import *
import requests
from MovieObject import MovieObject
import CreateExcelSpreadSheet

# SET API KEY
apiKey = 'ENTER YOUR API KEY'
poster_dir = '_Output/Movie Posters/'
excel_dir = '_Output/'
excel_file_name = 'Movie Stats.xlsx'


def get_main_directory():
    root = Tk()
    root.withdraw()
    main_directory = filedialog.askdirectory()
    return main_directory


def get_list_of_sub_directory_names(path):
    files = os.listdir(path)
    return files


def get_list_of_txt_lines(file):
    if os.path.exists(file):
        return [line.rstrip('\n') for line in open(file)]
    print('{0} doesnt exist'.format(file))


def create_movie_object_list(sub_names):
    movie_list = []
    for name in sub_names:
        movie = MovieObject(name)
        movie_list.append(movie)
    return movie_list


def search_omdb(movie_objects):
    for movie in movie_objects:
        params = dict(
            apikey=apiKey,
            t=movie.FileMovieTitle,
            y=movie.FileYear,
            r='json'
        )
        url = 'http://www.omdbapi.com/?'
        resp = requests.get(url=url, params=params)
        movie.OMDbJson = resp.json()

        print('{resp}: {filename}'.format(resp=get_json_field_data(movie, 'Response'), filename=movie.FileName))


# region Download Movie Posters
def create_folder(dir_name):
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)


def create_file_name(title, year):
    name = '{title} ({year}).jpg'.format(title=title, year=year)
    illegal = ['\',''/', ':', '*', '"', '<', '>', '|', '?']
    for i in illegal:
        name = name.replace(i, '')
    return name


def write_poster_to_file(r, name):
        with open(poster_dir + name, 'wb') as f:
            f.write(r.content)


def amazon_poster_download(movie, name):
    url = get_json_field_data(movie, 'Poster')
    if 'https://' in url:
        r = requests.get(url)
        if r.status_code == 200:
            write_poster_to_file(r, name)
            return True
    return False


def omdb_api_poster_download(movie, name):
    imdb_id = get_json_field_data(movie, 'imdbID')
    params = dict(
        apikey=apiKey,
        i=imdb_id
    )
    url = ' http://img.omdbapi.com/?'
    r = requests.get(url, params)
    if r.status_code == 200:
        write_poster_to_file(r, name)
        return True
    return False


'''
Tries to download Poster using given Amazon Poster URL
if unsuccessful tries to download poster using omdb api
'''
def download_posters(movie_objects):
    # Create Folder to store movie Posters
    create_folder(poster_dir)

    for movie in movie_objects:
        status = False
        if movie.OMDbJson['Response'] == 'True':
            title = get_json_field_data(movie, 'Title')
            year = get_json_field_data(movie, 'Year')
            name = create_file_name(title, year)  # Create valid file name, removes illegal characters

            status = amazon_poster_download(movie, name)

            if status is False:
                status = omdb_api_poster_download(movie, name)

        print('{status}: {filename}'.format(status=status, filename=movie.FileName))
# endregion


def get_json_field_data(movie, key):
    try:
        return movie.OMDbJson[key]
    except KeyError:
        print('Cant find key:{key} in json file: {filename}'.format(key=key, filename=movie.FileName))
        return None


if __name__ == "__main__":
    option = 1
    names_list = []

    # Search Folder
    # Using tkinter to open file dialog and select parent directory
    if option == 1:
        mainDirectory = get_main_directory()
        if mainDirectory != '':
            names_list = get_list_of_sub_directory_names(mainDirectory)

    # Reads in txt file to test file names
    if option == 2:
        names_list = get_list_of_txt_lines('Test-Movie-Names.txt')

    if names_list:
        movieObjects = create_movie_object_list(names_list)

        # Print out movies - title, year, filename
        print('READ IN MOVIES...')
        for m in movieObjects:
            print(str(m))

        # Search OMDb fill objects with data
        print('SEARCH OMDB...')
        search_omdb(movieObjects)

        # Create Excel Spread Sheet
        print('CREATE EXCEL SPREADSHEET...')
        CreateExcelSpreadSheet.main(movieObjects, excel_dir, excel_file_name)

        # Download Posters
        print('DOWNLOAD POSTERS...')
        download_posters(movieObjects)
