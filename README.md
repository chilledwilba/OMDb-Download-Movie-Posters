# OMDb-Download-Movie-Posters
This program reads in all file names of a given directory and searches OMDb for its corresponding movie data. Minimal file name scraping is in place to determine movie names given long file names.

<b>Example:</b><br>
File name: Zodiac.Signs.of.the.Apocalypse.2014.1080p.BluRay.H264.AAC.5.1.BADASSMEDIA<br>
Title: Zodiac Signs of the Apocalypse<br>
Year: 2014
<br>

<b>Spreadsheet</b><br>
XlsxWriter is used to output OMDb information to an Excel Spreadsheet. The header fields can be edited inside of the `CreateExcelSpreadSheet.py` file.<br>
![excel-image](https://github.com/chilledwilba/OMDb-Download-Movie-Posters/blob/master/images/excel-image.PNG)

<b>Download Posters</b><br>
Two methods are used to download the movie posters:
1. Using the fetched movie poster url
2. OMDb Poster API

![posters-image](https://github.com/chilledwilba/OMDb-Download-Movie-Posters/blob/master/images/Posters-image'.PNG)

## Dependencies

Requires [OMDb API Key](http://www.omdbapi.com/apikey.aspx)

Python version 3.7.1 or newer

Modules
* requests
* XlsxWriter

## Running The Project
How to run the project in a virtual environment

1. [Download & install python ](https://www.python.org/)
2. Open CMD and run: pip install virtualenv
3. Open CMD inside of project path
4. Run virtualenv venv
5. venv\Scripts\activate
6. pip install XlsxWriter
7. pip install requests
8. python Main.py

![venv-cmd-gif](https://github.com/chilledwilba/OMDb-Download-Movie-Posters/blob/master/images/venv-cmd-gif.gif)



