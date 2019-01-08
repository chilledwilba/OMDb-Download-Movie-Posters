import xlsxwriter
from Main import create_folder


# region Field Objects/ Child Objects | String, Number, ObjectVar, URL, dict
class ParentField:
    def __init__(self, omdb_key, column_width, cell_format, header=None):
        self.OMDbKey = omdb_key
        self.ColumnWidth = column_width
        self.Header = omdb_key if header is None else header
        self.CellFormat = cell_format

    def write_cell(self, worksheet, row_index, column_index, movie):
        data = self.get_data(movie)
        worksheet.write(row_index, column_index, data, self.CellFormat)

    def get_data(self, movie):
        try:
            return movie.OMDbJson[self.OMDbKey]
        except KeyError:
            print('File: {filename} - Key Error: {key} '.format(key=self.OMDbKey, filename=movie.FileName))
            return ''


class StringField(ParentField):
    pass


class NumberField(ParentField):
    def write_cell(self, worksheet, row_index, column_index, movie):
        data = self.get_data(movie)
        try:
            for char in ['$', ',']:
                data = data.replace(char, '')
            worksheet.write_number(row_index, column_index, float(data), self.CellFormat)
        except ValueError:
            worksheet.write(row_index, column_index, data, self.CellFormat)


class ObjectVarField(ParentField):
    def get_data(self, movie):
        return getattr(movie, self.OMDbKey)


class URLField(ParentField):
    def __init__(self, omdb_key, column_width, cell_format, alt_format, url_name='URL', header=None):
        super().__init__(omdb_key, column_width, cell_format, header)
        self.AltFormat = alt_format
        self.UrlName = url_name

    def write_cell(self, worksheet, row_index, column_index, movie):
        data = self.get_data(movie)
        if data == 'N/A':
            worksheet.write(row_index, column_index, data, self.AltFormat)
        else:
            worksheet.write_url(row_index, column_index, data, self.CellFormat, self.UrlName)


class DictField(ParentField):
    def write_cell(self, worksheet, row_index, column_index, movie):
        data = self.get_data(movie)
        string = ''
        for d in data:
            string += '{source} : {value}'.format(source=d['Source'], value=d['Value'])
        worksheet.write(row_index, column_index, string, self.CellFormat)
# endregion


def main(movies, excel_dir, excel_file_name):
    # Check if directory exists / create directory
    create_folder(excel_dir)
    workbook = xlsxwriter.Workbook('{excel_dir}{excel_file_name}'.format(excel_dir=excel_dir,
                                                                         excel_file_name=excel_file_name))
    worksheet = workbook.add_worksheet()

    # Formats
    title_format = workbook.add_format({'font_size': 14})
    url_format = workbook.add_format({'align': 'center', 'font_color': 'blue', 'underline': 1})
    center_format = workbook.add_format({'align': 'center'})
    left_format = workbook.add_format({'align': 'left'})
    bad_format = workbook.add_format({'bg_color': 'red'})
    comma_format = workbook.add_format({'num_format': 0x25, 'align': 'center'})
    money_format = workbook.add_format({'num_format': '$#,###,###', 'align': 'center'})

    # Spreadsheet Headers - displayed in list order
    # Comment out or re-order list objects to your needs
    column_fields = [
        StringField('Title', 55, title_format),
        ObjectVarField('FileName', 95, title_format),
        NumberField('Year', 15, center_format),
        StringField('Rated', 15, center_format),
        StringField('Released', 18, center_format),
        StringField('Runtime', 20, center_format),
        StringField('Genre', 45, center_format),
        StringField('Director', 55, center_format),
        StringField('Writer', 132, center_format),
        StringField('Actors', 70, center_format),
        StringField('Plot', 150, left_format),
        StringField('Language', 50, center_format),
        StringField('Country', 31, center_format),
        StringField('Awards', 65, center_format),
        URLField('Poster', 11, url_format, center_format, 'Poster'),
        DictField('Ratings', 80, center_format),
        NumberField('Metascore', 18, center_format),
        NumberField('imdbRating', 16, center_format),
        NumberField('imdbVotes', 18, comma_format),
        StringField('imdbID', 15, center_format),
        StringField('Type', 15, center_format),
        StringField('DVD', 17, center_format),
        NumberField('BoxOffice', 25, money_format),
        StringField('Production', 32, center_format),
        URLField('Website', 16, url_format, center_format, 'Website')]

    worksheet.set_zoom(70)

    # Set Column Widths
    for i in range(len(column_fields)):
        worksheet.set_column(i, i, column_fields[i].ColumnWidth)

    # Title
    worksheet.write('A1', "MOVIE'S", title_format)
    worksheet.write_formula('B1', '"Total : "&COUNTIF(Table1[Title],"*")', title_format)

    # Create Table
    table_headers = []
    for field in column_fields:
        table_headers.append({'header': field.Header})

    # Start Row, Start Column, End Row, End Column
    worksheet.add_table(1, 0, len(movies) + 1, len(table_headers) - 1,
                        {'style': 'Table Style Medium 1',
                         'banded_columns': True,
                         'banded_rows': False,
                         'columns': table_headers})

    # For each Movie write its field data in a new row, iterate over movies and fields
    row_index = 1
    for movie in movies:
        row_index += 1
        column_index = 0
        if movie.OMDbJson['Response'] == 'True':
            for field in column_fields:
                field.write_cell(worksheet, row_index, column_index, movie)
                column_index += 1
        else:
            # Print FileMovieTitle, FileName as error
            worksheet.write(row_index, 0, movie.FileMovieTitle, bad_format)
            worksheet.write(row_index, 1, movie.FileName, bad_format)

    workbook.close()
