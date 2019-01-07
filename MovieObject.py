import re


class MovieObject:
    def __init__(self, name):
        self.OMDbJson = None
        self.FileName = name

        # TODO: Fix removing '.' for acronyms eg The Man for U.N.C.L.E.
        # Remove Elements , [ ] { } ( ) - _
        name = self.remove_elements(name, [',', '[', ']', '{', '}', '(', ')', '-', '_', '.'])

        # Set File Year
        self.FileYear = self.find_year(name)

        # Set File Movie Title
        self.FileMovieTitle = self.find_movie_title(name, ['1080p', '720p', '480p'])

    def find_year(self, name):
        year = None

        for i in range(0, len(name) - 3):
            if self.check_next_four_numbers_are_digits(name, i):
                tmp = name[i: i + 4]
                if int(tmp) > 1800:
                    year = tmp
        return year

    def find_movie_title(self, name, split_elements):
        # Create string of elements to split name at eg '1080p |720p |480p |2008'
        split_string = split_elements[0]
        for i in range(1, len(split_elements)):
            split_string += ' |{element}'.format(element=split_elements[i])
        if self.FileYear is not None:
            split_string += ' |{file_year}'.format(file_year=self.FileYear)

        # Split name into a list
        tmp = re.split(split_string, name)

        # Checks for whitespace in case year is before movie name eg. "(2008) Iron Man"
        for x in tmp:
            x = x.strip()
            if len(x) > 0:
                return x

    @staticmethod
    def check_next_four_numbers_are_digits(name, index):
        for i in range(0, 4):
            if not name[index + i].isdigit():
                return False
        return True

    @staticmethod
    def remove_elements(name, elements):
        for ch in elements:
            name = name.replace(ch, ' ')
        return name

    def __str__(self):
        return '{file_movie_title} - {file_year} : {file_name}'.format(file_movie_title=self.FileMovieTitle,
                                                                       file_year=self.FileYear, file_name=self.FileName)
