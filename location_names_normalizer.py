#!C:\Python27
# -*- coding: utf-8 -*-
import sys
from openpyxl import load_workbook
from fuzzywuzzy import fuzz, process
from utils import get_unicode, write_ws


class Location(list):

    def __init__(self, id=None, values=[]):
        self.id = id
        self.extend(values)


class BaseLocations():

    def iterate_locations_sheet(self, wb, index=0):
        """Iterates a locations sheet yielding a Location object with id
        and values stored in it."""

        # load worksheet
        ws = wb.worksheets[index]

        # count hierarchical values after id
        num_values = self._count_fields(ws)

        # iterate through locations in worksheet
        i = 1
        while ws.cell(row=i, column=0).value:

            # take id
            id = ws.cell(row=i, column=0).value

            # take values
            values = []
            j = 1
            while j < num_values:
                values.append(ws.cell(row=i, column=j).value)
                j += 1

            # create location
            location = Location(id, values)

            i += 1

            # if location has values, yield it
            if len(values) > 0:
                yield location

    def _count_fields(self, ws):
        """Count number of fields in worksheet"""

        field_number = 1
        while ws.cell(row=0, column=field_number).value:
            field_number += 1

        return field_number


class NormalizedLocationsList(list, BaseLocations):

    def add(self, location, location_matched):
        """Adds normalized location based on location_matched way to write it,
        with its correspondence ids from both."""

        normalized_location = [location.id, location_matched.id]
        normalized_location.extend(location)
        normalized_location.extend(location_matched)
        print normalized_location
        self.append(normalized_location)


class LocationsList(list, BaseLocations):

    def __init__(self, wb):
        self._create_first_list(wb)

    # PRIVATE
    def _create_first_list(self, wb):
        """Creates a list of Locations from first worksheet of workbook."""

        # iterate through locations in first sheet and append to a list
        for location in self.iterate_locations_sheet(wb, 0):
            self.append(location)


class LocationsDict(dict, BaseLocations):

    def __init__(self, wb):
        self._create_second_list(wb)

    # PUBLIC
    def find(self, location, dictionary=None, key_index=0):
        """Returns the most similar location stored in locations dict to
        location passed into the function."""

        # dictionary is self object if no dictionary is provided
        if not dictionary:
            dictionary = self

        # take first value field to be found
        value = location[key_index]

        # extract matched value from
        value_matched = process.extractOne(value, dictionary.keys())

        if value_matched:
            key = value_matched[0]

            # if there are more values to evaluate, call recursively
            if len(location) > key_index + 1:
                return self.find(location, dictionary[key], key_index + 1)

            else:
                return dictionary[key]

        else:
            return None

    # PRIVATE
    def _create_second_list(self, wb):
        """Creates a dict with hierarchical keys to store locations from
        second worksheet of workbook."""

        # iterate through locations in second sheet and add to dictionary
        for location in self.iterate_locations_sheet(wb, 1):
            self._add_location(self, location, 0)

    def _add_location(self, dictionary, location, key_index):
        """Adds location using hierarchical dictionaries and storing id as
        the value in the last dictionary."""

        # takes value of key_index location field
        key_value = location[key_index]

        # if is the last value, add key_value with id, as value
        if len(location) == key_index + 1:
            dictionary[key_value] = location

        # else, add key to dictionary and recursively move on to next key_value
        else:

            # if key is not already, add a dictionary for it
            if key_value not in dictionary:
                dictionary[key_value] = {}

            # recursively add location with next key_index inside subdictionary
            self._add_location(dictionary[key_value], location, key_index + 1)


class LocationsFile():
    """Parse the location files into location lists and creates normalized
    list between two location lists."""

    def __init__(self, wb):
        self.wb = wb
        self.first_list = LocationsList(wb)
        self.second_list = LocationsDict(wb)
        self.normalized_list = NormalizedLocationsList()

    # PUBLIC
    def normalize_names(self):
        """Looks in second_list for each location in first_list, when there is
        a match, adds to normalized list."""

        # iterate through all first_list locations
        i = 1
        for location in self.first_list:

            print "Looking number", i, "of", len(self.first_list),
            location_matched = self.second_list.find(location)

            # if there is a match, add to normalized list
            if location_matched:

                print "added"
                self.normalized_list.add(location, location_matched)

            else:
                print "no match"

            i += 1

    def count_first_list(self):
        return self.first_list.count()

    def count_second_list(self):
        return self.second_list.count()

    def save(self):
        """Add normalized list to workbook and save it."""

        ws = self.wb.create_sheet(-1, "tbl_corresponde")

        for record in self.normalized_list:
            write_ws(ws, record)

        self.wb.save("location_names_normalized.xlsx")


# DATA
INPUT_FILE = "location_names.xlsx"
OUTPUT_FILE = "location_names.xlsx"


# USER METHODS
def normalize_location_names(input_file=None, output_file=None):
    """Takes a workbook with a list of locations at first two sheets and
    add a sheet with correspondence table between location names at them
    two lists."""

    # if not i/o files are passed, defaults are used
    input_file = input_file or INPUT_FILE
    output_file = output_file or OUTPUT_FILE

    # loads excel file
    wb = load_workbook(input_file)

    # creates a location names file object
    lf = LocationsFile(wb)

    # count records in each list
    print "First list has", lf.count_first_list(), "records"
    print "Second list has", lf.count_second_list(), "records"

    # creates normalized list of names
    lf.normalize_names()

    # count records in normalized list
    print "Normalized list has", lf.count_normalized_list(), "records"

    # save excel workbook with normalized sheet appended
    lf.save(output_file)


# executes main routine
if __name__ == '__main__':

    # if parameters are passed, use them
    input_file = None
    output_file = None

    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]

    # main call
    normalize_location_names(input_file, output_file)
