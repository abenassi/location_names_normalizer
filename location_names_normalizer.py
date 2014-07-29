#!C:\Python27
# -*- coding: utf-8 -*-
import sys
from openpyxl import load_workbook
from utils import write_ws
from location_lists import LocationsList, LocationsDict, NormalizedLocationsList


class LocationsFile():
    """Parse the location files into location lists and creates normalized
    list between two location lists."""

    # DATA
    INPUT_FILE = "location_names.xlsx"
    OUTPUT_FILE = "location_names_normalized.xlsx"

    def __init__(self, wb):
        self.wb = wb
        self.first_list = LocationsList(wb)
        self.second_list = LocationsDict(wb)
        self.normalized_list = NormalizedLocationsList()

    # PUBLIC
    def normalize_names(self):
        """Looks in second_list for each location in first_list, when there is
        a match, adds to normalized list."""

        # clear normalized list
        self.normalized_list = NormalizedLocationsList()

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
                print "no match,", location

            i += 1

    def count_first_list(self):
        return self.first_list.count()

    def count_second_list(self):
        return self.second_list.count()

    def save(self, output_file=None):
        """Add normalized list to workbook and save it."""

        # take default values
        output_file = output_file or self.OUTPUT_FILE

        # create new sheet
        ws = self.wb.create_sheet(-1, "tbl_normal")

        # write each record in excel sheet
        for record in self.normalized_list:
            write_ws(ws, record)

        # save excel with normalized list of locations
        self.wb.save(output_file)


# DATA
INPUT_FILE = "location_names.xlsx"
OUTPUT_FILE = "location_names_normalized.xlsx"


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
