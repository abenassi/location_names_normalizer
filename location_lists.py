#!C:\Python27
# -*- coding: utf-8 -*-
from fuzzywuzzy import process
from utils import normalize_name


# PRIVATE CLASSES
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

    def count(self):
        pass

    def _count_fields(self, ws):
        """Count number of fields in worksheet"""

        field_number = 1
        while ws.cell(row=0, column=field_number).value:
            field_number += 1

        return field_number


# PUBLIC CLASSES
class NormalizedLocationsList(list, BaseLocations):

    def add(self, location, location_matched):
        """Adds normalized location based on location_matched way to write it,
        with its correspondence ids from both."""

        # start with the two matched ids
        normalized_location = [location.id, location_matched.id]

        # extend with location and location_matched
        normalized_location.extend(location)
        normalized_location.extend(location_matched)

        # extend with location matched normalized
        location_matched_norm = [normalize_name(i) for i in location_matched]
        normalized_location.extend(location_matched_norm)

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

    THRESHOLD_RATIO = 75

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
        value = normalize_name(location[key_index])

        # extract matched value from
        value_matched = process.extractOne(value, dictionary.keys())

        if value_matched and value_matched[1] > self.THRESHOLD_RATIO:
            key = value_matched[0]

            # if there are more values to evaluate, call recursively
            if len(location) > key_index + 1:
                print value_matched[1],
                return self.find(location, dictionary[key], key_index + 1)

            else:
                print value_matched[1],
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
        key_value = normalize_name(location[key_index])

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
