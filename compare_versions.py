# -*- coding: utf-8 -*-
import xlrd
from xlwt import *
from datetime import datetime
import os


class ExcelFile(object):
    def __init__(self, workbook):
        self.workbook = xlrd.open_workbook(workbook)
        self.worksheet = self.workbook.sheet_by_index(0)


def return_list(worksheet):
    """Returns all rows from a worksheet as a nested list"""
    all_rows = []
    curr_row = -1
    num_rows = worksheet.nrows - 1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        # handling dates
        values = [xlrd.xldate_as_tuple(c.value, 0)
                  if c.ctype == 3 else c.value for c in row]
        # converting xldate tupels to actual dates in text format
        values = [datetime.strftime(datetime(*element), '%Y-%m-%d')
                  if type(element) == tuple else element for element in values]
        all_rows.append(values)
    return all_rows


def return_headers(elements):
    """Returns first list from :elements - a nested list"""
    return elements[0]


def encode_as_utf(elements):
    """Encoding each string element of the list as utf-8"""
    new_list = []
    for row in elements:
        new_list.append([element.encode('utf-8', 'replace')
                         if isinstance(element, str)
                         else element for element in row])
    return new_list


def join_primary_key_elements(elements):
    """Returns elements as string (joined by _)

    Param:
    :elements - list"""
    return '_'.join([str(element) if isinstance(element, int)
                     or isinstance(element, float)
                     else element for element in elements])


def convert_none_to_string(elements):
    """If any element of :elements is of NoneType
    it is replaced by the string 'Empty'

    Param:
    :elements - list"""
    elements_no_nulls = ['empty' if element is None else element
                         for element in elements]
    return elements_no_nulls


def get_id_column_indexes(input_dataset, columns):
    """Returns a list containing positions of :columns
    in the first element of :input_dataset.

    Params:
    :input_dataset - nested list, e.g. output of return_list
    :columns - list

    Example:
    input_dataset = [['a','b','c','d'], ['x1', 'x2', 'x3']]
    columns = ['b', 'd']
    output = [1, 3]
    """
    first_row = input_dataset[0]
    column_indexes = [first_row.index(column) for column in columns]
    return column_indexes


def get_row_primary_key_elements(row, indexes):
    """
    Returns elements which form a unique identifier of a row/

    Params:
    :row - list
    : indexes - list

    Example:
    row = ['col1', 'col2', 'col3', 'col4']
    indexes = [0, 2]
    output = ['col1', 'col3']
    """
    return [element for n, element in enumerate(row) if n in indexes]


def detect_changes(input_dataset_1, input_dataset_2):
    """Returns a list of columns not matching.

    Params:
    :input_dataset_1 - nested list, previous version of the document
    :input_dataset_2 - nested list, current version of the document

    Example:
    input_dataset_1 = [['element1', 'element2', 'primary_key1'],
                       ['element1', 'element2', 'primary_key2']]
    input_dataset_2 = [['element1', 'x', 'primary_key1'],
                       ['element1', 'element2', 'primary_key2']]
    output = [[1], []]
    """
    outer_list = []
    for row2 in input_dataset_2:
        inner_list = []
        relevant_row_from_dataset_1 = find_relevant_row(row2, input_dataset_1)
        for position, element in enumerate(row2):
            if element != relevant_row_from_dataset_1[position]:
                inner_list.append(position)
        outer_list.append(inner_list)
    return outer_list


def find_relevant_row(row, input_dataset):
    """Returns a record from :input_list which have the same
    unique identifier as :row. It is assumed, that the last
    element in both :input_dataset_1 and :input_dataset_2
    contains unique identifier.

    Params:
    :row - list
    :input_dataset - nested list

    Example:
    row = ['a', 'b', 'id1']
    input_dataset = [['a', 'b', 'id0'], ['a', 'a', 'id1']]
    output = ['a', 'a', 'id1']
    """
    try:
        relevant_row = [element for element in input_dataset
                        if element[-1] == row[-1]][0]
        return relevant_row
    except:
        print "Relevant row not found: %s" % row[-1]


def get_current_date(date_format="%Y%m%d_%H%M"):
    """Returns current date as a string"""
    now = datetime.now()
    today = now.strftime(date_format)
    return today


def write_to_file(previous_version, current_version, removed,
                  added, compared, columns_not_matching):
    style_cell_change = easyxf('pattern: pattern solid, fore_color red; font: color white;')

    # creating new object
    book = Workbook()
    # adding sheets
    previous_worksheet = book.add_sheet('Previous version')
    current_worksheet = book.add_sheet('Current version')
    removed_worksheet = book.add_sheet('Removed')
    added_worksheet = book.add_sheet('Added')
    differences_worksheet = book.add_sheet('Differences')
    # populating first four sheets
    data_per_sheet = dict(zip([previous_worksheet, current_worksheet,
                               removed_worksheet, added_worksheet],
                              [previous_version, current_version,
                               removed, added]))
    for key, value in data_per_sheet.iteritems():
        for row_index, row in enumerate(value):
            for cell_index, cell in enumerate(row):
                key.write(row_index, cell_index, cell)
    # populating the last sheet (differences)
    for row_index, row in enumerate(compared):
        wrong_columns = columns_not_matching[row_index]
        cells_not_matching = ((element[0], element[1]) for element in enumerate(row) if element[0] in wrong_columns)
        cells_matching = ((element[0], element[1]) for element in enumerate(row) if element[0] not in wrong_columns)
        for cell in cells_not_matching:
            differences_worksheet.write(row_index, cell[0],
                                        cell[1], style_cell_change)
        for cell in cells_matching:
            differences_worksheet.write(row_index, cell[0], cell[1])
    # save book object as Excel file
    now = get_current_date()
    file_name = '_'.join(['compared', now, '.xls'])
    work_dir = os.path.dirname(os.path.realpath(__file__))
    book.save(file_name)
    print "Saved output in %s" % os.path.join(work_dir, file_name)
    return os.path.join(work_dir, file_name)


def compare_excel_files(previous_file, current_file, id_columns):
    """
    Performs a comparison of two Excel files and saves output in
    the directory of the script.

    Params:
    :previous_file: string, file path
    :current_file: string, file path
    :param id_columns: list, column names

    Example:
    compare_excel_files(r'test_file_1.xlsx', r'test_file_2.xlsx',['col1', 'col2'])
    """
    # Excel files to Workbook objects
    previous_workbook = ExcelFile(previous_file)
    current_workbook = ExcelFile(current_file)
    # Workbook objects to nested lists
    previous_dataset_decoded = return_list(previous_workbook.worksheet)
    current_dataset_decoded = return_list(current_workbook.worksheet)
    previous_dataset = encode_as_utf(previous_dataset_decoded)
    current_dataset = encode_as_utf(current_dataset_decoded)
    # add primary keys, i.e. concatenated combination of multiple columns,
    # as last elements of each record
    id_column_indexes = get_id_column_indexes(previous_dataset, id_columns)
    for row in previous_dataset:
        row.append(join_primary_key_elements(get_row_primary_key_elements(row, id_column_indexes)))
    for row in current_dataset:
        row.append(join_primary_key_elements(get_row_primary_key_elements(row, id_column_indexes)))
    # getting unique identifiers of both datasets
    previous_dataset_ids = set([row[-1] for row in previous_dataset])
    current_dataset_ids = set([row[-1] for row in current_dataset])
    # rows removed in current_file
    removed = [row for row in previous_dataset if row[-1]
               not in current_dataset_ids]
    # rows added in current_file
    added = [row for row in current_dataset if row[-1]
             not in previous_dataset_ids]
    # rows existing in both datasets
    rows_to_analyze = [row for row in current_dataset
                       if row[-1] in previous_dataset_ids]
    # compare
    changes = detect_changes(previous_dataset, rows_to_analyze)
    # save
    output = write_to_file(previous_dataset, current_dataset,
                           removed, added, rows_to_analyze, changes)
    print "Output: %s" %output
    return
