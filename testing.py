import unittest
import compare_versions as cv


class testing_cv(unittest.TestCase):
    def test_join_primary_key_elements_pass(self):
        pk = ['test', 'primary', 'key']
        concat_pk = cv.join_primary_key_elements(pk)
        self.assertTrue(concat_pk == 'test_primary_key')

    def test_convert_none_to_string_pass(self):
        test_input = [None, 1, 'a', True]
        test_output = cv.convert_none_to_string(test_input)
        self.assertTrue(test_output == ['empty', 1, 'a', True])


    def test_get_id_column_indexes_pass(self):
        test_input_list = [['Row 1: Element 1', 'Row 1: Element 2', 'Row1: Element 3'],
                           ['Row 2', 'Row 2', 'Row2'],
                           ['Row 3', 'Row 3', 'Row3']]
        test_columns = ['Row 1: Element 1', 'Row 1: Element 2']
        test_output = cv.get_id_column_indexes(test_input_list, test_columns)
        self.assertTrue(test_output == [0, 1])

    def test_detect_changes_pass(self):
        test_input_dataset1 = [[1, 2, 'one'],
                               [2, 3, 'two']]
        test_input_dataset2 = [[1, 1, 'one'],
                               [2, 3, 'two']]
        test_output = cv.detect_changes(test_input_dataset1, test_input_dataset2)
        self.assertTrue(test_output == [[1], []])

    def test_find_relevant_row_pass(self):
        test_row = ['a', 'a', 'id1']
        test_input_dataset = [['b', 'b', 'id1'], ['a', 'a', 'id2']]
        test_output = cv.find_relevant_row(test_row, test_input_dataset)
        self.assertTrue(test_output == ['b', 'b', 'id1'])

    def test_return_headers_pass(self):
        wb = cv.ExcelFile(r'test_file_1.xlsx')
        wb_list = cv.return_list(wb.worksheet)
        self.assertTrue(cv.return_headers(wb_list) == [u'col1', u'col2', u'col3', u'col4'])

    def test_get_row_primary_key_elements_pass(self):
        test_row = ["element1", "element2", "element3"]
        self.assertTrue(cv.get_row_primary_key_elements(test_row, [0, 2]) == [u'element1', u'element3'])


if __name__ == '__cv__':
    unittest.cv()









