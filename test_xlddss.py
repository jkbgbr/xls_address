import unittest
import xlddss
import xlrd


class TestSingle(unittest.TestCase):

    def test_letters(self):
        """Letter range"""
        self.assertEqual(xlddss.LETTERS[0], 'A')
        self.assertEqual(xlddss.LETTERS[-1], 'Z')

    def test_address_matching(self):
        """tests the regexp"""
        # these will fail
        self.assertRaises(ValueError, xlddss._parse_address, 'a1')
        self.assertRaises(ValueError, xlddss._parse_address, 'a 1')
        self.assertRaises(ValueError, xlddss._parse_address, 'a:1')
        self.assertRaises(ValueError, xlddss._parse_address, '1')
        self.assertRaises(ValueError, xlddss._parse_address, 'A')
        self.assertRaises(ValueError, xlddss._parse_address, 'AAAA')
        self.assertRaises(ValueError, xlddss._parse_address, 'AAA')
        self.assertRaises(ValueError, xlddss._parse_address, 'A12345678')
        self.assertRaises(ValueError, xlddss._parse_address, 'A1B2')
        self.assertRaises(ValueError, xlddss._parse_address, 'Ä0')

        # these will return the address
        self.assertEqual(xlddss._parse_address('$A$1'), 'A1')
        self.assertEqual(xlddss._parse_address('A1'), 'A1')
        self.assertEqual(xlddss._parse_address('XXX1234567'), 'XXX1234567')
        self.assertEqual(xlddss._parse_address('A1'), 'A1')
        self.assertEqual(xlddss._parse_address('AAA1'), 'AAA1')
        self.assertEqual(xlddss._parse_address('AAA1'), 'AAA1')

        # these are OK fro parsing but not later
        self.assertEqual(xlddss._parse_address('A0'), 'A0')

    def test_cell_address(self):
        """the method returning the address"""
        self.assertEqual(xlddss._cell_address('A1'), (0, 0))
        self.assertEqual(xlddss._cell_address('C1'), (0, 2))
        self.assertEqual(xlddss._cell_address('C5'), (4, 2))
        self.assertEqual(xlddss._cell_address('XFD1'), (0, 16383))
        self.assertEqual(xlddss._cell_address('XFD1048576'), (1048575, 16383))

    def test_range(self):
        """Returning a range"""
        self.assertTupleEqual(xlddss._cell_range('A1:B2'), ((0, 0), (1, 1)))
        # special case: not a range really
        self.assertTupleEqual(xlddss._cell_range('A1'), (0, 0))
        self.assertTupleEqual(xlddss._cell_range('B3'), (2, 1))

        # these are invalid, will fail
        self.assertRaises(ValueError, xlddss._cell_range, '')
        self.assertRaises(ValueError, xlddss._cell_range, 'A1:B2:C3')


class TestReadingFromFile(unittest.TestCase):

    def setUp(self) -> None:
        wb = xlrd.open_workbook('sample.xlsx')
        self.sheet = wb.sheet_by_index(0)

    def test_reading_given_column(self):
        """Reading from a given given column"""
        for i in range(self.sheet.nrows):
            _addr = xlddss._cell_address('A{}'.format(i+1))
            try:
                self.assertEqual(self.sheet.cell_value(*_addr), i)
            except IndexError:
                print('no data in {}'.format(_addr))

    def test_reading_given_row(self):
        """Reading from a given given row"""
        for i in xlddss.LETTERS:
            _addr = xlddss._cell_address('{}1'.format(i))
            try:
                self.assertEqual(self.sheet.cell_value(*_addr), xlddss.LETTERS.index(i))
            except IndexError:
                print('no data in {}'.format(_addr))

    def test_reading_range_singlecell(self):
        """Using the method meant for range reading to read a single cell"""
        _addr = xlddss._cell_address('C4')
        self.assertEqual(self.sheet.cell_value(*_addr), 1)

        self.assertEqual(xlddss.get_value(self.sheet, addr='C4', value_only=True), 1)
        self.assertEqual(xlddss.get_value(self.sheet, addr='F7', value_only=True), 5)
        self.assertEqual(xlddss.get_value(self.sheet, addr='I7', value_only=True), 'text')
        self.assertEqual(xlddss.get_value(self.sheet, addr='J7', value_only=True), 'data')

        self.assertEqual(xlddss.get_value(self.sheet, addr='C4:D5', value_only=True), [[1, 2], [3, 4]])
        self.assertEqual(xlddss.get_value(self.sheet, addr='F7:F11', value_only=True), [[5], [5], [5], [5], [5]])

        self.assertRaises(ValueError, xlddss.get_value, *(self.sheet, 'D5:C4', True))


if __name__ == '__main__':
    unittest.main()
