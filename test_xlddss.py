import unittest
import xlddss


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
        self.assertEqual(xlddss._cell_address('C1'), (2, 0))
        self.assertEqual(xlddss._cell_address('C5'), (2, 4))
        self.assertEqual(xlddss._cell_address('XFD1'), (16383, 0))
        self.assertEqual(xlddss._cell_address('XFD1048576'), (16383, 1048575))

    def test_range(self):
        """Returning a range"""
        self.assertTupleEqual(xlddss._cell_range('A1:B2'), ((0, 0), (1, 1)))
        # special case: not a range really
        self.assertTupleEqual(xlddss._cell_range('A1'), (0, 0))
        self.assertTupleEqual(xlddss._cell_range('B3'), (1, 2))

        # these are invalid, will fail
        self.assertRaises(ValueError, xlddss._cell_range, '')
        self.assertRaises(ValueError, xlddss._cell_range, 'A1:B2:C3')


if __name__ == '__main__':
    unittest.main()
