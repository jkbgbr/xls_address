"""
Makes working with excel files a bit easier by allowing to read from an xls file using excel's
row-column designation
"""

import re
import xlwt


LETTERS = tuple([chr(x) for x in range(65, 65+26)])  # A to Z


def _column_adress(addr='A'):
    """returns the column number for a column adress"""
    return _cell_address(''.join([addr, '1']))[0]


def _row_adress(addr='1'):
    """returns the rown number for a column adress"""
    return _cell_address(''.join(['A', addr]))[1]


def _parse_address(addr='XFD1048576') -> str:
    """
    Parses the address given using a regexp and returns a cleaned string that can be split
    Accepts any valid excel cell address definition, incl. absoulte adresses with $.
    :param addr:
    :return:
    """
    patt = re.match(r'^(\$?[A-Z]{1,3}\$?\d{1,7})$', addr)
    if patt is None:
        raise ValueError('Could not parse the address {}'.format(addr))
    else:
        return patt.group(0).replace('$', '')


def _cell_address(addr='XFD1048576', rev=False):
    """parses the input address and returns the column, row tuple corresponding"""

    # check the address. Expected is something between 'A1' and 'XFD1048576'
    try:
        _ret = _parse_address(addr)
    except ValueError:
        raise

    _row, _col = None, None

    addr = _ret
    _letters = ''.join([x for x in addr if x in LETTERS])
    _numbers = ''.join([x for x in addr if x not in _letters])

    # getting the row number
    # try:
    _row = int(_numbers) - 1
    if _row < 0:
        raise ValueError('Incorrect row position in the address: {}!'.format(_numbers))

    # getting the column. len(LETTERS)-base arithmetic
    _col = 0
    for col in range(len(_letters), 0, -1):
        he = len(_letters) - col  # position, 1, 2 ...
        val = _letters[col - 1]  # value at position
        _col += (LETTERS.index(val) + 1) * (len(LETTERS) ** he)
    _col -= 1

    if rev:
        _col, _row = _row, _col

    return _row, _col


def _cell_range(rnge='A1:XFD1048576') -> tuple:
    """Returns the addresses from range"""

    # splitting
    rnge = rnge.split(':')
    # the split results a single value - fall back to cell
    if len(rnge) == 1:
        return _cell_address(addr=rnge[0])

    elif len(rnge) > 2 or len(rnge) <= 0:
        raise ValueError('The provided range "{}" is not correct'.format(rnge))

    else:  # len(rnge) == 2
        return _cell_address(rnge[0]), _cell_address(rnge[1])


def _get_cell_range(sheet, start_row, start_col, end_row, end_col):
    """Returns the values from a range
    https://stackoverflow.com/a/33938163
    """
    return [sheet.row_slice(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]


def get_value(sheet, addr, value_only=False):
    """Use this to retreive stuff"""
    # this does not work at all
    try:
        _addr = _cell_range(addr)
    except ValueError:
        raise

    try:
        _addr = _addr[0][0], _addr[0][1], _addr[1][0], _addr[1][1]

        # checking a range definition validity: start_ is smaller of same as end_
        if not (_addr[0] <= _addr[2]) or not(_addr[1] <= _addr[3]):
            raise ValueError('The range definition {} is not correct'.format(addr))

        if value_only:
            _ret = []
            for r in _get_cell_range(sheet, *_addr):
                _ret.append([x.value for x in r])
            return _ret
        else:
            return _get_cell_range(sheet, *_addr)

    except TypeError:
        if value_only:
            return sheet.cell_value(*_addr)
        else:
            return sheet.cell(*_addr)


def write_samplefile():
    book = xlwt.Workbook()
    sheet = book.add_sheet("Table1")

    for rownum in range(41):
        row = sheet.row(rownum)
        row.write(0, rownum)

    for colnum in range(1, 26):
        row = sheet.row(0)
        # col = sheet.col(colnum)
        row.write(colnum, colnum)

    # C4:D5
    sheet.row(3).write(2, 1)
    sheet.row(3).write(3, 2)
    sheet.row(4).write(2, 3)
    sheet.row(4).write(3, 4)

    # G4:J4
    for i in range(6, 10):
        sheet.row(3).write(i, 1)

    # F7:F11
    for i in range(6, 11):
        sheet.row(i).write(5, 5)

    # I7
    sheet.row(6).write(8, 'text')
    # J7
    sheet.row(6).write(9, 'data')

    book.save('test.xls')


if __name__ == '__main__':
    import xlrd
    import os
    write_samplefile()
    wb = xlrd.open_workbook('test.xls')
    sheet = wb.sheet_by_index(0)
    print(get_value(sheet, addr='C4:D5', value_only=True))
    os.remove('test.xls')