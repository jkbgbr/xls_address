"""
Makes working with excel files a bit easier by allowing to read from an xls file using excel's
row-column designation
"""

import re


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
        addr = _cell_range(addr)
    except ValueError:
        raise

    try:
        addr = addr[0][0], addr[0][1], addr[1][0], addr[1][1]
        if value_only:
            _ret = []
            for r in _get_cell_range(sheet, *addr):
                _ret.append([x.value for x in r])
            return _ret
        else:
            return _get_cell_range(sheet, *addr)

    except TypeError:
        if value_only:
            return sheet.cell_value(*addr)
        else:
            return sheet.cell(*addr)


if __name__ == '__main__':
    # print(_cell_address('$B$3'))
    # print(_cell_address('AD3'))
    print(_cell_address('XFD1'))
    print(_column_adress('XFD'))
    print(_row_adress('1'))


    # print(LETTERS.index('X'))
