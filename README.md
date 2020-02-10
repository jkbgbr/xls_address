# xls_address
A module to make accessing data excel files in xlrd easier: addressing just like in excel

## Getting Started

I never found a simple way to retreive data by excel-type cell addressing when using [xlrd](https://pypi.org/project/xlrd/), whick can be a pain if the data is sparse. 

So I created mine.

This allows to define unique cells or ranges using Excel's notation when reading data using xlrd, which is - IMHO - easier to read and debug.
The retreived data may either be xlrd's native Cell object or its value, directly. 

### Prerequisites

To use it you'll obviously need xlrd, for testing [xlwt](https://pypi.org/project/xlwt/) is also required; both are included in the requirements.txt file. 

### Installing

Currently no PyPI package, so just download the files in a separate directory and install them via

```
pip install -r requirements.txt <path>
```

with <path> being the location of the downloaded files. 

### Usage

The function `get_values()` is to be used to retreive values from a cell or a range of a sheet previously opened.

```
>>> import xlrd
>>> wb = xlrd.open_workbook('test.xls')  # opens the file at the path
>>> sheet = wb.sheet_by_index(0)  # takes the first sheet

>>> print(get_value(sheet, addr='C4:D5', value_only=True))  # reads the values from the range.
>>> [[1.0, 2.0], [3.0, 4.0]]
>>> print(get_value(sheet, addr='C4', value_only=True))  # single cell
>>> 1.0
>>> print(get_value(sheet, addr='C4', value_only=False))  # single cell, returned is an xlrd.Cell
>>> number:1.0
```

for `value_only=False` the return value is the Cell. 

The value of addr may be either a single cell or a range as in the examples.
If the address or range is invalid, a `ValueError` is raised.
Results from ranges yield a nested lists with the inner lists being the rows.
Requesting data from outside the range where there are some returns empty lists the size of the range.

No upper bound checking is performed, but in general the sanity of the input address or range is checked. See the tests for more info. 

## Running the tests

Tests are included in `test_xlddss.py`. The tests require xlwt as a test file is created during testing.
You can specify the path to this file at the top of the test file. 

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
