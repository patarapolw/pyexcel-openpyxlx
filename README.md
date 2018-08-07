# pyexcel-openpyxlx

[![Build Status](https://travis-ci.org/patarapolw/pyexcel-openpyxlx.svg?branch=master)](https://travis-ci.org/patarapolw/pyexcel-openpyxlx)
[![PyPI version shields.io](https://img.shields.io/pypi/v/pyexcel_openpyxlx.svg)](https://pypi.python.org/pypi/pyexcel_openpyxlx/)
[![PyPI license](https://img.shields.io/pypi/l/pyexcel_openpyxlx.svg)](https://pypi.python.org/pypi/pyexcel_openpyxlx/)
[![PyPI pyversions](https://img.shields.io/pypi/pyversions/pyexcel_openpyxlx.svg)](https://pypi.python.org/pypi/pyexcel_openpyxlx/)

Export styles from Excel using [OpenPyXL](http://openpyxl.readthedocs.io/en/stable/) (for [XlsxWriter](https://xlsxwriter.readthedocs.io) -- [pyexcel-xlsxwx](https://github.com/patarapolw/pyexcel-xlsxwx)).

## Installation

```commandline
$ pip install pyexcel-openpyxlx
```

## Usage

```python
>>> import pyexcel_openpyxlx
>>> pyexcel_openpyxlx.get_styles('example.xlsx')
{'worksheet': {'hanzi': {'freeze_panes': 'A2', 'column_width': [7.7109375, 28.7109375, 30.7109375, 8.7109375, 10.7109375, 16.7109375, 7.7109375, 6.7109375, 6.7109375, 30.7109375]}, 'sentences': {'freeze_panes': 'A2', 'column_width': [30.7109375, 30.7109375, 30.7109375, 30.7109375, 6.7109375, 6.7109375, 30.7109375]}, 'vocab': {'freeze_panes': 'A2', 'column_width': [7.7109375, 24.7109375, 24.7109375, 30.7109375, 30.7109375, 30.7109375, 7.7109375, 6.7109375, 6.7109375, 8.7109375, 30.7109375]}}}
```
Exporting to YAML is also possible by specifying `out_file`.

```python
>>> pyexcel_openpyxlx.get_styles('example.xlsx', out_file='example.yaml')
```

```yaml
worksheet:
  hanzi:
    column_width: [7.7109375, 28.7109375, 30.7109375, 8.7109375, 10.7109375, 16.7109375,
      7.7109375, 6.7109375, 6.7109375, 30.7109375]
    freeze_panes: A2
  sentences:
    column_width: [30.7109375, 30.7109375, 30.7109375, 30.7109375, 6.7109375, 6.7109375,
      30.7109375]
    freeze_panes: A2
  vocab:
    column_width: [7.7109375, 24.7109375, 24.7109375, 30.7109375, 30.7109375, 30.7109375,
      7.7109375, 6.7109375, 6.7109375, 8.7109375, 30.7109375]
    freeze_panes: A2
```

## Related projects

- [pyexcel-xlsxwx](https://github.com/patarapolw/pyexcel-xlsxwx) -- Save pyexcel data with XlsxWriter, while retaining good formatting.
