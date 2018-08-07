from pathlib import Path

import pyexcel_openpyxlx

if __name__ == '__main__':
    print(pyexcel_openpyxlx.get_styles(str(Path('../tests/input').joinpath('user.xlsx')),
                                       skip_cell_formatting=True))
