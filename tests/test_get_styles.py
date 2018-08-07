import pytest
from pathlib import Path

import pyexcel_openpyxlx


@pytest.mark.parametrize('in_file', [
    'user.xlsx'
])
def test_get_styles(in_file, request):
    pyexcel_openpyxlx.get_styles(str(Path('tests/input').joinpath(in_file)),
                                 str(Path('tests/output').joinpath(request.node.name).with_suffix('.yaml')),
                                 skip_cell_formatting=True)
