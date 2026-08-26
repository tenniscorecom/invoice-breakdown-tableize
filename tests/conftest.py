"""tests/conftest.py — テスト共通の前準備

複数シート（正常・失敗）を持つ ``tmp_path`` 入力 Excel を作る fixture を置く。
"""

from collections.abc import Iterable
from pathlib import Path

import pytest
from comken import Config
from openpyxl import Workbook


@pytest.fixture
def make_input_book():
    """``{sheet_name: rows_of_lists_or_empty}`` の dict から複数シートの Excel を作る。

    各シートの 1 行目を見出しとして扱う。``rows`` が空のときは見出しのみのシートを作る。
    """

    def _make(path: Path, sheets: dict[str, Iterable[Iterable[object]]]) -> Path:
        workbook = Workbook()
        # デフォルトで存在する Sheet を 1 番目の名前に rename してから残りを追加。
        first_name = next(iter(sheets))
        default_sheet = workbook.active
        assert default_sheet is not None
        default_sheet.title = first_name
        first_rows = list(sheets[first_name])
        for row in [("col_a", "col_b", "col_c"), *first_rows]:
            default_sheet.append(list(row))
        for name, rows in list(sheets.items())[1:]:
            worksheet = workbook.create_sheet(name)
            for row in [("col_a", "col_b", "col_c"), *list(rows)]:
                worksheet.append(list(row))
        workbook.save(path)
        workbook.close()
        return path

    return _make


@pytest.fixture
def use_config(monkeypatch):
    """``src.run.config`` を指定パスの ``Config`` で差し替える。"""

    def _use(path: Path) -> Config:
        config = Config(path)
        monkeypatch.setattr("src.run.config", config)
        return config

    return _use
