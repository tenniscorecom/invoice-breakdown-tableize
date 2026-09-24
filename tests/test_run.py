"""run.py の業務フロー全体テスト。

複数シート（正常・失敗）を持つ入力 Excel を ``run()`` に通し、
正常シートにはテーブルが登録され、失敗シートはスキップされて処理が止まらないことを確認する。
"""

from pathlib import Path

from openpyxl import load_workbook


def _write_config(tmp_path: Path, *, header_row: str | None = "1") -> Path:
    """テスト用の config.ini を ``tmp_path`` 配下に書き出す。"""
    input_file = tmp_path / "input.xlsx"
    output_folder = tmp_path / "output"
    lines = [
        "[FILES]",
        f"INPUT_FILE = {input_file}",
        f"OUTPUT_FOLDER = {output_folder}",
        "",
        "[EXCEL]",
        "OUTPUT_PREFIX = テーブル化_",
    ]
    if header_row is not None:
        lines.append(f"HEADER_ROW = {header_row}")
    config_path = tmp_path / "config.ini"
    config_path.write_text("\n".join(lines), encoding="utf-8")
    return config_path


def test_run_converts_normal_sheets_and_skips_failing_sheet(
    tmp_path: Path, make_input_book, use_config
) -> None:
    # 1 枚目: 正常な表。2 枚目: 見出し行に空セルがあり ``ExcelHeaderError``（空セル） で失敗する。
    # 3 枚目: 見出し行に重複があり ``ExcelHeaderError``（重複） で失敗する。
    input_path = make_input_book(
        tmp_path / "input.xlsx",
        {
            "正常A": [("1", "りんご", "100"), ("2", "みかん", "80")],
            "見出し空": [("1", "りんご", "100")],
            "見出し重複": [("1", "りんご", "100")],
        },
    )
    # 2 枚目・3 枚目シートの見出しを書き換えて失敗条件を満たす。
    workbook = load_workbook(input_path)
    workbook["見出し空"]["B1"] = None  # type: ignore[index]
    workbook["見出し重複"]["B1"] = "col_a"  # type: ignore[index]
    workbook.save(input_path)
    workbook.close()

    use_config(_write_config(tmp_path))

    from src.run import run

    output_path = run()

    # 元ファイルは変更されない。
    assert (tmp_path / "input.xlsx").exists()
    # 出力ファイル名は ``OUTPUT_PREFIX + 元ファイル名``。
    assert output_path == tmp_path / "output" / "テーブル化_input.xlsx"
    assert output_path.exists()

    result = load_workbook(output_path)
    try:
        # 正常シートにはテーブルが登録されている（シート名がそのまま使える形なので接頭辞なし）。
        normal_sheet = result["正常A"]
        assert "正常A" in normal_sheet.tables  # type: ignore[operator]
        # 失敗シートにはテーブルが登録されていない。
        assert "見出し空" not in result["見出し空"].tables  # type: ignore[operator]
        assert "見出し重複" not in result["見出し重複"].tables  # type: ignore[operator]
    finally:
        result.close()


def test_run_uses_header_row_from_config(tmp_path: Path, make_input_book, use_config) -> None:
    """``[EXCEL] HEADER_ROW = 2`` のとき 2 行目を見出しとして登録される。"""
    input_path = make_input_book(
        tmp_path / "input.xlsx",
        {
            "Sheet1": [("date", "item", "amount"), ("1", "りんご", "100")],
        },
    )
    # 1 行目はタイトル列（空ではないが使わない）、2 行目を見出しとして扱う。
    workbook = load_workbook(input_path)
    workbook["Sheet1"]["A1"] = "タイトル行"  # type: ignore[index]
    workbook.save(input_path)
    workbook.close()

    use_config(_write_config(tmp_path, header_row="2"))

    from src.run import run

    output_path = run()

    result = load_workbook(output_path)
    try:
        sheet = result["Sheet1"]
        table_names = list(sheet.tables)  # type: ignore[attr-defined]
        assert len(table_names) == 1
        # テーブル範囲が ``A2:C3`` 始まり（タイトル行を含まない）ことを確認。
        table = sheet.tables[table_names[0]]  # type: ignore[index]
        assert table.ref.split(":")[0].startswith("A2")  # type: ignore[attr-defined]
    finally:
        result.close()


def test_run_treats_missing_header_row_as_auto(tmp_path: Path, make_input_book, use_config) -> None:
    """``HEADER_ROW`` 未設定（キーなし）のときは A2 ルール自動推定で止まらずに処理できる。"""
    make_input_book(
        tmp_path / "input.xlsx",
        {
            "Sheet1": [("1", "りんご", "100")],
        },
    )
    use_config(_write_config(tmp_path, header_row=None))

    from src.run import run

    # A2 ルールでは 1 行目が見出し（結合なし）として扱われるので、止まらずにテーブル化される。
    output_path = run()
    assert output_path.exists()

    result = load_workbook(output_path)
    try:
        assert len(list(result["Sheet1"].tables)) == 1  # type: ignore[attr-defined]
    finally:
        result.close()


def test_sanitize_table_name_adds_prefix_for_digit_leader() -> None:
    """先頭が数字のシート名には ``T_`` を付ける（セル参照と紛らわしい）。"""
    from src.run import _sanitize_table_name

    assert _sanitize_table_name("2024年1月").startswith("T_")
    # 全角数字も「先頭が数字」相当として接頭辞を付ける。
    assert _sanitize_table_name("１月").startswith("T_")


def test_sanitize_table_name_replaces_forbidden_characters() -> None:
    """空白・禁則文字は ``_`` に置換される。"""
    from src.run import _sanitize_table_name

    sanitized = _sanitize_table_name("sheet 1/2")
    assert " " not in sanitized and "/" not in sanitized
    # 全体が ``_`` 以外で構成されていること（セル参照パターンと判定されない形）。
    assert "_" in sanitized


def test_sanitize_table_name_handles_cell_reference_like_sheet() -> None:
    """シート名が ``A1`` / ``R1C1`` のようなセル参照パターンなら接頭辞を付ける。"""
    from src.run import _sanitize_table_name

    assert _sanitize_table_name("A1").startswith("T_")
    assert _sanitize_table_name("R1C1").startswith("T_")
