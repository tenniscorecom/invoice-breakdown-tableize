"""src/run.py — 処理の本体

main.py から呼ばれる。ここに「実際にやりたいこと」を書く。
外注先から届く請求書内訳 Excel の各シート（テーブル化されていないただの表）を、
comken の ``Excel.convert_range_to_table`` で Excel テーブル（ListObject）に変換する。
"""

import logging
import re
import shutil
from pathlib import Path

from comken import config
from comken.exceptions import (
    ComkenError,
    ConfigKeyNotFoundError,
)
from comken.toolbox.excel import Excel

logger = logging.getLogger(__name__)

# comken の ``Sheet._validate_table_name`` と同じ制約の抜粋。
# ここを通る名前を組み立てれば convert_range_to_table の ``InvalidTableNameError`` を避けられる。
# _CELL_REFERENCE_PATTERN は ``A1`` / ``R1C1`` などセル参照と紛らわしい形を弾くもの
# （数字始まりは別の条件で弾くが、ここでは一応まとめて接頭辞付けの対象にする）。
_TABLE_NAME_FORBIDDEN_CHARS = re.compile(r"[\s\[\]/\\:*?\"<>|'`#%@$&+={}~]")
_TABLE_NAME_CELL_REF_PATTERN = re.compile(r"^[A-Z]+[0-9]+$|^R[0-9]+C[0-9]+$", re.IGNORECASE)
_TABLE_NAME_DIGIT_PREFIX = "T_"


def run() -> Path:
    """処理の入口。main.py から呼ばれる。

    戻り値は出力ファイルのパス。1 シートの変換失敗は ``ComkenError`` として扱い、
    そのシートだけをスキップして次のシートへ進む。それ以外の例外はそのまま送出する。
    """
    input_file = Path(config.FILES.INPUT_FILE)
    output_folder = Path(config.FILES.OUTPUT_FOLDER)
    output_prefix = str(config.EXCEL.OUTPUT_PREFIX)
    header_row = _resolve_header_row(config.EXCEL)

    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / f"{output_prefix}{input_file.name}"
    # 元ファイルへは書き込まない。作業用にコピーしてから openpyxl で開く。
    shutil.copy2(input_file, output_path)

    success_count = 0
    failure_count = 0
    skipped_sheets: list[str] = []

    with Excel(output_path, engine="openpyxl") as excel:
        for sheet_name in excel.list_sheets():
            try:
                sheet = excel.sheet(sheet_name)
                start, end = sheet.get_used_range()
                table_name = _sanitize_table_name(sheet_name)
                excel.convert_range_to_table(
                    sheet_name,
                    range=f"{start}:{end}",
                    table_name=table_name,
                    header_row=header_row,
                )
                success_count += 1
                logger.info(
                    "シート「%s」をテーブル化しました（テーブル名: %s）",
                    sheet_name,
                    table_name,
                )
            except ComkenError as error:
                # 想定内のエラー（空見出し・重複見出し・範囲不正など）は
                # 1 シートだけスキップして次へ進む。
                failure_count += 1
                skipped_sheets.append(sheet_name)
                logger.warning(
                    "シート「%s」をスキップしました: %s",
                    sheet_name,
                    error,
                )
        # with を抜けるときに保存されるが、失敗時の例外パスでは保存されないよう明示しておく。
        excel.save()

    logger.info(
        "全シート処理完了: 成功 %d 件 / 失敗 %d 件 / スキップ %s",
        success_count,
        failure_count,
        skipped_sheets if skipped_sheets else "なし",
    )
    return output_path


def _resolve_header_row(excel_section: object) -> int | None:
    """``[EXCEL] HEADER_ROW`` を ``int | None`` で返す。

    未設定（キーなし）または空文字のときは ``None`` を返し、
    ``convert_range_to_table`` の A2 ルール自動推定に任せる。
    """
    try:
        value = excel_section.HEADER_ROW  # type: ignore[attr-defined]
    except ConfigKeyNotFoundError:
        return None
    if value is None or value == "":
        return None
    return int(value)


def _sanitize_table_name(sheet_name: str) -> str:
    """シート名を Excel テーブル名として使える文字列へ整形する。

    制約は ``comken/toolbox/excel/sheet.py`` の ``_validate_table_name`` に従う:
        - 空文字は不可
        - 空白・禁則文字 ``[]/\\:*?\"<>|'`#%@$&+={}~`` は不可
        - 先頭が数字は不可（セル参照と紛らわしい）
        - セル参照パターン（``A1`` / ``R1C1``）は不可

    ここでは: 禁則文字を ``_`` に置換 → 先頭が数字またはセル参照パターンなら接頭辞を付与。
    """
    sanitized = _TABLE_NAME_FORBIDDEN_CHARS.sub("_", sheet_name)
    if not sanitized:
        return f"{_TABLE_NAME_DIGIT_PREFIX}DEFAULT"
    if sanitized[0].isdigit() or _TABLE_NAME_CELL_REF_PATTERN.fullmatch(sanitized):
        sanitized = f"{_TABLE_NAME_DIGIT_PREFIX}{sanitized}"
    return sanitized
