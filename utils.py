"""
utils.py - 列名正規化・日付解析ユーティリティ
"""
import unicodedata
import re
import datetime
import pandas as pd


# ==================== 列名マッピング ====================

COLUMN_ALIASES: dict[str, list[str]] = {
    "ステータス": ["ステータス", "status", "Status", "ｽﾃｰﾀｽ"],
    "明細ステータス": ["明細ステータス", "明細status", "明細Status", "明細ｽﾃｰﾀｽ"],
    "商品名": ["商品名", "商品　名", "商 品 名", "item_name"],
    "商品グループ": ["商品グループ", "商品group", "商品Group", "商品ｸﾞﾙｰﾌﾟ", "商品ｸﾞﾙｰﾌﾟ名"],
    "単価（税抜）": [
        "単価（税抜）", "単価(税抜)", "単価（税抜き）", "単価(税抜き)",
        "税抜単価", "単価（税抜額）", "単価(税抜額)",
    ],
    "請求開始月": ["請求開始月", "請求開始", "課金開始月", "請求　開始月"],
    "請求終了月": ["請求終了月", "請求終了", "課金終了月", "請求　終了月"],
}

# 追加管理列の内部名 → 表示名
EXTRA_COLS_INTERNAL = [
    "__blank1__",
    "除外理由",
    "__blank2__",
    "契約有無",
    "除外",
    "除外番号",
    "単価0円案件判別",
]

EXTRA_COLS_DISPLAY = ["", "除外理由", "", "契約有無", "除外", "除外番号", "単価0円案件判別"]

INTERNAL_TO_DISPLAY: dict[str, str] = {
    "__blank1__": "",
    "__blank2__": "",
}


# ==================== 列名正規化 ====================

def normalize_col_name(name: str) -> str:
    """列名を正規化する（全角→半角、前後空白除去、不可視文字除去）。"""
    name = str(name).strip()
    name = unicodedata.normalize("NFKC", name)
    # ゼロ幅スペース等の不可視文字を除去
    name = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", name)
    return name


def find_column(df_columns, canonical_name: str) -> str | None:
    """DataFrameの列リストから正規名に対応する実際の列名を返す。見つからなければ None。"""
    normalized_cols = {normalize_col_name(c): c for c in df_columns}

    aliases = COLUMN_ALIASES.get(canonical_name, [canonical_name])
    for alias in aliases:
        normalized_alias = normalize_col_name(alias)
        if normalized_alias in normalized_cols:
            return normalized_cols[normalized_alias]

    # 部分一致フォールバック（正規名が列名に含まれる、またはその逆）
    canonical_normalized = normalize_col_name(canonical_name)
    for norm_col, orig_col in normalized_cols.items():
        if canonical_normalized in norm_col or norm_col in canonical_normalized:
            return orig_col

    return None


# ==================== 日付・月解析 ====================

def parse_month(value) -> tuple[int, int] | None:
    """
    各種形式の月表現を解析して (year, month) タプルを返す。
    解析できない場合は None を返す。
    対応形式:
      - YYYY/MM, YYYY-MM, YYYY年MM月
      - YYYY/MM/DD, YYYY年MM月DD日
      - YYYYMM (6桁数字)
      - Excelシリアル値 (数値)
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    value_str = str(value).strip()
    if not value_str or value_str.lower() in ("nan", "none", "nat", ""):
        return None

    value_str = unicodedata.normalize("NFKC", value_str)

    # YYYY/MM または YYYY-MM または YYYY年MM月 (+オプションで日付)
    m = re.match(r"(\d{4})[/\-年](\d{1,2})(?:[/\-月日]|$)", value_str)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return (y, mo)

    # YYYYMM (6桁数字のみ)
    m = re.match(r"^(\d{4})(\d{2})$", value_str)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return (y, mo)

    # Excelシリアル値（数値）
    try:
        val = float(value_str.replace(",", ""))
        if 1 <= val <= 2958465:  # 1900/01/01 〜 9999/12/31
            excel_epoch = datetime.date(1899, 12, 30)
            delta = datetime.timedelta(days=int(val))
            d = excel_epoch + delta
            return (d.year, d.month)
    except (ValueError, TypeError):
        pass

    return None


def month_to_int(year_month: tuple[int, int] | None) -> int | None:
    """(year, month) を YYYYMM 整数に変換する。None の場合は None を返す。"""
    if year_month is None:
        return None
    return year_month[0] * 100 + year_month[1]


def prev_month(year_month: tuple[int, int]) -> tuple[int, int]:
    """前月の (year, month) を返す。"""
    y, m = year_month
    if m == 1:
        return (y - 1, 12)
    return (y, m - 1)
