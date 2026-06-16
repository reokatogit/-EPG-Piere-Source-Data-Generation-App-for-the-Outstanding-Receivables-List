"""
fast_xlsx_reader.py - 巨大 xlsx ファイル用の高速リーダー

bulk re.findall で全セルを一括抽出し、DataFrame を直接構築する。
サブプロセス不要（C レベルの regex は GIL を保持しない I/O 部分で解放される）。
"""

from __future__ import annotations

import logging
import os
import re
import time
import zipfile
from typing import Callable

import numpy as np
import pandas as pd
from lxml import etree

logger = logging.getLogger(__name__)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ROW_RE = re.compile(rb"<row [^>]*?>(.*?)</row>", re.DOTALL)
CELL_RE = re.compile(
    rb'<c r="([A-Z]{1,3})\d+"([^>]*)>(?:.*?<v>([^<]*)</v>)?',
    re.DOTALL,
)

# 全セル一括抽出用の正規表現
ALL_CELLS_RE = re.compile(
    rb'<c r="([A-Z]{1,3})(\d+)"([^>]*)>(?:.*?<v>([^<]*)</v>)?',
    re.DOTALL,
)

T_RE = re.compile(rb"<t[^>]*>([^<]*)</t>")

LARGE_FILE_THRESHOLD = 30 * 1024 * 1024  # 30 MB


def _col_letter_to_index(letters: bytes) -> int:
    n = 0
    for c in letters:
        n = n * 26 + (c - 64)
    return n - 1


def _get_sheet_map(zf: zipfile.ZipFile) -> list[tuple[str, str, int]]:
    wb_xml = zf.read("xl/workbook.xml")
    root = etree.fromstring(wb_xml)
    ns = {"s": NS, "r": NS_REL}
    sheet_elems = root.findall(".//s:sheet", ns)

    rels_xml = zf.read("xl/_rels/workbook.xml.rels")
    rels_root = etree.fromstring(rels_xml)
    rid_to_file = {}
    for rel in rels_root:
        rid_to_file[rel.get("Id")] = "xl/" + rel.get("Target")

    result = []
    for s in sheet_elems:
        name = s.get("name")
        rid = s.get(f"{{{NS_REL}}}id")
        target = rid_to_file.get(rid, "")
        size = 0
        if target in zf.namelist():
            size = zf.getinfo(target).file_size
        result.append((name, target, size))
    return result


def _read_header_from_sheet(zf: zipfile.ZipFile, xml_path: str) -> list[str]:
    """シートXMLの先頭からヘッダ行をバイト正規表現で読み取る。"""
    with zf.open(xml_path) as sf:
        head = sf.read(200_000)

    row_match = ROW_RE.search(head)
    if not row_match:
        return []

    row_xml = row_match.group(1)
    needs_ss = b't="s"' in row_xml

    if not needs_ss:
        cells = {}
        for m in CELL_RE.finditer(row_xml):
            ci = _col_letter_to_index(m.group(1))
            val = m.group(3).decode("utf-8", errors="replace") if m.group(3) else ""
            cells[ci] = val
        ncols = max(cells.keys(), default=-1) + 1
        return [cells.get(i, "") for i in range(ncols)]

    ss_indices = set()
    for m in CELL_RE.finditer(row_xml):
        if b't="s"' in m.group(2) and m.group(3):
            ss_indices.add(int(m.group(3)))

    max_idx = max(ss_indices) if ss_indices else 0
    ss_partial = _load_shared_strings_partial(zf, max_idx + 1)

    cells = {}
    for m in CELL_RE.finditer(row_xml):
        ci = _col_letter_to_index(m.group(1))
        if m.group(3):
            if b't="s"' in m.group(2):
                idx = int(m.group(3))
                val = ss_partial[idx] if idx < len(ss_partial) else ""
            else:
                val = m.group(3).decode("utf-8", errors="replace")
        else:
            val = ""
        cells[ci] = val

    ncols = max(cells.keys(), default=-1) + 1
    return [cells.get(i, "") for i in range(ncols)]


def _load_shared_strings_partial(zf: zipfile.ZipFile, count: int) -> list[str]:
    """共有文字列テーブルの先頭 count 件だけ読み込む。"""
    si_tag = f"{{{NS}}}si"
    t_tag = f"{{{NS}}}t"
    ss: list[str] = []

    if "xl/sharedStrings.xml" not in zf.namelist():
        return ss

    with zf.open("xl/sharedStrings.xml") as ssf:
        for _, elem in etree.iterparse(ssf, events=("end",), tag=si_tag):
            parts = []
            for t in elem.iter(t_tag):
                if t.text:
                    parts.append(t.text)
            ss.append("".join(parts))
            elem.clear()
            while elem.getprevious() is not None:
                del elem.getparent()[0]
            if len(ss) >= count:
                break

    return ss


def _load_shared_strings_bulk(zf: zipfile.ZipFile, progress: Callable[[str], None] | None = None) -> list[str]:
    """共有文字列テーブルを高速バイト分割で全件読み込む。"""
    ss: list[str] = []

    if "xl/sharedStrings.xml" not in zf.namelist():
        return ss

    with zf.open("xl/sharedStrings.xml") as ssf:
        buf = b""
        while True:
            chunk = ssf.read(4 * 1024 * 1024)
            if not chunk:
                # 残りバッファ処理
                parts = buf.split(b"</si>")
                for part in parts[:-1]:
                    texts = T_RE.findall(part)
                    ss.append(b"".join(texts).decode("utf-8", errors="replace"))
                break
            buf += chunk
            parts = buf.split(b"</si>")
            buf = parts[-1]
            for part in parts[:-1]:
                texts = T_RE.findall(part)
                ss.append(b"".join(texts).decode("utf-8", errors="replace"))

            if progress and len(ss) % 200000 == 0 and len(ss) > 0:
                progress(f"文字列テーブル読込中... {len(ss):,} 件")

    return ss


def find_target_sheet(
    path: str,
    required_columns: list[str],
    find_column_fn,
    normalize_fn,
) -> tuple[str, str, int] | None:
    """必須列を含むシートを高速に特定する。複数候補がある場合は最大シートを優先。"""
    candidates = []

    with zipfile.ZipFile(path) as zf:
        sheet_map = _get_sheet_map(zf)

        for name, xml_path, size in sheet_map:
            if not xml_path or xml_path not in zf.namelist():
                continue

            headers = _read_header_from_sheet(zf, xml_path)
            if not headers:
                continue

            norm_headers = [normalize_fn(h) for h in headers]
            non_empty = [h for h in norm_headers if h.strip()]
            all_found = all(
                find_column_fn(non_empty, req) is not None
                for req in required_columns
            )
            if all_found:
                candidates.append((name, xml_path, size))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates[0]


def read_large_xlsx(
    path: str,
    required_columns: list[str],
    find_column_fn,
    normalize_fn,
    progress: Callable[[str], None] | None = None,
) -> pd.DataFrame:
    """
    巨大 xlsx を bulk re.findall で高速読み込み。
    全セルを一括抽出して DataFrame を直接構築する。
    """
    if progress:
        progress("ファイル構造を解析中...")

    target = find_target_sheet(path, required_columns, find_column_fn, normalize_fn)
    if target is None:
        raise ValueError("必須列を含むシートが見つかりません。ファイル形式を確認してください。")

    sheet_name, xml_path, total_size = target
    size_mb = total_size / 1024 / 1024
    if progress:
        progress(f"シート '{sheet_name}' を読み込みます（{size_mb:.0f} MB）")

    with zipfile.ZipFile(path) as zf:
        # 1. 共有文字列テーブル読み込み（高速バイト分割）
        if progress:
            progress("文字列テーブル読込中...")
        t0 = time.time()
        ss = _load_shared_strings_bulk(zf, progress)
        if progress:
            progress(f"文字列テーブル完了: {len(ss):,} 件 ({time.time()-t0:.0f}秒)")

        # 2. シートXMLを読み込んで全セルを一括抽出
        if progress:
            progress("データ読込中...")
        t1 = time.time()

        sheet_info = zf.getinfo(xml_path)
        file_size = sheet_info.file_size
        CHUNK = 8 * 1024 * 1024
        leftover = b""
        bytes_read = 0

        # 全セルを格納: col_indices[], row_nums[], values[]
        all_col_indices = []
        all_row_nums = []
        all_values = []
        cell_count = 0

        with zf.open(xml_path) as sf:
            while True:
                raw = sf.read(CHUNK)
                if not raw:
                    # 残りバッファ処理
                    if leftover:
                        matches = ALL_CELLS_RE.findall(leftover)
                        for col_letters, row_num, attrs, val_bytes in matches:
                            ci = _col_letter_to_index(col_letters)
                            ri = int(row_num)
                            if val_bytes:
                                if b't="s"' in attrs:
                                    idx = int(val_bytes)
                                    val = ss[idx] if idx < len(ss) else ""
                                else:
                                    val = val_bytes.decode("utf-8", errors="replace")
                            else:
                                val = ""
                            all_col_indices.append(ci)
                            all_row_nums.append(ri)
                            all_values.append(val)
                        cell_count += len(matches)
                    break

                data = leftover + raw
                bytes_read += len(raw)

                # 最後の完全な </row> で分割
                last_end = data.rfind(b"</row>")
                if last_end == -1:
                    leftover = data
                    continue
                last_end += len(b"</row>")
                process_part = data[:last_end]
                leftover = data[last_end:]

                # 一括抽出
                matches = ALL_CELLS_RE.findall(process_part)
                for col_letters, row_num, attrs, val_bytes in matches:
                    ci = _col_letter_to_index(col_letters)
                    ri = int(row_num)
                    if val_bytes:
                        if b't="s"' in attrs:
                            idx = int(val_bytes)
                            val = ss[idx] if idx < len(ss) else ""
                        else:
                            val = val_bytes.decode("utf-8", errors="replace")
                    else:
                        val = ""
                    all_col_indices.append(ci)
                    all_row_nums.append(ri)
                    all_values.append(val)

                cell_count += len(matches)

                if progress and file_size > 0:
                    pct = min(int(bytes_read / file_size * 100), 99)
                    progress(f"データ読込中... {cell_count:,} セル ({pct}%)")

        t_extract = time.time() - t1
        if progress:
            progress(f"セル抽出完了: {cell_count:,} セル ({t_extract:.0f}秒)")

        # 3. DataFrame 構築
        if progress:
            progress("DataFrame構築中...")
        t2 = time.time()

        if not all_row_nums:
            return pd.DataFrame()

        row_arr = np.array(all_row_nums, dtype=np.int32)
        col_arr = np.array(all_col_indices, dtype=np.int32)

        min_row = int(row_arr.min())
        max_row = int(row_arr.max())
        max_col = int(col_arr.max())
        num_cols = max_col + 1

        # ヘッダ行 (row=min_row) とデータ行を分離
        header_mask = row_arr == min_row
        header_cols = col_arr[header_mask]
        header_vals = [all_values[i] for i in range(len(all_values)) if header_mask[i]]

        headers = [""] * num_cols
        for ci, val in zip(header_cols, header_vals):
            headers[ci] = val

        # データ行を構築
        data_mask = ~header_mask
        data_rows = row_arr[data_mask]
        data_cols = col_arr[data_mask]
        data_vals_indices = np.where(data_mask)[0]

        if len(data_rows) == 0:
            df = pd.DataFrame(columns=[normalize_fn(h) for h in headers])
            if progress:
                progress(f"ファイル読込完了: 0 件")
            return df

        # 行番号を 0-based の連続インデックスに変換
        unique_rows = np.unique(data_rows)
        num_data_rows = len(unique_rows)
        row_to_idx = {}
        for i, r in enumerate(unique_rows):
            row_to_idx[r] = i

        # 列ごとに配列を構築（メモリ効率的）
        if progress:
            progress(f"DataFrame構築中... {num_data_rows:,} 行 × {num_cols} 列")

        # 空文字列の numpy 配列で初期化
        col_arrays = [np.empty(num_data_rows, dtype=object) for _ in range(num_cols)]
        for ca in col_arrays:
            ca[:] = ""

        # データを列配列に格納
        for k in range(len(data_rows)):
            ri = row_to_idx[data_rows[k]]
            ci = data_cols[k]
            col_arrays[ci][ri] = all_values[data_vals_indices[k]]

        # DataFrame 構築
        col_dict = {}
        normalized_headers = [normalize_fn(h) for h in headers]
        for ci in range(num_cols):
            col_name = normalized_headers[ci] if ci < len(normalized_headers) else f"col_{ci}"
            # 重複列名の処理
            if col_name in col_dict:
                suffix = 2
                while f"{col_name}_{suffix}" in col_dict:
                    suffix += 1
                col_name = f"{col_name}_{suffix}"
            col_dict[col_name] = col_arrays[ci]

        df = pd.DataFrame(col_dict)

        t_build = time.time() - t2
        total_time = time.time() - t0
        if progress:
            progress(f"ファイル読込完了: {len(df):,} 件 (合計 {total_time:.0f}秒)")

        return df


def is_large_xlsx(path: str) -> bool:
    _, ext = os.path.splitext(path)
    if ext.lower() not in (".xlsx", ".xlsm"):
        return False
    return os.path.getsize(path) > LARGE_FILE_THRESHOLD
