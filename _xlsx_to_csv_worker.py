"""
_xlsx_to_csv_worker.py - xlsx → CSV 変換ワーカー（サブプロセスで実行）

bulk re.findall で全セルを一括抽出し、高速に CSV を生成する。
進捗を stdout に出力し、親プロセスが読み取ってGUIに表示する。
出力形式: PROGRESS:メッセージ  または  ERROR:メッセージ
"""
import sys
import os
import zipfile
import re
import csv
import time

T_RE = re.compile(rb"<t[^>]*>([^<]*)</t>")
ALL_CELLS_RE = re.compile(
    rb'<c r="([A-Z]{1,3})(\d+)"([^>]*)>(?:.*?<v>([^<]*)</v>)?',
    re.DOTALL,
)


def col_idx(letters: bytes) -> int:
    n = 0
    for c in letters:
        n = n * 26 + (c - 64)
    return n - 1


def progress(msg: str) -> None:
    print(f"PROGRESS:{msg}", flush=True)


def main():
    if len(sys.argv) != 4:
        print("ERROR:Usage: _xlsx_to_csv_worker.py <xlsx_path> <sheet_xml> <output_csv>", flush=True)
        sys.exit(1)

    xlsx_path = sys.argv[1]
    sheet_xml = sys.argv[2]
    output_csv = sys.argv[3]

    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            # 1. 共有文字列テーブル読み込み（byte split方式）
            progress("文字列テーブル読込中...")
            t0 = time.time()
            ss = _load_shared_strings(zf)
            progress(f"文字列テーブル完了: {len(ss):,} 件 ({time.time()-t0:.0f}秒)")

            # 2. シートデータをチャンク単位で bulk findall → CSV
            progress("データ変換中...")
            sheet_info = zf.getinfo(sheet_xml)
            total_size = sheet_info.file_size
            t1 = time.time()

            row_count = _extract_to_csv(zf, sheet_xml, output_csv, ss, total_size)

            elapsed = time.time() - t1
            progress(f"CSV変換完了: {row_count:,} 行 ({elapsed:.0f}秒)")

    except Exception as e:
        print(f"ERROR:{e}", flush=True)
        sys.exit(1)


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """共有文字列テーブルを byte split 方式で高速読み込み。"""
    ss: list[str] = []
    if "xl/sharedStrings.xml" not in zf.namelist():
        return ss

    with zf.open("xl/sharedStrings.xml") as ssf:
        buf = b""
        while True:
            chunk = ssf.read(4 * 1024 * 1024)
            if not chunk:
                for part in buf.split(b"</si>")[:-1]:
                    texts = T_RE.findall(part)
                    ss.append(b"".join(texts).decode("utf-8", errors="replace"))
                break
            buf += chunk
            parts = buf.split(b"</si>")
            buf = parts[-1]
            for part in parts[:-1]:
                texts = T_RE.findall(part)
                ss.append(b"".join(texts).decode("utf-8", errors="replace"))
            if len(ss) % 500000 == 0 and len(ss) > 0:
                progress(f"文字列テーブル読込中... {len(ss):,} 件")
    return ss


def _extract_to_csv(
    zf: zipfile.ZipFile,
    sheet_xml: str,
    output_csv: str,
    ss: list[str],
    total_size: int,
) -> int:
    """
    チャンク単位で sheet XML を読み、bulk findall で全セルを抽出して CSV に書き出す。
    50MB チャンクごとに findall → 行グループ化 → CSV書き出し。
    """
    CHUNK_SIZE = 50 * 1024 * 1024  # 50MB chunks for bulk processing
    bytes_read = 0
    row_count = 0
    num_cols = 0

    with zf.open(sheet_xml) as sf, \
         open(output_csv, "w", newline="", encoding="utf-8") as csvf:
        writer = csv.writer(csvf)
        leftover = b""

        while True:
            raw = sf.read(CHUNK_SIZE)
            if not raw:
                # 残りデータを処理
                if leftover.strip():
                    cnt, num_cols = _process_chunk(leftover, ss, writer, num_cols)
                    row_count += cnt
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

            cnt, num_cols = _process_chunk(process_part, ss, writer, num_cols)
            row_count += cnt

            if total_size > 0:
                pct = min(int(bytes_read / total_size * 100), 99)
                progress(f"データ読込中... {row_count:,} 行 ({pct}%)")

    return row_count


def _process_chunk(
    data: bytes,
    ss: list[str],
    writer,
    num_cols: int,
) -> tuple[int, int]:
    """
    チャンク内の全セルを bulk findall で一括抽出し、行ごとにCSVに書き出す。
    戻り値: (このチャンクの行数, 更新後の最大列数)
    """
    matches = ALL_CELLS_RE.findall(data)
    if not matches:
        return 0, num_cols

    # マッチ結果を行番号でグループ化
    rows: dict[int, list[tuple[int, str]]] = {}
    max_ci = num_cols - 1

    for col_letters, row_num_bytes, attrs, val_bytes in matches:
        row_num = int(row_num_bytes)
        ci = col_idx(col_letters)

        if val_bytes:
            if b't="s"' in attrs:
                idx = int(val_bytes)
                val = ss[idx] if idx < len(ss) else ""
            else:
                val = val_bytes.decode("utf-8", errors="replace")
        else:
            val = ""

        if row_num not in rows:
            rows[row_num] = []
        rows[row_num].append((ci, val))
        if ci > max_ci:
            max_ci = ci

    new_num_cols = max_ci + 1

    # 行番号順にソートして CSV 出力
    for row_num in sorted(rows):
        cells = rows[row_num]
        row = [""] * new_num_cols
        for ci, val in cells:
            row[ci] = val
        writer.writerow(row)

    return len(rows), new_num_cols


if __name__ == "__main__":
    main()
