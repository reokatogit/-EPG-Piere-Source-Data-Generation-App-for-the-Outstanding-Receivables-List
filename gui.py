"""
gui.py - tkinter GUI
Piere未収リスト元データ作成ツール
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import processor
from processor import ProcessingResult


class Application(tk.Tk):
    """メインアプリケーションウィンドウ。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("Piere未収リスト元データ作成ツール")
        self.geometry("720x620")
        self.minsize(640, 540)
        self.resizable(True, True)

        self._msg_queue: queue.Queue = queue.Queue()
        self._processing = False
        self._configure_style()
        self._build_ui()
        self._set_default_month()

    # ── スタイル ──────────────────────────────────────────

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")

        style.configure("Title.TLabel", font=("Yu Gothic UI", 14, "bold"))
        style.configure("RunButton.TButton", font=("Yu Gothic UI", 11, "bold"), padding=8)
        style.configure("Status.TLabel", font=("Yu Gothic UI", 9), foreground="#555555")

    # ── UI 構築 ───────────────────────────────────────────

    def _build_ui(self) -> None:
        root_frame = ttk.Frame(self, padding=12)
        root_frame.pack(fill=tk.BOTH, expand=True)

        # タイトル
        ttk.Label(root_frame, text="Piere未収リスト元データ作成ツール", style="Title.TLabel").pack(
            pady=(0, 10)
        )

        # ── 入力設定 ──────────────────────────────────
        input_frame = ttk.LabelFrame(root_frame, text="入力設定", padding=10)
        input_frame.pack(fill=tk.X, pady=(0, 8))
        input_frame.columnconfigure(1, weight=1)

        # 今月分 CSV
        self.csv_var = tk.StringVar()
        self._add_file_row(input_frame, row=0, label="今月分ファイル:", var=self.csv_var,
                           command=self._select_csv)

        # 前月分完成 Excel
        self.prev_excel_var = tk.StringVar()
        self._add_file_row(input_frame, row=1, label="前月分完成Excel:", var=self.prev_excel_var,
                           command=self._select_prev_excel)

        # 処理対象年月
        ttk.Label(input_frame, text="処理対象年月:").grid(row=2, column=0, sticky=tk.W, pady=4)
        month_frame = ttk.Frame(input_frame)
        month_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=(6, 0))

        self.year_var = tk.StringVar()
        ttk.Spinbox(month_frame, from_=2020, to=2099, textvariable=self.year_var,
                    width=7, font=("Yu Gothic UI", 9)).pack(side=tk.LEFT)
        ttk.Label(month_frame, text="年").pack(side=tk.LEFT, padx=(2, 10))

        self.month_var = tk.StringVar()
        ttk.Combobox(
            month_frame,
            textvariable=self.month_var,
            values=[str(i) for i in range(1, 13)],
            width=5,
            state="readonly",
            font=("Yu Gothic UI", 9),
        ).pack(side=tk.LEFT)
        ttk.Label(month_frame, text="月").pack(side=tk.LEFT, padx=(2, 0))

        # 出力先フォルダ
        self.output_folder_var = tk.StringVar()
        self._add_folder_row(input_frame, row=3, label="出力先フォルダ:", var=self.output_folder_var,
                             command=self._select_output_folder)

        # 出力ファイル名
        ttk.Label(input_frame, text="出力ファイル名:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.output_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.output_name_var, font=("Yu Gothic UI", 9)).grid(
            row=4, column=1, columnspan=2, sticky=tk.EW, padx=(6, 4)
        )

        # ── 実行ボタン ───────────────────────────────
        btn_frame = ttk.Frame(root_frame)
        btn_frame.pack(pady=8)
        self.run_btn = ttk.Button(
            btn_frame, text="  実行  ", style="RunButton.TButton", command=self._run
        )
        self.run_btn.pack(ipadx=16, ipady=4)

        # プログレスバー
        self.progress_bar = ttk.Progressbar(root_frame, mode="determinate", length=400, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 6))

        # プログレス表示ラベル
        self.progress_label = ttk.Label(root_frame, text="", font=("Yu Gothic UI", 8))
        self.progress_label.pack(anchor=tk.E, padx=2, pady=(0, 2))

        # ── 処理結果 ─────────────────────────────────
        result_frame = ttk.LabelFrame(root_frame, text="処理ログ", padding=8)
        result_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            result_frame,
            state="disabled",
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#f8f8f8",
            height=12,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ステータスバー
        self.status_var = tk.StringVar(value="待機中")
        ttk.Label(root_frame, textvariable=self.status_var, style="Status.TLabel").pack(
            anchor=tk.W, pady=(4, 0)
        )

    def _add_file_row(self, parent, row: int, label: str, var: tk.StringVar,
                      command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=4)
        ttk.Entry(parent, textvariable=var, font=("Yu Gothic UI", 9)).grid(
            row=row, column=1, sticky=tk.EW, padx=(6, 4)
        )
        ttk.Button(parent, text="参照...", command=command).grid(row=row, column=2)

    def _add_folder_row(self, parent, row: int, label: str, var: tk.StringVar,
                        command) -> None:
        self._add_file_row(parent, row, label, var, command)

    # ── デフォルト年月 ────────────────────────────────────

    def _set_default_month(self) -> None:
        now = datetime.now()
        self.year_var.set(str(now.year))
        self.month_var.set(str(now.month))
        self._update_default_filename()
        # 年月変更時にファイル名を自動更新
        self.year_var.trace_add("write", lambda *_: self._update_default_filename())
        self.month_var.trace_add("write", lambda *_: self._update_default_filename())

    def _update_default_filename(self) -> None:
        """年月に連動してデフォルトファイル名を更新する。"""
        try:
            y = int(self.year_var.get())
            m = int(self.month_var.get())
            self.output_name_var.set(f"【Piere契約有効案件{y}{m:02d}まで】nv_contract_v_SJIS")
        except (ValueError, TypeError):
            pass

    # ── ファイル選択 ─────────────────────────────────────

    def _select_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="今月分ファイルを選択（CSV または Excel）",
            filetypes=[("CSV/Excelファイル", "*.csv *.xlsx *.xlsm"), ("CSVファイル", "*.csv"), ("Excelファイル", "*.xlsx *.xlsm"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.csv_var.set(path)

    def _select_prev_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="前月分完成Excelを選択",
            filetypes=[("Excelファイル", "*.xlsx *.xlsm *.xls"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.prev_excel_var.set(path)

    def _select_output_folder(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.output_folder_var.set(path)

    # ── バリデーション ────────────────────────────────────

    def _validate_inputs(self) -> bool:
        csv = self.csv_var.get().strip()
        prev_excel = self.prev_excel_var.get().strip()
        output_folder = self.output_folder_var.get().strip()

        if not csv:
            messagebox.showerror("入力エラー", "今月分ファイル（CSV または Excel）を選択してください。")
            return False
        if not os.path.isfile(csv):
            messagebox.showerror("入力エラー", f"今月分ファイルが見つかりません:\n{csv}")
            return False
        if not prev_excel:
            messagebox.showerror("入力エラー", "前月分完成Excelを選択してください。")
            return False
        if not os.path.isfile(prev_excel):
            messagebox.showerror("入力エラー", f"前月分完成Excelが見つかりません:\n{prev_excel}")
            return False
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            if not (2000 <= year <= 2099 and 1 <= month <= 12):
                raise ValueError
        except ValueError:
            messagebox.showerror("入力エラー", "処理対象年月を正しく入力してください。")
            return False
        if not output_folder:
            messagebox.showerror("入力エラー", "出力先フォルダを選択してください。")
            return False
        return True

    # ── ログ出力 ─────────────────────────────────────────
    def _log(self, msg: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _flush_log_buffer(self) -> None:
        pass  # バッファなし、_logが即時表示

    def _clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")

    def _set_status(self, msg: str) -> None:
        self.status_var.set(msg)

    def _poll_queue(self) -> None:
        """50ms ごとにキューを確認してGUIを更新（メインスレッドで実行）。"""
        try:
            while True:
                kind, data = self._msg_queue.get_nowait()
                if kind == "progress":
                    self._update_progress(data)
                elif kind == "complete":
                    self._processing = False
                    self._on_complete(data)
                    return
                elif kind == "fatal_error":
                    self._processing = False
                    self._on_fatal_error(data)
                    return
                elif kind == "cancelled":
                    self._processing = False
                    self._on_cancelled()
                    return
                elif kind == "exception":
                    self._processing = False
                    self._on_exception(*data)
                    return
        except queue.Empty:
            pass
        if self._processing:
            self.after(50, self._poll_queue)

    def _update_progress(self, msg: str) -> None:
        """進捗メッセージから進捗値を抽出してバーを更新。"""
        self._log(msg)

        # データ読込中... X 行 (Y%) パターン（大容量xlsx）
        match = re.search(r'データ読込中.*?(\d+)%', msg)
        if match:
            pct = int(match.group(1))
            # ファイル読込は全体の 0-60% に割り当て
            progress_pct = int(pct * 0.6)
            self.progress_bar["value"] = progress_pct
            self.progress_label.config(text=f"{progress_pct}%")
            return

        # 文字列テーブル読込中
        if "文字列テーブル読込中" in msg:
            self.progress_bar["value"] = 5
            self.progress_label.config(text="5%")
            return

        if "文字列テーブル完了" in msg:
            self.progress_bar["value"] = 10
            self.progress_label.config(text="10%")
            return

        # 除外判定完了: 除外 {excluded} 件 / 残 {output} 件 パターン
        match = re.search(r'除外 (\d+) 件 / 残 (\d+) 件', msg)
        if match:
            excluded = int(match.group(1))
            output = int(match.group(2))
            total = excluded + output
            if total > 0:
                progress_pct = int((excluded / total) * 50) + 50  # 50-100%
                self.progress_bar["value"] = progress_pct
                self.progress_label.config(text=f"{progress_pct}%")
            return

        # ファイル読込完了: {total_count} 件 パターン
        match = re.search(r'ファイル読込完了.*?(\d[\d,]*)\s*件', msg)
        if match:
            self.progress_bar["value"] = 65
            self.progress_label.config(text="65%")
            return

        # 前月分Excel読込中、除外判定中などのステップを認識
        if "前月分Excel読込中" in msg:
            self.progress_bar["value"] = 68
            self.progress_label.config(text="68%")
        elif "除外判定中" in msg:
            self.progress_bar["value"] = 75
            self.progress_label.config(text="75%")
        elif "Excelワークブック生成中" in msg:
            self.progress_bar["value"] = 85
            self.progress_label.config(text="85%")
        elif "ファイル保存中" in msg:
            self.progress_bar["value"] = 95
            self.progress_label.config(text="95%")

    # ── 実行制御 ─────────────────────────────────────────

    def _run(self) -> None:
        if not self._validate_inputs():
            return

        self._clear_log()
        self.run_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="")
        self._set_status("処理中...")
        self._processing = True
        self._poll_queue()  # キューポーリング開始

        csv_path = self.csv_var.get().strip()
        prev_excel_path = self.prev_excel_var.get().strip()
        target_year = int(self.year_var.get())
        target_month = int(self.month_var.get())
        output_folder = self.output_folder_var.get().strip()
        output_name = self.output_name_var.get().strip()

        self._log(f"処理開始: {target_year}年{target_month}月")
        self._log(f"入力ファイル  : {os.path.basename(csv_path)}")
        self._log(f"前月分Excel  : {os.path.basename(prev_excel_path)}")
        self._log(f"出力先       : {output_folder}")
        self._log("-" * 55)

        def worker() -> None:
            try:
                result = processor.process_data(
                    csv_path=csv_path,
                    prev_excel_path=prev_excel_path,
                    target_year=target_year,
                    target_month=target_month,
                    progress_callback=lambda m: self._msg_queue.put(("progress", m)),
                )

                if result.errors:
                    self._msg_queue.put(("fatal_error", result))
                    return

                # 警告がある場合、メインスレッドで確認ダイアログを表示
                if result.warnings:
                    confirmed_event = threading.Event()
                    user_ok: list[bool] = [False]

                    def ask_warnings() -> None:
                        warn_text = "\n".join(f"・{w}" for w in result.warnings)
                        msg = (
                            f"以下の警告が発生しました。\n\n{warn_text}\n\n"
                            "処理結果を保存しますか？"
                        )
                        user_ok[0] = messagebox.askyesno("警告確認", msg, icon="warning")
                        confirmed_event.set()

                    self.after(0, ask_warnings)
                    confirmed_event.wait()

                    if not user_ok[0]:
                        self._msg_queue.put(("cancelled", None))
                        return

                # 上書き確認
                output_path = os.path.join(output_folder, output_name + ".xlsx")
                if os.path.exists(output_path):
                    overwrite_event = threading.Event()
                    overwrite_ok: list[bool] = [False]

                    def ask_overwrite() -> None:
                        overwrite_ok[0] = messagebox.askyesno(
                            "上書き確認",
                            f"同名ファイルが既に存在します。\n\n{os.path.basename(output_path)}\n\n上書きしますか？",
                            icon="warning",
                        )
                        overwrite_event.set()

                    self.after(0, ask_overwrite)
                    overwrite_event.wait()

                    if not overwrite_ok[0]:
                        self._msg_queue.put(("cancelled", None))
                        return

                # 保存
                self._msg_queue.put(("progress", "ファイル保存中..."))
                err = processor.save_result(result, output_folder, output_name)
                if err:
                    result.errors.append(err)
                    self._msg_queue.put(("fatal_error", result))
                    return

                self._msg_queue.put(("complete", result))

            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                self._msg_queue.put(("exception", (str(e), tb)))

        threading.Thread(target=worker, daemon=True).start()

    # ── 完了・エラーハンドラ ─────────────────────────────

    def _on_complete(self, result: ProcessingResult) -> None:
        self._stop_progress()
        self._log("-" * 55)
        self._log("【処理完了】")
        self._log(f"  読込件数    : {result.total_count} 件")
        self._log(f"  除外件数    : {result.excluded_count} 件")
        self._log(f"  出力件数    : {result.output_count} 件")
        if result.tab_counts:
            for sheet, cnt in result.tab_counts.items():
                self._log(f"    └ {sheet}: {cnt} 件")
        self._log(f"  生成ファイル: {os.path.basename(result.output_file)}")
        self._flush_log_buffer()
        self._set_status("完了")

        summary = (
            f"処理が完了しました。\n\n"
            f"読込: {result.total_count} 件 / 除外: {result.excluded_count} 件 / 出力: {result.output_count} 件\n\n"
            f"生成ファイル:\n{result.output_file}"
        )
        if messagebox.askokcancel("処理完了", summary + "\n\nファイルを開きますか？", icon="info"):
            self._open_file(result.output_file)

    def _on_fatal_error(self, result: ProcessingResult) -> None:
        self._stop_progress()
        self._log("-" * 55)
        self._log("【エラー】処理を中断しました。")
        for err in result.errors:
            self._log(f"  ✗ {err}")
        self._flush_log_buffer()  # バッファをフラッシュ
        self._set_status("エラー終了")
        messagebox.showerror("処理エラー", "\n".join(result.errors))

    def _on_cancelled(self) -> None:
        self._stop_progress()
        self._log("処理をキャンセルしました。ファイルは保存されていません。")
        self._flush_log_buffer()  # バッファをフラッシュ
        self._set_status("キャンセル")

    def _on_exception(self, msg: str, tb: str) -> None:
        self._stop_progress()
        self._log(f"予期しないエラー: {msg}")
        self._log(tb)
        self._flush_log_buffer()  # バッファをフラッシュ
        self._set_status("エラー終了")
        messagebox.showerror("予期しないエラー", f"エラーが発生しました:\n\n{msg}\n\n詳細はログを確認してください。")

    def _stop_progress(self) -> None:
        self.progress_bar["value"] = 100
        self.progress_label.config(text="100%")
        self.run_btn.config(state="normal")

    # ── ファイルを開く ────────────────────────────────────

    def _open_file(self, path: str) -> None:
        if sys.platform == "win32":
            os.startfile(path)  # noqa: S606
        else:
            subprocess.run(["xdg-open", path], check=False)
