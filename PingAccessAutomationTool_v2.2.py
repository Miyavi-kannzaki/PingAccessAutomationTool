import customtkinter as ctk
import tkinter as tk
import requests
import pyperclip
from datetime import datetime
import os
import subprocess
import threading
import time
import re
import requests.exceptions
from concurrent.futures import ThreadPoolExecutor, as_completed

MAX_LINES = 10
MAX_WORKERS = 10 # 並列処理で同時に実行する最大タスク数

# --- アプリケーション設定定数 ---
RT_HUB_BASE_PORT = 50000
AP_BASE_PORT = 60000
# デフォルトのタイムアウト値（GUIでの入力が不正だった場合に使われる）
DEFAULT_TIMEOUT_SEC = 3
DEFAULT_HUB_TIMEOUT_SEC = 7


def open_url_in_chrome_force_tab(url, mainapp_instance):
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    chrome_path = None
    for path in chrome_paths:
        if os.path.exists(path):
            chrome_path = path
            break
    if chrome_path:
        try:
            subprocess.Popen([chrome_path, "--new-tab", url])
            return True
        except OSError as e:
            # --- 変更点1：エラーをGUIログに出力 ---
            mainapp_instance.append_log(f"Chrome起動失敗: {e}", level="warn")
            return False
    else:
        # --- 変更点1：エラーをGUIログに出力 ---
        mainapp_instance.append_log("Chromeが見つかりません", level="warn")
        return False

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("疎通確認ツール ver2.2 (並列処理・停止機能付き)")
        self.geometry("1130x700")
        ctk.set_appearance_mode("dark")

        # --- 変更点3：緊急停止用のイベントフラグ ---
        self.stop_event = threading.Event()

        self.base_bg = "#212121"
        self.frame_border = "#424242"
        self.card_bg = "#232C33"
        self.section_bg = "#263238"
        self.log_bg = "#222B33"
        self.log_fg = "#ECEFF1"

        self.log_tags = (
            "success", "fail", "warn", "info", "rt_success", "hub_success",
            "cleared", "summary_rt", "summary_hub", "summary_ap"
        )

        self.configure(fg_color=self.base_bg)

        self.lines_var = tk.IntVar(value=1)
        self.lines_var.trace_add("write", self.update_lines_count)

        top_frame = ctk.CTkFrame(self, fg_color=self.base_bg)
        top_frame.pack(fill="x", pady=(8, 2), padx=10)
        ctk.CTkLabel(top_frame, text="回線本数", width=80).pack(side="left", padx=5)
        self.lines_entry = ctk.CTkEntry(top_frame, textvariable=self.lines_var, width=50)
        self.lines_entry.pack(side="left")
        ctk.CTkLabel(top_frame, text="（1〜10）", width=60).pack(side="left", padx=5)

        self.access_mode = tk.StringVar(value="browser")
        access_frame = ctk.CTkFrame(self, fg_color=self.base_bg)
        access_frame.pack(fill="x", pady=(2, 6), padx=10)
        ctk.CTkLabel(access_frame, text="アクセス方式", width=100).pack(side="left", padx=(0, 5))
        ctk.CTkRadioButton(access_frame, text="web", variable=self.access_mode, value="browser").pack(side="left")
        ctk.CTkRadioButton(access_frame, text="requests", variable=self.access_mode, value="requests").pack(side="left")

        main_frame = ctk.CTkFrame(self, fg_color=self.base_bg)
        main_frame.pack(fill="both", expand=True, padx=10, pady=(2, 2))

        self.left_card = ctk.CTkFrame(main_frame, fg_color=self.card_bg, border_width=0, corner_radius=18, width=770)
        self.left_card.pack(side="left", fill="both", padx=(0, 12), pady=8)
        self.left_card.pack_propagate(False)

        self.tabview = ctk.CTkTabview(self.left_card, fg_color=self.card_bg, width=730, height=600)
        self.tabview.pack(fill="both", expand=True, padx=(10, 10), pady=(16, 16))
        
        self.line_frames = []

        sidebar = ctk.CTkFrame(main_frame, fg_color=self.log_bg, width=340, corner_radius=14)
        sidebar.pack(side="right", fill="both", expand=True, padx=(2, 4), pady=8)
        sidebar.pack_propagate(False)

        self.log_marker_var = tk.StringVar(value="")
        self.log_marker_label = ctk.CTkLabel(sidebar, textvariable=self.log_marker_var,
                                             text_color="#FFEB3B", font=ctk.CTkFont(size=15, weight="bold"))
        self.log_marker_label.pack(anchor="n", pady=(16, 3))

        log_title_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        log_title_frame.pack(anchor="n", fill="x", pady=(5, 2), padx=8)

        log_label = ctk.CTkLabel(log_title_frame, text="=== LOG ===", text_color="#90CAF9",
                                 font=ctk.CTkFont(size=13, weight="bold"))
        log_label.pack(side="left", anchor="w")

        clear_log_btn = ctk.CTkButton(log_title_frame, text="ログクリア",
                                      command=self.clear_log,
                                      width=80, height=20,
                                      fg_color="#616161", hover_color="#757575")
        clear_log_btn.pack(side="right", anchor="e")
        
        ctk.CTkFrame(sidebar, height=1, fg_color="#405060").pack(fill="x", padx=8, pady=(1, 8))

        log_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        log_frame.pack(fill="both", padx=2, pady=3, expand=True)
        self.log_text = tk.Text(
            log_frame, height=22, state="disabled", font=("Consolas", 11),
            bg=self.log_bg, fg=self.log_fg, insertbackground="#ECEFF1",
            wrap="none", width=36, relief="flat", borderwidth=0
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        log_x_scroll = ctk.CTkScrollbar(log_frame, orientation="horizontal", command=self.log_text.xview)
        log_x_scroll.pack(side="bottom", fill="x")
        self.log_text.config(xscrollcommand=log_x_scroll.set)
        
        self.log_text.bind("<Shift-MouseWheel>", self._on_horizontal_scroll)

        self.log_text.tag_config("success", foreground="#8BC34A")
        self.log_text.tag_config("fail", foreground="#E57373")
        self.log_text.tag_config("warn", foreground="#FFA726")
        self.log_text.tag_config("info", foreground="#ECEFF1")
        self.log_text.tag_config("rt_success", foreground="#FFA726")
        self.log_text.tag_config("hub_success", foreground="#4FC3F7")
        self.log_text.tag_config("cleared", foreground="#B0BEC5")
        self.log_text.tag_config("summary_rt", foreground="#4FC3F7")
        self.log_text.tag_config("summary_hub", foreground="#81C784")
        self.log_text.tag_config("summary_ap", foreground="#FFD54F")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(sidebar, variable=self.progress_var, width=220, height=10)
        self.progress_bar.pack(pady=(12, 10), anchor="s")

        # --- 変更点3：緊急停止ボタンの追加 ---
        self.stop_button = ctk.CTkButton(sidebar, text="緊急停止", command=self.request_stop,
                                         fg_color="#C62828", hover_color="#B71C1C", width=210, state="disabled")
        self.stop_button.pack(anchor="s", pady=(4, 4))
        
        exit_btn = ctk.CTkButton(sidebar, text="アプリ終了", command=self.on_exit,
                                 fg_color="#424242", hover_color="#607D8B", width=210)
        exit_btn.pack(anchor="s", pady=(4, 16))

        self.update_lines_count()
    
    def _on_horizontal_scroll(self, event):
        if event.delta > 0:
            self.log_text.xview_scroll(-2, "units")
        else:
            self.log_text.xview_scroll(2, "units")

    # --- 変更点3：緊急停止ボタンのコマンド ---
    def request_stop(self):
        self.append_log("停止リクエスト受信。現在の処理が完了次第、中断します...", level="warn")
        self.stop_event.set()
        self.stop_button.configure(state="disabled", text="停止中...")

    def update_lines_count(self, *_):
        try:
            n = int(self.lines_var.get())
        except ValueError:
            n = 1
        n = max(1, min(MAX_LINES, n))
        while len(self.line_frames) < n:
            line_num = len(self.line_frames) + 1
            tab = self.tabview.add(f"回線#{line_num}")
            self.tabview.tab(f"回線#{line_num}").configure(fg_color=self.card_bg)
            frame = LineTabFrame(tab, line_num, self)
            frame.pack(fill="both", expand=True)
            self.line_frames.append(frame)
        while len(self.line_frames) > n:
            self.tabview.delete(f"回線#{len(self.line_frames)}")
            self.line_frames.pop()

    def append_log(self, message, level="info"):
        ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        tag = level if level in self.log_tags else "info"
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{ts} - {message}\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.append_log("ログをクリアしました", level="cleared")
        self.log_text.config(state="disabled")

    def set_log_marker(self, message, color="#FFEB3B"):
        self.log_marker_label.configure(text_color=color)
        self.log_marker_var.set(message)

    def clear_log_marker(self):
        self.log_marker_var.set("")

    def update_progress(self, percent):
        self.progress_var.set(percent)
        self.progress_bar.update_idletasks()

    def on_exit(self):
        self.destroy()

class LineTabFrame(ctk.CTkFrame):
    def __init__(self, parent, number, mainapp):
        super().__init__(parent, fg_color=mainapp.card_bg, border_width=0, corner_radius=12)
        self.mainapp = mainapp
        self.number = number
        self.success_hub_urls = []
        self.success_ap_urls = []

        self.WIDTH_IP_ENTRY = 180
        self.WIDTH_NUM_ENTRY = 60
        self.WIDTH_BTN_BATCH = 145
        self.WIDTH_BTN_CLEAR = 100
        self.WIDTH_BTN_EXEC = 80
        self.WIDTH_BTN_COPY = 140
        self.PAD_NORMAL = 5
        self.PAD_LARGE = 8
        self.PAD_XLARGE = 18
        
        self.hub_status_var = tk.StringVar(value="")
        self.ap_status_var = tk.StringVar(value="")

        iprow = ctk.CTkFrame(self, fg_color="transparent")
        iprow.pack(anchor="w", pady=(6, 2), fill="x", padx=24)
        ctk.CTkLabel(iprow, text="IP/ホスト名", width=105).pack(side="left", padx=(0, self.PAD_NORMAL))
        self.ip_entry = ctk.CTkEntry(iprow, placeholder_text="192.168.1.100 or xxxx.test.jp", width=self.WIDTH_IP_ENTRY)
        self.ip_entry.pack(side="left")
        self.batch_btn = ctk.CTkButton(iprow, text="一括実行", fg_color="#4FC3F7", hover_color="#0091EA",
                                       width=self.WIDTH_BTN_BATCH, font=ctk.CTkFont(weight="bold", size=13),
                                       command=self.on_batch_execute)
        self.batch_btn.pack(side="left", padx=self.PAD_XLARGE)

        self.clear_inputs_btn = ctk.CTkButton(iprow, text="入力クリア",
                                              command=self.clear_inputs,
                                              fg_color="#616161", hover_color="#757575",
                                              width=self.WIDTH_BTN_CLEAR)
        self.clear_inputs_btn.pack(side="left", padx=self.PAD_NORMAL)

        hub_group = ctk.CTkFrame(self, fg_color="#2B3A45", border_width=0, corner_radius=18)
        hub_group.pack(anchor="center", fill="x", padx=24, pady=(10, 5))

        ctk.CTkLabel(
            hub_group, text="🗂 HUB設定",
            font=ctk.CTkFont(size=12, weight="bold"), text_color="#4FC3F7"
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=6, pady=(8, 4))

        ctk.CTkLabel(hub_group, text="台数").grid(row=1, column=0, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.hub_count_entry = ctk.CTkEntry(hub_group, width=self.WIDTH_NUM_ENTRY)
        self.hub_count_entry.grid(row=1, column=1, padx=self.PAD_LARGE, pady=4)
        self.hub_count_entry.insert(0, "1")

        ctk.CTkLabel(hub_group, text="開始末尾番号").grid(row=1, column=2, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.hub_start_entry = ctk.CTkEntry(hub_group, width=self.WIDTH_NUM_ENTRY)
        self.hub_start_entry.grid(row=1, column=3, padx=self.PAD_LARGE, pady=4)
        self.hub_start_entry.insert(0, "1")
        
        ctk.CTkLabel(hub_group, text="最大待機時間 秒").grid(row=2, column=0, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.hub_timeout_entry = ctk.CTkEntry(hub_group, width=self.WIDTH_NUM_ENTRY)
        self.hub_timeout_entry.grid(row=2, column=1, padx=self.PAD_LARGE, pady=4)
        self.hub_timeout_entry.insert(0, str(DEFAULT_HUB_TIMEOUT_SEC))

        ctk.CTkButton(
            hub_group, text="実行", width=self.WIDTH_BTN_EXEC,
            command=self.on_hub_execute, fg_color="#42A5F5", hover_color="#1976D2",
            text_color="#212121"
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=self.PAD_LARGE, pady=(4, 8))

        ctk.CTkButton(
            hub_group, text="成功URLコピー", width=self.WIDTH_BTN_COPY,
            command=self.copy_success_hub_urls, fg_color="#42A5F5", hover_color="#388E3C",
            text_color="#212121"
        ).grid(row=3, column=2, columnspan=2, sticky="e", padx=self.PAD_LARGE, pady=(4, 8))

        hub_group.columnconfigure((0,1,2,3,4,5), weight=1)

        ap_group = ctk.CTkFrame(self, fg_color="#2B3A45", border_width=0, corner_radius=18)
        ap_group.pack(anchor="center", fill="x", padx=24, pady=(5, 10))

        ctk.CTkLabel(
            ap_group, text="📶 AP設定",
            font=ctk.CTkFont(size=12, weight="bold"), text_color="#AED581"
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=6, pady=(8, 4))
        
        ctk.CTkLabel(ap_group, text="台数").grid(row=1, column=0, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.ap_count_entry = ctk.CTkEntry(ap_group, width=self.WIDTH_NUM_ENTRY)
        self.ap_count_entry.grid(row=1, column=1, padx=self.PAD_LARGE, pady=4)
        self.ap_count_entry.insert(0, "6")

        ctk.CTkLabel(ap_group, text="開始末尾番号").grid(row=1, column=2, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.ap_start_entry = ctk.CTkEntry(ap_group, width=self.WIDTH_NUM_ENTRY)
        self.ap_start_entry.grid(row=1, column=3, padx=self.PAD_LARGE, pady=4)
        self.ap_start_entry.insert(0, "1")
        
        ctk.CTkLabel(ap_group, text="最大待機時間　秒").grid(row=2, column=0, sticky="e", padx=self.PAD_LARGE, pady=4)
        self.ap_timeout_entry = ctk.CTkEntry(ap_group, width=self.WIDTH_NUM_ENTRY)
        self.ap_timeout_entry.grid(row=2, column=1, padx=self.PAD_LARGE, pady=4)
        self.ap_timeout_entry.insert(0, str(DEFAULT_TIMEOUT_SEC))

        ctk.CTkButton(
            ap_group, text="実行", width=self.WIDTH_BTN_EXEC,
            command=self.on_ap_execute, fg_color="#81C784", hover_color="#388E3C",
            text_color="#212121"
        ).grid(row=3, column=0, columnspan=2, sticky="w", padx=self.PAD_LARGE, pady=(4, 8))

        ctk.CTkButton(
            ap_group, text="成功URLコピー", width=self.WIDTH_BTN_COPY,
            command=self.copy_success_ap_urls, fg_color="#66BB6A", hover_color="#388E3C",
            text_color="#212121"
        ).grid(row=3, column=2, columnspan=2, sticky="e", padx=self.PAD_LARGE, pady=(4, 8))

        ap_group.columnconfigure((0,1,2,3,4,5), weight=1)
        ctk.CTkLabel(ap_group, textvariable=self.ap_status_var, text_color="#AED581",
            font=ctk.CTkFont(size=11, weight="bold")
        ).grid(row=4, column=0, columnspan=6, pady=(2,2), sticky="w")

        ctk.CTkLabel(hub_group, textvariable=self.hub_status_var, text_color="#4FC3F7",
                    font=ctk.CTkFont(size=11, weight="bold")
        ).grid(row=4, column=0, columnspan=6, pady=(2, 2), sticky="w")

        self.generator_frame = TemplateGeneratorFrame(self)
        self.generator_frame.pack(side="bottom", fill="x", expand=True, padx=24, pady=(10, 0))

    def clear_inputs(self):
        self.ip_entry.delete(0, "end")
        self.hub_count_entry.delete(0, "end")
        self.hub_start_entry.delete(0, "end")
        self.ap_count_entry.delete(0, "end")
        self.ap_start_entry.delete(0, "end")
        self.hub_timeout_entry.delete(0, "end")
        self.ap_timeout_entry.delete(0, "end")
        
        self.hub_count_entry.insert(0, "1")
        self.hub_start_entry.insert(0, "1")
        self.ap_count_entry.insert(0, "6")
        self.ap_start_entry.insert(0, "1")
        self.hub_timeout_entry.insert(0, str(DEFAULT_HUB_TIMEOUT_SEC))
        self.ap_timeout_entry.insert(0, str(DEFAULT_TIMEOUT_SEC))

        self.mainapp.append_log(f"回線#{self.number}: 入力内容をクリアしました", level="cleared")

    def _exec_threaded(self, func, *args):
        # --- 変更点3：ボタンの状態管理を修正 ---
        self.batch_btn.configure(state="disabled")
        self.clear_inputs_btn.configure(state="disabled")
        self.mainapp.stop_button.configure(state="normal", text="緊急停止")
        self.mainapp.stop_event.clear()

        def run_and_reenable():
            try:
                func(*args)
            finally:
                # 処理が正常終了、エラー、緊急停止のいずれでもUIを元に戻す
                self.batch_btn.configure(state="normal")
                self.clear_inputs_btn.configure(state="normal")
                self.mainapp.stop_button.configure(state="disabled", text="緊急停止")

        t = threading.Thread(target=run_and_reenable, daemon=True)
        t.start()

    def _check_connection(self, url, timeout_sec):
        try:
            res = requests.get(url, timeout=timeout_sec)
            # --- 変更点4：並列処理用に結果をタプルで返す ---
            return True, url
        except requests.exceptions.RequestException:
            return False, url

    def on_batch_execute(self):
        self._exec_threaded(self._batch_execute)

    def on_hub_execute(self):
        self._exec_threaded(self._hub_execute)

    def on_ap_execute(self):
        self._exec_threaded(self._ap_execute)

    def _validate_and_get_inputs(self, check_ap=True, check_hub=True):
        """入力フォームの値を検証し、辞書として返す。失敗時はNoneを返す。"""
        ip = self.ip_entry.get().strip()
        if not ip:
            self.mainapp.append_log(f"回線#{self.number}: IPアドレスまたはホスト名が入力されていません", level="warn")
            return None
        
        inputs = {"ip": ip}
        try:
            if check_hub:
                inputs["hub_count"] = int(self.hub_count_entry.get())
                inputs["hub_start"] = int(self.hub_start_entry.get())
                try:
                    inputs["hub_timeout"] = int(self.hub_timeout_entry.get())
                except (ValueError, TypeError):
                    self.mainapp.append_log(f"回線#{self.number}: HUBタイムアウト値不正。デフォルト値({DEFAULT_HUB_TIMEOUT_SEC}秒)を使用", level="warn")
                    inputs["hub_timeout"] = DEFAULT_HUB_TIMEOUT_SEC

            if check_ap:
                inputs["ap_count"] = int(self.ap_count_entry.get())
                inputs["ap_start"] = int(self.ap_start_entry.get())
                try:
                    inputs["ap_timeout"] = int(self.ap_timeout_entry.get())
                except (ValueError, TypeError):
                    self.mainapp.append_log(f"回線#{self.number}: APタイムアウト値不正。デフォルト値({DEFAULT_TIMEOUT_SEC}秒)を使用", level="warn")
                    inputs["ap_timeout"] = DEFAULT_TIMEOUT_SEC

            return inputs
        except ValueError:
            self.mainapp.append_log(f"回線#{self.number}: 台数または開始番号に数字でない値が入力されています", level="warn")
            return None

    # --- 変更点4：並列処理を行うように大幅に修正 ---
    # --- 変更点：ログの表示順を担保するように修正 ---
    def _perform_connectivity_checks(self, ip, base_port, count, start_num, device_name, success_level, success_list, timeout_sec):
        """指定された機器群への疎通確認をまとめて実行し、結果をソートしてログ出力する"""
        success_list.clear()
        mode = self.mainapp.access_mode.get()
        
        # --- 処理結果を一時的に保存するリスト ---
        results_buffer = []

        urls_to_check = []
        for i in range(count):
            port = base_port + start_num + i
            if not (0 < port < 65536):
                self.mainapp.append_log(f"回線#{self.number}: {device_name}ポート不正: {port}", level="warn")
                continue
            urls_to_check.append(f"http://{ip}:{port}")
        
        if not urls_to_check:
            return 0, 0

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_url = {executor.submit(self._check_connection, url, timeout_sec): url for url in urls_to_check}
            
            total_tasks = len(urls_to_check)
            completed_tasks = 0

            for future in as_completed(future_to_url):
                if self.mainapp.stop_event.is_set():
                    break 

                url = future_to_url[future]
                try:
                    is_success, received_url = future.result()
                    # --- ログ出力せず、結果をバッファに保存 ---
                    results_buffer.append({'url': received_url, 'success': is_success, 'error': None})
                except Exception as exc:
                    # --- エラーもバッファに保存 ---
                    results_buffer.append({'url': url, 'success': False, 'error': exc})
                
                completed_tasks += 1
                self.mainapp.update_progress(completed_tasks / total_tasks)
        
        # --- 処理完了後、URLでソート ---
        results_buffer.sort(key=lambda r: r['url'])

        success_count = 0
        fail_count = 0

        # --- ソートした結果を順番にログ出力 ---
        for result in results_buffer:
            if result['success']:
                success_count += 1
                success_list.append(result['url'])
                self.mainapp.append_log(f"回線#{self.number}: {device_name} 成功: {result['url']}", level=success_level)
                if mode == "browser":
                    open_url_in_chrome_force_tab(result['url'], self.mainapp)
            else:
                fail_count += 1
                if result['error']:
                    self.mainapp.append_log(f"回線#{self.number}: {device_name} エラー: {result['url']} ({result['error']})", level="fail")
                else:
                    self.mainapp.append_log(f"回線#{self.number}: {device_name} 失敗: {result['url']}", level="fail")

        return success_count, fail_count


    def _batch_execute(self):
        inputs = self._validate_and_get_inputs()
        if not inputs:
            return

        self.mainapp.set_log_marker("⚙️ 一括実行中...", "#FFEB3B")
        self.mainapp.update_progress(0.0)

        # --- RT ---
        rt_url = f"http://{inputs['ip']}:{RT_HUB_BASE_PORT}"
        rt_success_count, rt_fail_count = 0, 0
        if self.mainapp.access_mode.get() == "browser":
            open_url_in_chrome_force_tab(rt_url, self.mainapp)
        is_rt_success, _ = self._check_connection(rt_url, timeout_sec=inputs["ap_timeout"])
        if is_rt_success:
            rt_success_count = 1
            self.mainapp.append_log(f"回線#{self.number}: RT 成功: {rt_url}", level="rt_success")
        else:
            rt_fail_count = 1
            self.mainapp.append_log(f"回線#{self.number}: RT 失敗: {rt_url}", level="fail")
        
        if self.mainapp.stop_event.is_set(): return

        # --- HUB ---
        self.hub_status_var.set("HUB実行中...")
        self.mainapp.update_progress(0)
        hub_success_count, hub_fail_count = self._perform_connectivity_checks(
            inputs["ip"], RT_HUB_BASE_PORT, inputs["hub_count"], inputs["hub_start"],
            "HUB", "hub_success", self.success_hub_urls,
            timeout_sec=inputs["hub_timeout"]
        )
        self.hub_status_var.set("HUB完了")
        
        if self.mainapp.stop_event.is_set(): return

        # --- AP ---
        self.ap_status_var.set("AP実行中...")
        self.mainapp.update_progress(0)
        ap_success_count, ap_fail_count = self._perform_connectivity_checks(
            inputs["ip"], AP_BASE_PORT, inputs["ap_count"], inputs["ap_start"],
            "AP", "success", self.success_ap_urls,
            timeout_sec=inputs["ap_timeout"]
        )
        self.ap_status_var.set("AP完了")

        # --- Summary & Cleanup ---
        self.mainapp.update_progress(1.0)
        self.mainapp.append_log(f"回線#{self.number}: RT 成功 {rt_success_count}件 / 失敗 {rt_fail_count}件", level="summary_rt")
        self.mainapp.append_log(f"回線#{self.number}: HUB 成功 {hub_success_count}件 / 失敗 {hub_fail_count}件", level="summary_hub")
        self.mainapp.append_log(f"回線#{self.number}: AP 成功 {ap_success_count}件 / 失敗 {ap_fail_count}件", level="summary_ap")

        self.mainapp.set_log_marker("✅ 完了", "#00E676")
        time.sleep(1.1)
        self.mainapp.clear_log_marker()
        self.mainapp.update_progress(0.0)
        self.hub_status_var.set("")
        self.ap_status_var.set("")

    def _hub_execute(self):
        inputs = self._validate_and_get_inputs(check_ap=False)
        if not inputs:
            return
            
        self.mainapp.set_log_marker("⚙️ HUBアクセス中...", "#FFEB3B")
        self.mainapp.update_progress(0.0)
        self.hub_status_var.set("HUB実行中...")

        # --- RT ---
        rt_url = f"http://{inputs['ip']}:{RT_HUB_BASE_PORT}"
        if self.mainapp.access_mode.get() == "browser":
            open_url_in_chrome_force_tab(rt_url, self.mainapp)
        is_rt_success, _ = self._check_connection(rt_url, timeout_sec=inputs["hub_timeout"])
        if is_rt_success:
            self.mainapp.append_log(f"回線#{self.number}: RT 成功: {rt_url}", level="rt_success")
        else:
            self.mainapp.append_log(f"回線#{self.number}: RT 失敗: {rt_url}", level="fail")
        
        if self.mainapp.stop_event.is_set(): return

        # --- HUB ---
        self._perform_connectivity_checks(
            inputs["ip"], RT_HUB_BASE_PORT, inputs["hub_count"], inputs["hub_start"],
            "HUB", "hub_success", self.success_hub_urls,
            timeout_sec=inputs["hub_timeout"]
        )

        self.mainapp.update_progress(1.0)
        self.hub_status_var.set("HUB完了")
        self.mainapp.set_log_marker("✅ HUB完了", "#00E676")
        time.sleep(1)
        self.mainapp.clear_log_marker()
        self.mainapp.update_progress(0.0)
        self.hub_status_var.set("")

    def _ap_execute(self):
        inputs = self._validate_and_get_inputs(check_hub=False)
        if not inputs:
            return

        self.mainapp.set_log_marker("⚙️ APアクセス中...", "#FFEB3B")
        self.mainapp.update_progress(0.0)
        self.ap_status_var.set("AP実行中...")

        self._perform_connectivity_checks(
            inputs["ip"], AP_BASE_PORT, inputs["ap_count"], inputs["ap_start"],
            "AP", "success", self.success_ap_urls,
            timeout_sec=inputs["ap_timeout"]
        )
        
        self.mainapp.update_progress(1.0)
        self.ap_status_var.set("AP完了")
        self.mainapp.set_log_marker("✅ AP完了", "#00E676")
        time.sleep(1)
        self.mainapp.clear_log_marker()
        self.mainapp.update_progress(0.0)
        self.ap_status_var.set("")

    def copy_success_hub_urls(self):
        if not self.success_hub_urls:
            self.mainapp.append_log(f"回線#{self.number}: コピー対象URLなし", level="warn")
            return
        pyperclip.copy("\n".join(self.success_hub_urls))
        self.mainapp.append_log(f"回線#{self.number}: 成功URLコピー完了（{len(self.success_hub_urls)}件）", level="cleared")

    def copy_success_ap_urls(self):
        if not self.success_ap_urls:
            self.mainapp.append_log(f"回線#{self.number}: コピー対象URLなし", level="warn")
            return
        pyperclip.copy("\n".join(self.success_ap_urls))
        self.mainapp.append_log(f"回線#{self.number}: 成功URLコピー完了（{len(self.success_ap_urls)}件）", level="cleared")


class TemplateGeneratorFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="#263238", border_width=0, corner_radius=14)
        self.mainapp = parent.mainapp

        self.TEMPLATES = {
            "回線 既存": "{date}test部{staff_name}：回線工事完了\n{company} {worker}({phone})\nRT・HUB疎通OK",
            "回線 新築": "{date}test部{staff_name}：回線工事完了\nONU試験済み 電源{power}\n{company} {worker}({phone})",
            "PT 新築": "{date}test部{staff_name}：共用部工事完了\n{power}電源 {company} {worker}({phone})\nRT・HUB・AP({ip_list})疎通OK",
            "回線 SIM": "{date}test部{staff_name}：NTT工事完了連絡あり\nSIMのためRTのみ疎通確認OK\n{company} {worker}({phone})",
            "PT SIM": "{date}test部{staff_name}：\n仮SIM回収、本開通({company})\nRT・HUB疎通確認OK Wi-Fi WEB疎通確認OK"
        }
        
        self.TEMPLATE_NAMES = list(self.TEMPLATES.keys())
        
        self.VISIBILITY_MAP = {
            "回線 既存": ["name", "company", "worker", "phone"],
            "回線 新築":   ["name", "company", "worker", "phone", "power"],
            "PT 新築":    ["name", "company", "worker", "phone", "ap_count", "power"],
            "回線 SIM":   ["name", "company", "worker", "phone"],
            "PT SIM":     ["name", "company", "worker", "phone"],
        }
        
        self.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self._create_widgets()
        self._on_template_change()

    def _validate_numeric_input(self, P):
        return P.isdigit() or P == ""

    def _create_widgets(self):
        BTN_WIDTH = 80
        BTN_RESET_WIDTH = 160
        PAD = 5

        ctk.CTkLabel(self, text="伝達特記事項 テンプレート選択").grid(row=0, column=0, padx=10, pady=(10, PAD), sticky="w")
        self.template_menu = ctk.CTkOptionMenu(self, values=self.TEMPLATE_NAMES, command=self._on_template_change)
        self.template_menu.grid(row=0, column=1, columnspan=3, padx=10, pady=(10, PAD), sticky="ew")

        validation_cmd = (self.register(self._validate_numeric_input), '%P')

        self.name_label = ctk.CTkLabel(self, text="名前")
        self.name_entry = ctk.CTkEntry(self)
        self.company_label = ctk.CTkLabel(self, text="会社名")
        self.company_entry = ctk.CTkEntry(self)
        self.worker_label = ctk.CTkLabel(self, text="作業員名")
        self.worker_entry = ctk.CTkEntry(self)
        self.phone_label = ctk.CTkLabel(self, text="電話番号")
        self.phone_entry = ctk.CTkEntry(self, validate="key", validatecommand=validation_cmd)
        self.ap_count_label = ctk.CTkLabel(self, text="AP台数")
        self.ap_count_entry = ctk.CTkEntry(self, validate="key", validatecommand=validation_cmd)
        self.power_label = ctk.CTkLabel(self, text="本設/仮設")
        self.power_menu = ctk.CTkOptionMenu(self, values=["本設", "仮設"])
        self.power_menu.set("本設")
        
        self.widgets = {
            "name":     (self.name_label, self.name_entry),
            "company":  (self.company_label, self.company_entry),
            "worker":   (self.worker_label, self.worker_entry),
            "phone":    (self.phone_label, self.phone_entry),
            "ap_count": (self.ap_count_label, self.ap_count_entry),
            "power":    (self.power_label, self.power_menu),
        }

        self.name_label.grid(row=1, column=0, padx=PAD, pady=PAD, sticky="e")
        self.name_entry.grid(row=1, column=1, padx=PAD, pady=PAD, sticky="ew")
        self.company_label.grid(row=1, column=2, padx=PAD, pady=PAD, sticky="e")
        self.company_entry.grid(row=1, column=3, padx=PAD, pady=PAD, sticky="ew")
        self.worker_label.grid(row=2, column=0, padx=PAD, pady=PAD, sticky="e")
        self.worker_entry.grid(row=2, column=1, padx=PAD, pady=PAD, sticky="ew")
        self.phone_label.grid(row=2, column=2, padx=PAD, pady=PAD, sticky="e")
        self.phone_entry.grid(row=2, column=3, padx=PAD, pady=PAD, sticky="ew")
        self.ap_count_label.grid(row=3, column=0, padx=PAD, pady=PAD, sticky="e")
        self.ap_count_entry.grid(row=3, column=1, padx=PAD, pady=PAD, sticky="ew")
        self.power_label.grid(row=3, column=2, padx=PAD, pady=PAD, sticky="e")
        self.power_menu.grid(row=3, column=3, padx=PAD, pady=PAD, sticky="w")

        self.output_textbox = ctk.CTkTextbox(self, height=85, state="disabled")
        self.output_textbox.grid(row=4, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=5, column=0, columnspan=4, pady=(0, 10))

        self.generate_button = ctk.CTkButton(buttons_frame, text="生成", width=BTN_WIDTH, command=self._on_generate)
        self.generate_button.pack(side="left", padx=PAD)
        self.copy_button = ctk.CTkButton(buttons_frame, text="コピー", width=BTN_WIDTH, command=self._on_copy)
        self.copy_button.pack(side="left", padx=PAD)
        self.reset_button = ctk.CTkButton(buttons_frame, text="リセット(名前は保持)", width=BTN_RESET_WIDTH, command=self._on_reset)
        self.reset_button.pack(side="left", padx=PAD)

    def _on_template_change(self, selected_template: str = None):
        if selected_template is None:
            selected_template = self.template_menu.get()
        visible_widgets = self.VISIBILITY_MAP.get(selected_template, [])
        for key, (label, widget) in self.widgets.items():
            if key in visible_widgets:
                label.grid()
                widget.grid()
            else:
                label.grid_remove()
                widget.grid_remove()
                if isinstance(widget, ctk.CTkEntry):
                    widget.delete(0, "end")

    def _on_generate(self):
        try:
            ap_count_val = int(self.ap_count_entry.get())
        except (ValueError, TypeError):
            ap_count_val = 0
        
        ip_list_str = "IP" + ",".join(str(i + 1) for i in range(ap_count_val))

        params = {
            "staff_name": self.name_entry.get(),
            "company": self.company_entry.get(),
            "worker": self.worker_entry.get(),
            "phone": self.phone_entry.get(),
            "ap_count": ap_count_val,
            "power": self.power_menu.get(),
            "date": datetime.now().strftime("%Y/%m/%d"),
            "ip_list": ip_list_str
        }

        template_string = self.TEMPLATES.get(self.template_menu.get())
        generated_text = template_string.format(**params)

        self.output_textbox.configure(state="normal")
        self.output_textbox.delete("1.0", "end")
        self.output_textbox.insert("1.0", generated_text)
        self.output_textbox.configure(state="disabled")
    
    def _on_copy(self):
        text_to_copy = self.output_textbox.get("1.0", "end-1c")
        if not text_to_copy:
            return
        pyperclip.copy(text_to_copy)
        
        original_text = self.copy_button.cget("text")
        self.copy_button.configure(text="コピー完了！", state="disabled", fg_color="#66BB6A")
        self.after(1500, lambda: self.copy_button.configure(text=original_text, state="normal", fg_color=("#3B8ED0", "#1F6AA5")))

    def _on_reset(self):
        self.company_entry.delete(0, "end")
        self.worker_entry.delete(0, "end")
        self.phone_entry.delete(0, "end")
        self.ap_count_entry.delete(0, "end")
        self.power_menu.set("本設")
        
        self.output_textbox.configure(state="normal")
        self.output_textbox.delete("1.0", "end")
        self.output_textbox.configure(state="disabled")
        
        self.mainapp.append_log("文章生成欄をリセットしました（名前除く）", level="cleared")


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
