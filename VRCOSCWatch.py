import importlib
import os
import sys
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from time import sleep
from threading import Thread
from math import ceil
import logging
import json
import ADLXPybind as ADLX  # ADLXPybind.pydをインポート


def import_with_install(package, module_name=None):
    """
    パッケージをインポートし、必要であれば自動でインストールする関数。
    :param package: インストールするパッケージ名（pipで指定する名前）
    :param module_name: インポート時に使用するモジュール名（デフォルトはpackageと同じ）
    """
    module_name = module_name or package
    try:
        globals()[module_name] = importlib.import_module(module_name)
    except ImportError:
        os.system(f"{sys.executable} -m pip install {package}")
        globals()[module_name] = importlib.import_module(module_name)

# 必要なライブラリをインポートまたはインストール
import_with_install("python-osc", "pythonosc")
from pythonosc.udp_client import SimpleUDPClient
import_with_install("nvidia-ml-py3", "pynvml")
import pynvml

class OSCWatchApp:
    # OSCパラメータ
    AVATAR_PARAMS = {
        "HourTenPlace"   : "/avatar/parameters/HourTenPlace",
        "HourZeroPlace"  : "/avatar/parameters/HourZeroPlace",
        "MinuteTenPlace" : "/avatar/parameters/MinuteTenPlace",
        "MinuteZeroPlace": "/avatar/parameters/MinuteZeroPlace",
        "GPUTenPlace"    : "/avatar/parameters/GPUTenPlace",
        "GPUZeroPlace"   : "/avatar/parameters/GPUZeroPlace",
        "VRAMTenPlace"   : "/avatar/parameters/VRAMTenPlace",
        "VRAMZeroPlace"  : "/avatar/parameters/VRAMZeroPlace",
    }
    # カレント取得
    currentDir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))

    def __init__(self, root):
        self.root = root
        self.root.title("VRC-OSC-Watch")
        self.setup_logging()
        
        # チャットプリセットファイルのパスを設定
        script_dir = self.currentDir
        self.CHAT_PRESETS_FILE = os.path.join(script_dir, "chat_presets.json")
        self.SETTINGS_FILE = os.path.join(script_dir, "settings.json")

        self.gpu_vendor = self.detect_gpu_vendor()
        self.load_chat_presets()
        self.load_settings()
        self.create_widgets()
        self.client = None
        self.running = False
        if self.defaultStart_var.get():
            self.start()

    def setup_logging(self):
        """ログファイル"""
        # logフォルダを作成
        script_dir = self.currentDir
        log_dir = os.path.join(script_dir, "log")
        os.makedirs(log_dir, exist_ok=True)
        
        # ログファイル名（日時付き）
        log_filename = f"vrc_osc_watch_{datetime.now().strftime('%Y%m%d')}.log"
        log_filepath = os.path.join(log_dir, log_filename)

        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        if not logger.handlers:
            fmt = logging.Formatter('%(asctime)s - %(message)s')
            file_handler = logging.FileHandler(log_filepath, mode='a', encoding='utf-8')
            file_handler.setFormatter(fmt)
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(fmt)

            logger.addHandler(file_handler)
            logger.addHandler(stream_handler)

        self.logger = logger

    def detect_gpu_vendor(self):
        """GPU検出"""
        # 1. NVIDIA外付け
        try:
            import pynvml
            pynvml.nvmlInit()
            count = pynvml.nvmlDeviceGetCount()
            if count > 0:
                handle = pynvml.nvmlDeviceGetHandleByIndex(0)
                self.gpu_name = pynvml.nvmlDeviceGetName(handle).decode("utf-8")
                pynvml.nvmlShutdown()
                return "NVIDIA"
        except:
            pass

        # 2. AMD外付け
        try:            
            # ADLXHelperで初期化
            self.adlx_helper = ADLX.ADLXHelper()
            self.ret = self.adlx_helper.Initialize()
            if self.ret != ADLX.ADLX_RESULT.ADLX_OK:
                pass
            
            # System Services取得
            system = self.adlx_helper.GetSystemServices()
            if system is None:
                print("Failed to get system services")
                self.adlx_helper.Terminate()
                return

            # Performance Monitoring Services取得
            self.perf_monitoring = system.GetPerformanceMonitoringServices()
            if self.perf_monitoring is None:
                print("Failed to get performance monitoring services")
                self.adlx_helper.Terminate()
                return
            
            # GPUリスト取得
            gpu_list = system.GetGPUs()
            if gpu_list is None:
                print("Failed to get GPU list")
                self.adlx_helper.Terminate()
                return

            # 最初のGPUでメトリクス取得（複数GPUの場合ループ）
            self.gpu = gpu_list[0]  # 最初のGPU
            self.metrics_support = self.perf_monitoring.GetSupportedGPUMetrics(self.gpu)
            if self.ret != ADLX.ADLX_RESULT.ADLX_OK:
                print("Failed to get metrics support")
                self.adlx_helper.Terminate()
                return
            
            self.console("ADLXHelperで初期化成功")
            self.gpu_name = self.gpu.Name()
            return "RADEON"
        except:
            pass

        # 3. 内蔵GPU
        try:
            import wmi
            c = wmi.WMI()
            gpus = c.Win32_VideoController()
            if gpus:
                self.gpu_name = gpus[0].Name
                return "INTEGRATED"
        except:
            return None

    def create_widgets(self):
        # 基本設定エリア
        # IP
        tk.Label(self.root, text="IP Address").grid(row=0, column=0, sticky="w", pady=2)
        self.ip_entry = tk.Entry(self.root)
        self.ip_entry.grid(row=0, column=1, sticky="ew", pady=2)
        self.ip_entry.insert(0, "127.0.0.1")

        # Port
        tk.Label(self.root, text="Port").grid(row=1, column=0, sticky="w", pady=2)
        self.port_entry = tk.Entry(self.root)
        self.port_entry.grid(row=1, column=1, sticky="ew", pady=2)
        self.port_entry.insert(0, "9000")

        # Interval
        tk.Label(self.root, text="送信間隔 (秒)").grid(row=2, column=0, sticky="w", pady=2)
        self.interval_entry = tk.Entry(self.root)
        self.interval_entry.grid(row=2, column=1, sticky="ew", pady=2)
        self.interval_entry.insert(0, "5")

        # Default Start
        tk.Checkbutton(self.root, text="起動時にStart開始", variable=self.defaultStart_var, command=self.save_settings).grid(row=3, column=0, columnspan=2, sticky="w")

        # Start/Stop buttons frame
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=2, sticky="w")
        tk.Button(button_frame, text="Start", command=self.start).pack(side='left', padx=(0,5))
        tk.Button(button_frame, text="Stop", command=self.stop).pack(side='left')

        # 詳細設定の折りたたみボタン
        initial_text = "おまけ機能チャット送信 OFF" if self.chat_enabled_var.get() else "おまけ機能チャット送信 ON"
        self.toggle_button = tk.Button(self.root, text=initial_text, command=self.toggle_advanced_settings)
        self.toggle_button.grid(row=5, column=0, columnspan=2, pady=2, sticky="w")

        # 詳細設定フレーム（最初は非表示）
        self.advanced_frame = tk.Frame(self.root)
        # 保存された設定に基づいて表示/非表示を決定
        if self.chat_enabled_var.get():
            self.advanced_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=2)

        # 詳細設定の中身
        # チャットメッセージ入力エリア
        tk.Label(self.advanced_frame, text="チャットメッセージ:").grid(row=0, column=0, columnspan=2, sticky="w", pady=2)
        
        # スクロール付きテキストボックス
        self.chat_frame = tk.Frame(self.advanced_frame)
        self.chat_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        
        self.chat_text = tk.Text(self.chat_frame, height=4, width=40, wrap=tk.WORD, state=tk.DISABLED)
        self.chat_scrollbar = tk.Scrollbar(self.chat_frame, orient=tk.VERTICAL, command=self.chat_text.yview)
        self.chat_text.config(yscrollcommand=self.chat_scrollbar.set)
        
        # チャットテキストが変更された際にステータス更新をバインド
        self.chat_text.bind('<KeyRelease>', self.on_chat_text_change)
        self.chat_text.bind('<Button-1>', self.on_chat_text_change)
        self.chat_text.bind('<Control-v>', self.on_chat_text_change)
        
        self.chat_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.chat_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # プリセット保存ボタン
        self.save_preset_button = tk.Button(self.advanced_frame, text="現在のメッセージをプリセットに保存", command=self.save_current_chat_as_preset, state=tk.DISABLED)
        self.save_preset_button.grid(row=2, column=0, columnspan=2, pady=2)

        # プリセット表示エリア
        self.preset_label = tk.Label(self.advanced_frame, text="プリセットメッセージ:")
        self.preset_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=2)
        self.presets_frame = tk.Frame(self.advanced_frame)
        self.presets_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=2)
        self.update_preset_buttons()

        # 詳細設定フレームの列の重みを設定
        self.advanced_frame.columnconfigure(1, weight=1)

        # ステータス表示
        self.status_label = tk.Label(self.root, text="ステータス: Stop", fg="red", justify="left")
        self.status_label.grid(row=7, column=0, columnspan=2, sticky="w", pady=0)

        # GPU情報表示
        gpu_info_text = f"GPU: {self.gpu_vendor}\n{self.gpu_name}"
        self.gpu_info_label = tk.Label(self.root, text=gpu_info_text, fg="blue", justify="left")
        self.gpu_info_label.grid(row=8, column=0, columnspan=2, sticky="w", pady=0)

        # グリッドの列の重みを設定（エントリーフィールドが伸縮するように）
        self.root.columnconfigure(1, weight=1)
        
        # 保存された設定に基づいてチャット入力の初期状態を設定
        self.toggle_chat_input()

    def toggle_advanced_settings(self):
        """詳細設定の表示/非表示とチャット機能のON/OFFを切り替える"""
        # 状態を反転
        is_enabled = not self.chat_enabled_var.get()
        self.chat_enabled_var.set(is_enabled)

        if is_enabled:
            # 詳細設定を表示し、機能を有効化
            self.advanced_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=2)
            self.toggle_button.config(text="おまけ機能チャット送信 OFF")
            self.toggle_chat_input()
        else:
            # 詳細設定を隠し、機能を無効化
            self.advanced_frame.grid_remove()
            self.toggle_button.config(text="おまけ機能チャット送信 ON")
            self.toggle_chat_input()
        
        # ステータス表示を更新
        self.update_status_display()
        
    def toggle_chat_input(self):
        """チャット入力欄の有効/無効を切り替える"""
        new_state = tk.NORMAL if self.chat_enabled_var.get() else tk.DISABLED
        
        self.chat_text.config(state=new_state)
        self.save_preset_button.config(state=new_state)
        
        # プリセットボタンのUIを再構築（適切な状態で）
        self.update_preset_buttons()

    def on_chat_text_change(self, event=None):
        """チャットテキスト変更時のイベントハンドラー"""
        if self.running and self.chat_enabled_var.get():
            # 少し遅延させてステータス更新（連続入力時の負荷軽減）
            self.root.after(500, self.update_status_display)

    def start(self):
        try:
            ip = self.ip_entry.get()
            port = int(self.port_entry.get())
            interval = float(self.interval_entry.get())
            sync = interval  # intervalと同じ値を使用
            self.client = SimpleUDPClient(ip, port)
            self.running = True
            self.update_status_display()
            Thread(target=self.send_messages, args=(interval, sync), daemon=True).start()
        except Exception as e:
            error_msg = f"Start error: {str(e)}"
            self.console(error_msg)
            messagebox.showerror("Error", str(e))

    def stop(self):
        self.running = False
        self.update_status_display()

    def update_status_display(self):
        """ステータス表示を更新する"""
        if self.running:
            status_text = "ステータス: Start"
            color = "green"
            
            # 送信情報を作成
            send_info = ["時計"]
            if self.chat_enabled_var.get():
                # チャットメッセージが空でない場合のみ追加
                message = self.chat_text.get("1.0", tk.END).strip()
                if message:
                    send_info.append("チャット")
            send_info_text = ", ".join(send_info)
            
            full_text = f"{status_text}\n送信情報: {send_info_text}"
        else:
            full_text = "ステータス: Stop"
            color = "red"
        
        self.status_label.config(text=full_text, fg=color)

    def get_gpu_usage_v2(self):
        if self.gpu_vendor == "NVIDIA":
            return self.get_nvidia_gpu_usage()
        elif self.gpu_vendor == "RADEON":
            return self.get_amd_gpu_usage()
        elif self.gpu_vendor == "INTEGRATED":
            self.console("内蔵GPUを検出")
        else:
            self.console("警告: 対応していないGPUです。")
        self.console("GPU使用率は0%として表示されます。")
        self.console("gpu_vendor:" + self.gpu_vendor)        
        return 0, 0

    def get_nvidia_gpu_usage(self):
        """NVIDIA GPU使用率を取得"""
        try:
            pynvml.nvmlInit()
            handle = pynvml.nvmlDeviceGetHandleByIndex(0)
            utilization = pynvml.nvmlDeviceGetUtilizationRates(handle)
            memory_info = pynvml.nvmlDeviceGetMemoryInfo(handle)
            pynvml.nvmlShutdown()
            return int(utilization.gpu), int(memory_info.used / memory_info.total * 100)
        except pynvml.NVMLError_LibraryNotFound:
            self.show_copyable_command_dialog(
                "グラボエラー",
                
                "NVIDIAドライバ (nvml.dll) が見つかりませんでした。\n"
                "NVIDIA公式情報などをご確認ください。\n"
                "\n"
                "2025/7 時点の情報: C:\\Program Files\\NVIDIA Corporation\\NVSMI\\ 配下に nvml.dll が無い場合は、\n"
                "次の手順で手動コピーすると解決する場合があります (自己責任)。\n"
                "1) エクスプローラーで C:\\Windows\\System32\\ を開く\n"
                "2) nvml.dll をコピー\n"
                "3) C:\\Program Files\\NVIDIA Corporation\\NVSMI\\ を開く\n"
                "4) ファイルを貼り付け\n"
                "このアプリを再起動してこのエラーがでないか確認してください。",
            )
            sys.exit(1)
        except Exception as e:
            error_code = e.args[0] if e.args else "Unknown"
            error_message = pynvml.nvmlErrorString(error_code) if hasattr(pynvml, "nvmlErrorString") else str(e)
            self.console(f"Error Code: {error_code}")
            self.console(f"Error Message: {error_message}")
            messagebox.showerror("グラボエラー", f"想定されていないエラーです。\n製作者にお問い合わせお願いします。\nError Code: {error_code}\nError Message: {error_message}")
            sys.exit(1)

    def get_amd_gpu_usage(self):
        # GPUUsageとVRAMUsageがサポートされているか確認
        if self.metrics_support.IsSupportedGPUUsage() and self.metrics_support.IsSupportedGPUVRAM():
            current_metrics = self.perf_monitoring.GetCurrentGPUMetrics(self.gpu)
            if self.ret == ADLX.ADLX_RESULT.ADLX_OK and current_metrics is not None:
                gpu_usage = current_metrics.GPUUsage()  # GPU利用率 (%)
                vram_usage = current_metrics.GPUVRAM()  # VRAM使用量 (MB)
                self.console(f"GpuName: {self.gpu.Name()} ")

                self.console(f"GPU Utilization: {gpu_usage}%")
                self.console(f"VRAM Usage: {vram_usage} MB")
                # Total VRAM取得
                total_vram = self.gpu.TotalVRAM()  # IADLXGPUのVRAMメソッドで総VRAM取得 (MB)
                self.console(f"Total VRAM: {total_vram} MB")
        return int(gpu_usage), int(vram_usage / total_vram * 100)
        
    def send_messages(self, interval, sync):
        sync_count = int(sync / interval)
        counters = {key: 0 for key in self.AVATAR_PARAMS.keys()}
        self.console(f"GPU Vendor detected: {self.gpu_vendor}")
        
        while self.running:
            now = datetime.now()
            self.send_param("HourTenPlace", now.hour // 10, counters, sync_count)
            self.send_param("HourZeroPlace", now.hour % 10, counters, sync_count)
            self.send_param("MinuteTenPlace", now.minute // 10, counters, sync_count)
            self.send_param("MinuteZeroPlace", now.minute % 10, counters, sync_count)
            self.console(f"Sent: {now.strftime('%Y-%m-%d %H:%M:%S')}")
            gpu, vram = self.get_gpu_usage_v2()
            gpu = min(gpu, 99)
            vram = min(vram, 99)
            self.console(f"gpu:{gpu}% vram:{vram}% (vendor:{self.gpu_vendor})")
            self.send_param("GPUTenPlace", gpu // 10, counters, sync_count)
            self.send_param("GPUZeroPlace", gpu % 10, counters, sync_count)
            self.send_param("VRAMTenPlace", vram // 10, counters, sync_count)
            self.send_param("VRAMZeroPlace", vram % 10, counters, sync_count)
            
            # チャット送信処理
            if self.chat_enabled_var.get():
                self.send_chat_message()
            
            sleep(interval)

    def send_param(self, param_name, value, counters, sync_count):
        if counters[param_name] <= 0 or value != counters.get(f"prev_{param_name}", None):
            self.client.send_message(self.AVATAR_PARAMS[param_name], value)
            counters[param_name] = sync_count
            counters[f"prev_{param_name}"] = value
            self.console(f"Param: {param_name}, Address:{self.AVATAR_PARAMS[param_name]} Value: {value}")
        else:
            counters[param_name] -= 1

    def send_chat_message(self):
        """チャットメッセージをVRChatに送信する"""
        try:
            # テキストウィジェットからメッセージを取得
            message = self.chat_text.get("1.0", tk.END).strip()
            if message:
                # VRChatのチャットボックスにメッセージを送信
                self.client.send_message("/chatbox/input", [message, True, False])
                self.console(f"Chat sent: {message}")
            else:
                self.console("Chat message is empty, skipping send")
        except Exception as e:
            error_msg = f"Chat send error: {str(e)}"
            self.console(error_msg)

    def load_chat_presets(self):
        """チャットプリセットをJSONファイルから読み込む"""
        try:
            if os.path.exists(self.CHAT_PRESETS_FILE):
                with open(self.CHAT_PRESETS_FILE, 'r', encoding='utf-8') as f:
                    self.chat_presets = json.load(f)
            else:
                # 空のプリセットファイルを作成
                self.chat_presets = []
                self.save_chat_presets()
        except Exception as e:
            self.console(f"Error loading chat presets: {e}")
            self.chat_presets = []

    def save_chat_presets(self):
        """チャットプリセットをJSONファイルに保存する"""
        try:
            with open(self.CHAT_PRESETS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.chat_presets, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.console(f"Error saving chat presets: {e}")

    def load_settings(self):
        """設定をJSONファイルから読み込む"""
        try:
            if os.path.exists(self.SETTINGS_FILE):
                with open(self.SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # デフォルトStart設定を読み込み
                    self.defaultStart_var = tk.BooleanVar(value=settings.get('defaultStart', True))
            else:
                # デフォルト設定
                self.defaultStart_var = tk.BooleanVar(value=True)
                self.save_settings()
            # チャット機能は初期状態で非活性
            self.chat_enabled_var = tk.BooleanVar(value=False)
        except Exception as e:
            self.console(f"Error loading settings: {e}")
            self.defaultStart_var = tk.BooleanVar(value=True)
            self.chat_enabled_var = tk.BooleanVar(value=False)

    def save_settings(self):
        """設定をJSONファイルに保存する"""
        try:
            settings = {
                'defaultStart': self.defaultStart_var.get()
            }
            with open(self.SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.console(f"Error saving settings: {e}")

    def update_preset_buttons(self):
        """プリセットボタンのUIを更新する"""
        # 既存のウィジェットをクリア
        for widget in self.presets_frame.winfo_children():
            widget.destroy()

        button_state = tk.NORMAL if self.chat_enabled_var.get() else tk.DISABLED

        # プリセットボタンを再作成
        for i, preset in enumerate(self.chat_presets):
            preset_frame = tk.Frame(self.presets_frame)
            preset_frame.pack(fill='x', pady=1)
            
            # 順序変更ボタン用のフレーム
            order_frame = tk.Frame(preset_frame)
            order_frame.pack(side='left')
            
            # 上へボタン（最初の要素以外で有効）
            up_btn_state = button_state if (i > 0 and self.chat_enabled_var.get()) else tk.DISABLED
            up_btn = tk.Button(order_frame, text="↑", width=2, 
                             command=lambda idx=i: self.move_preset_up(idx), 
                             state=up_btn_state)
            up_btn.pack(side='left')
            
            # 下へボタン（最後の要素以外で有効）
            down_btn_state = button_state if (i < len(self.chat_presets) - 1 and self.chat_enabled_var.get()) else tk.DISABLED
            down_btn = tk.Button(order_frame, text="↓", width=2, 
                               command=lambda idx=i: self.move_preset_down(idx), 
                               state=down_btn_state)
            down_btn.pack(side='left')
            
            # プリセットメッセージボタン
            btn = tk.Button(preset_frame, text=preset, command=lambda p=preset: self.add_preset_to_chat(p), state=button_state, anchor='w')
            btn.pack(side='left', expand=True, fill='x', padx=(1,0))
            
            # 削除ボタン
            del_btn = tk.Button(preset_frame, text="削除", command=lambda p=preset: self.delete_preset(p), state=button_state)
            del_btn.pack(side='right')

    def add_preset_to_chat(self, preset_text):
        """プリセットテキストをチャット入力欄に設定する（既存テキストを置き換え）"""
        if self.chat_enabled_var.get():
            # 既存のテキストをクリアしてから新しいテキストを設定
            self.chat_text.delete("1.0", tk.END)
            self.chat_text.insert("1.0", preset_text)
            # ステータス表示を更新
            if self.running:
                self.update_status_display()

    def delete_preset(self, preset_text):
        """プリセットを削除してUIとファイルを更新する"""
        if preset_text in self.chat_presets:
            self.chat_presets.remove(preset_text)
            self.save_chat_presets()
            self.update_preset_buttons()

    def move_preset_up(self, index):
        """プリセットを一つ上に移動する"""
        if index > 0 and index < len(self.chat_presets):
            # 要素を入れ替え
            self.chat_presets[index], self.chat_presets[index - 1] = self.chat_presets[index - 1], self.chat_presets[index]
            self.save_chat_presets()
            self.update_preset_buttons()

    def move_preset_down(self, index):
        """プリセットを一つ下に移動する"""
        if index >= 0 and index < len(self.chat_presets) - 1:
            # 要素を入れ替え
            self.chat_presets[index], self.chat_presets[index + 1] = self.chat_presets[index + 1], self.chat_presets[index]
            self.save_chat_presets()
            self.update_preset_buttons()

    def save_current_chat_as_preset(self):
        """現在のチャットメッセージをプリセットとして保存する"""
        message = self.chat_text.get("1.0", tk.END).strip()
        if message and message not in self.chat_presets:
            self.chat_presets.append(message)
            self.save_chat_presets()
            self.update_preset_buttons()
            self.chat_text.delete("1.0", tk.END) # 保存後に入力欄をクリア
            # ステータス表示を更新
            if self.running:
                self.update_status_display()
            messagebox.showinfo("成功", "プリセットを保存しました。")
        elif not message:
            messagebox.showwarning("警告", "保存するメッセージが空です。")
        else:
            messagebox.showwarning("警告", "このメッセージは既にプリセットに存在します。")

    def console(self, value):
        self.logger.info(value)

    @staticmethod
    def ceil_minifloat(value):
        return ceil(value * 128) / 128

    def show_copyable_command_dialog(self, title, message):
        win = tk.Toplevel(self.root)
        win.title(title)
        tk.Label(win, text=message, justify="left").pack(padx=10, pady=(10,0))
        tk.Button(win, text="閉じる", command=lambda: (win.destroy(), sys.exit(1))).pack(pady=10)
        win.grab_set()
        win.wait_window()

if __name__ == "__main__":
    root = tk.Tk()
    app = OSCWatchApp(root)
    root.mainloop()
