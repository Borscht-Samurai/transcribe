import os
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from pathlib import Path
# from dotenv import load_dotenv, set_key  # dotenv関連のインポートをコメントアウト
import threading
import json

# 既存のtranscribe.pyから関数をインポート
from transcribe import load_audio_file, transcribe_audio

# 環境変数をロード
# load_dotenv() # コメントアウト

# 設定ファイルのパス
CONFIG_FILE = "config.json"

def get_resource_path(relative_path):
    """リソースファイルの絶対パスを取得する"""
    try:
        # PyInstallerでビルドされた場合のパスを取得
        base_path = sys._MEIPASS
    except Exception:
        # 通常のPython実行時は、現在のディレクトリを返す
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_working_directory():
    """作業ディレクトリを取得する"""
    if getattr(sys, 'frozen', False):
        # exeファイル実行時は、exeファイルのディレクトリを返す
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        # 通常のPython実行時は、現在のディレクトリを返す
        return os.path.abspath(os.getcwd())

class TranscribeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("音声文字起こしツール")
        self.root.geometry("800x900")  # 初期サイズを調整
        self.root.minsize(800, 700)  # 最小サイズも調整
        
        # フォントとスタイルの設定
        self.font_default = ("Yu Gothic UI", 10)
        self.font_heading = ("Yu Gothic UI", 12, "bold")
        
        # 設定の読み込み
        self.config = self.load_config()
        
        # メインフレーム
        self.main_frame = tk.Frame(root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        self.title_label = tk.Label(
            self.main_frame, 
            text="音声文字起こし・議事録作成ツール", 
            font=("Yu Gothic UI", 16, "bold"),
            pady=10
        )
        self.title_label.pack(fill=tk.X)
        
        # ファイル選択セクション
        self.file_frame = tk.LabelFrame(
            self.main_frame, 
            text="音声ファイルの選択", 
            font=self.font_heading,
            padx=10, 
            pady=10
        )
        self.file_frame.pack(fill=tk.X, pady=10)
        
        # ファイルパス表示
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(
            self.file_frame, 
            textvariable=self.path_var, 
            font=self.font_default,
            width=50
        )
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # 参照ボタン
        self.browse_button = tk.Button(
            self.file_frame, 
            text="参照...", 
            font=self.font_default,
            command=self.browse_file,
            width=10
        )
        self.browse_button.pack(side=tk.RIGHT)
        
        # APIキー設定セクション
        self.api_frame = tk.LabelFrame(
            self.main_frame, 
            text="API設定", 
            font=self.font_heading,
            padx=10, 
            pady=10
        )
        self.api_frame.pack(fill=tk.X, pady=10)
        
        # APIキー入力
        self.api_label = tk.Label(
            self.api_frame, 
            text="Google API キー:", 
            font=self.font_default
        )
        self.api_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.api_var = tk.StringVar(value=self.config.get("api_key", os.getenv("GOOGLE_API_KEY", "")))
        self.api_var = tk.StringVar(value=self.config.get("api_key", ""))
        self.api_entry = tk.Entry(
            self.api_frame, 
            textvariable=self.api_var, 
            font=self.font_default,
            width=40,
            show="*"  # パスワードのように表示
        )
        self.api_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # APIキー表示切り替えチェックボックス
        self.show_api_var = tk.BooleanVar(value=False)
        self.show_api_check = tk.Checkbutton(
            self.api_frame,
            text="APIキーを表示",
            variable=self.show_api_var,
            font=self.font_default,
            command=self.toggle_api_visibility
        )
        self.show_api_check.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        # APIキー保存ボタン
        self.save_api_button = tk.Button(
            self.api_frame,
            text="APIキーを保存",
            font=self.font_default,
            command=self.save_api_key,
            width=15
        )
        self.save_api_button.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # オプションフレーム
        self.options_frame = tk.LabelFrame(
            self.main_frame, 
            text="オプション", 
            font=self.font_heading,
            padx=10, 
            pady=10
        )
        self.options_frame.pack(fill=tk.X, pady=10)
        
        # モデル選択
        self.model_label = tk.Label(
            self.options_frame, 
            text="モデル:", 
            font=self.font_default
        )
        self.model_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.model_var = tk.StringVar(value="gemini-2.0-flash")
        self.model_combo = ttk.Combobox(
            self.options_frame, 
            textvariable=self.model_var, 
            font=self.font_default,
            values=["gemini-2.0-flash"],
            state="readonly",
            width=20
        )
        self.model_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 言語選択の追加
        self.language_label = tk.Label(
            self.options_frame, 
            text="言語:", 
            font=self.font_default
        )
        self.language_label.grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
        
        self.language_var = tk.StringVar(value="japanese")
        self.language_combo = ttk.Combobox(
            self.options_frame, 
            textvariable=self.language_var, 
            font=self.font_default,
            values=["japanese", "english"],
            state="readonly",
            width=10
        )
        self.language_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # タイムスタンプオプション
        self.timestamp_var = tk.BooleanVar(value=True)
        self.timestamp_check = tk.Checkbutton(
            self.options_frame, 
            text="タイムスタンプを付ける", 
            variable=self.timestamp_var,
            font=self.font_default
        )
        self.timestamp_check.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # 自動保存オプション
        self.autosave_var = tk.BooleanVar(value=False)
        self.autosave_check = tk.Checkbutton(
            self.options_frame, 
            text="結果を自動的にファイルに保存する", 
            variable=self.autosave_var,
            font=self.font_default
        )
        self.autosave_check.grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # 議事録生成オプション
        self.minutes_var = tk.BooleanVar(value=True)
        self.minutes_check = tk.Checkbutton(
            self.options_frame, 
            text="議事録も生成する", 
            variable=self.minutes_var,
            font=self.font_default
        )
        self.minutes_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # ステータス表示
        self.status_var = tk.StringVar(value="準備完了")
        self.status_label = tk.Label(
            self.main_frame,
            textvariable=self.status_var,
            font=self.font_default,
            fg="blue",
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, pady=(0, 5))
        
        # 実行ボタン
        self.execute_button = tk.Button(
            self.main_frame, 
            text="文字起こしを実行", 
            font=self.font_heading,
            bg="#4CAF50", 
            fg="white",
            command=self.execute_transcription,
            height=2
        )
        self.execute_button.pack(fill=tk.X, pady=10)
        
        # 進捗バー
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.main_frame, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='indeterminate', 
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # 結果表示ペイン (Panedウィンドウを使用して分割)
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))
        
        # 文字起こし結果保存ボタン
        self.save_button = tk.Button(
            self.button_frame,
            text="文字起こし結果を保存",
            font=self.font_default,
            command=self.save_result,
            state=tk.DISABLED,
            width=20
        )
        self.save_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 議事録保存ボタン
        self.save_minutes_button = tk.Button(
            self.button_frame,
            text="議事録を保存",
            font=self.font_default,
            command=self.save_minutes,
            state=tk.DISABLED,
            width=20
        )
        self.save_minutes_button.pack(side=tk.LEFT)
        
        # PanedWindowの高さを制限して、ボタンが見えるようにする
        results_frame = tk.Frame(self.main_frame)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 5))
        
        self.results_paned = ttk.PanedWindow(results_frame, orient=tk.VERTICAL)
        self.results_paned.pack(fill=tk.BOTH, expand=True)
        
        # 文字起こし結果フレーム
        self.transcription_frame = tk.LabelFrame(
            self.results_paned,
            text="文字起こし結果",
            font=self.font_heading,
            padx=10,
            pady=10
        )
        
        # 議事録フレーム
        self.minutes_frame = tk.LabelFrame(
            self.results_paned,
            text="議事録",
            font=self.font_heading,
            padx=10,
            pady=10
        )
        
        # ペインに追加
        self.results_paned.add(self.transcription_frame, weight=1)
        self.results_paned.add(self.minutes_frame, weight=1)
        
        # 文字起こし結果テキストエリア
        self.result_text = scrolledtext.ScrolledText(
            self.transcription_frame,
            wrap=tk.WORD,
            font=self.font_default,
            height=20  # 高さを少し小さくする
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 議事録テキストエリア
        self.minutes_text = scrolledtext.ScrolledText(
            self.minutes_frame,
            wrap=tk.WORD,
            font=self.font_default,
            height=20  # 高さを少し小さくする
        )
        self.minutes_text.pack(fill=tk.BOTH, expand=True)
        
        # 初期状態ではプログレスバーを非表示に
        self.progress_bar.pack_forget()
        
        # 現在の処理状態
        self.processing = False
        
        # 現在の文字起こし結果と議事録
        self.current_result = ""
        self.current_minutes = ""
        
        # APIキーのチェック
        if not os.getenv("GOOGLE_API_KEY"):
            messagebox.showerror(
                "APIキーエラー", 
                "Google API Keyが設定されていません。\n.envファイルを確認してください。"
            )
        # config.jsonからAPIキーが読み込まれていない場合のエラーメッセージ
        if not self.config.get("api_key"):
            messagebox.showerror(
                "APIキーエラー",
                "Google API Keyが設定されていません。\nconfig.jsonファイルを確認してください。"
            )

    def browse_file(self):
        """ファイル選択ダイアログを表示して音声ファイルを選択する"""
        filetypes = [
            ("すべての音声ファイル", "*.wav *.flac *.mp3 *.ogg *.webm *.mp4 *.amr *.3gp *.m4a *.opus *.speex"),
            ("WAVファイル", "*.wav"),
            ("FLACファイル", "*.flac"),
            ("MP3ファイル", "*.mp3"),
            ("OGGファイル", "*.ogg"),
            ("すべてのファイル", "*.*")
        ]
        
        try:
            # initialdirをNoneにしておくことで、OSが前回参照したディレクトリを記憶したり、
            # ユーザーの自由な場所からファイルを選択できるようになる
            print(f"\n=== ファイル選択ダイアログ ===")
            filepath = filedialog.askopenfilename(
                title="音声ファイルを選択",
                filetypes=filetypes,
                initialdir=None  # 修正: get_working_directory() を使用せず None に設定
            )
            
            if filepath:
                # 絶対パスに変換して保存
                abs_path = os.path.abspath(filepath)
                print(f"\n2. 選択されたファイル情報:")
                print(f"- 選択されたパス: {filepath}")
                print(f"- 絶対パス: {abs_path}")
                print(f"- ファイルの存在: {os.path.exists(abs_path)}")
                print(f"- ファイルサイズ: {os.path.getsize(abs_path) if os.path.exists(abs_path) else 'N/A'} bytes")
                print(f"- ファイルの親ディレクトリ: {os.path.dirname(abs_path)}")
                print(f"- 親ディレクトリの存在: {os.path.exists(os.path.dirname(abs_path))}")
                
                if not os.path.exists(abs_path):
                    error_msg = f"ファイルが見つかりません: {abs_path}\n"
                    error_msg += f"現在の作業ディレクトリ: {os.getcwd()}\n"
                    error_msg += f"ファイルの親ディレクトリ: {os.path.dirname(abs_path)}\n"
                    error_msg += f"親ディレクトリの存在: {os.path.exists(os.path.dirname(abs_path))}\n"
                    error_msg += f"親ディレクトリの内容: {os.listdir(os.path.dirname(abs_path)) if os.path.exists(os.path.dirname(abs_path)) else 'N/A'}"
                    print(f"\n3. エラー情報:")
                    print(error_msg)
                    messagebox.showerror("エラー", error_msg)
                    return
                
                self.path_var.set(abs_path)
                print(f"\n3. ファイルパス設定完了")
                
        except Exception as e:
            error_msg = f"ファイル選択中にエラーが発生しました:\n"
            error_msg += f"エラーの種類: {type(e).__name__}\n"
            error_msg += f"エラーメッセージ: {str(e)}\n"
            error_msg += f"現在の作業ディレクトリ: {os.getcwd()}"
            print(f"\n3. エラー情報:")
            print(error_msg)
            messagebox.showerror("エラー", error_msg)
    
    def execute_transcription(self):
        """文字起こし処理を実行する"""
        filepath = self.path_var.get().strip()
        
        if not filepath:
            messagebox.showerror("エラー", "音声ファイルを選択してください。")
            return
        
        # ファイルパスを絶対パスに変換
        filepath = os.path.abspath(filepath)
        
        if not os.path.exists(filepath):
            messagebox.showerror("エラー", f"ファイルが見つかりません: {filepath}")
            return
        
        if self.processing:
            return
        
        # APIキーが入力されているか確認
        api_key = self.api_var.get().strip()
        if not api_key:
            messagebox.showerror("エラー", "APIキーが設定されていません。")
            return
        
        # 環境変数にAPIキーを設定
        os.environ["GOOGLE_API_KEY"] = api_key
        
        # UIを処理中状態に更新
        self.processing = True
        self.execute_button.config(state=tk.DISABLED, text="処理中...")
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        self.progress_bar.start(10)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "文字起こし処理中です。しばらくお待ちください...")
        self.minutes_text.delete(1.0, tk.END)
        if self.minutes_var.get():
            self.minutes_text.insert(tk.END, "議事録を生成中です。しばらくお待ちください...")
        self.save_button.config(state=tk.DISABLED)
        self.save_minutes_button.config(state=tk.DISABLED)
        
        # バックグラウンドスレッドで文字起こし処理を実行
        thread = threading.Thread(target=self.process_transcription, args=(filepath,))
        thread.daemon = True
        thread.start()
    
    def process_transcription(self, filepath):
        """バックグラウンドで文字起こし処理を実行する"""
        try:
            print("\n=== 文字起こし処理開始 ===")
            print(f"1. 初期情報:")
            print(f"- 入力ファイルパス: {filepath}")
            print(f"- 現在の作業ディレクトリ: {os.getcwd()}")
            
            # 作業ディレクトリを設定 (削除対象)
            # working_dir = get_working_directory()
            # os.chdir(working_dir)  # ← この行を削除
            print(f"\n2. 作業ディレクトリ設定(不要なので削除):")
            # print(f"- 設定された作業ディレクトリ: {working_dir}")
            # print(f"- 現在の作業ディレクトリ: {os.getcwd()}")
            
            # 絶対パスに変換
            filepath = os.path.abspath(filepath)
            print(f"\n3. ファイルパス情報:")
            print(f"- 絶対パス: {filepath}")
            print(f"- ファイルの存在: {os.path.exists(filepath)}")
            print(f"- ファイルサイズ: {os.path.getsize(filepath) if os.path.exists(filepath) else 'N/A'} bytes")
            print(f"- ファイルの親ディレクトリ: {os.path.dirname(filepath)}")
            print(f"- 親ディレクトリの存在: {os.path.exists(os.path.dirname(filepath))}")
            print(f"- 親ディレクトリの内容: {os.listdir(os.path.dirname(filepath)) if os.path.exists(os.path.dirname(filepath)) else 'N/A'}")
            
            # APIキーの確認
            api_key = self.api_var.get().strip()
            if not api_key:
                error_msg = "APIキーが設定されていません"
                print(f"\n4. APIキーエラー:")
                print(f"- {error_msg}")
                self.update_status("エラー: APIキーが設定されていません")
                self.update_result("エラー: APIキーが設定されていません。API設定セクションでAPIキーを入力してください。", is_error=True)
                self.finish_processing()
                return
            
            # 環境変数にAPIキーを設定
            os.environ["GOOGLE_API_KEY"] = api_key
            print(f"\n4. APIキー設定完了")
            
            model = self.model_var.get()
            language = self.language_var.get()
            with_timestamps = self.timestamp_var.get()
            generate_minutes = self.minutes_var.get()
            
            print(f"\n5. 処理設定:")
            print(f"- モデル: {model}")
            print(f"- 言語: {language}")
            print(f"- タイムスタンプ: {'あり' if with_timestamps else 'なし'}")
            print(f"- 議事録生成: {'あり' if generate_minutes else 'なし'}")
            
            # ステータス更新
            self.update_status("音声ファイルを読み込み中...")
            self.update_result("音声ファイルを読み込み中です。しばらくお待ちください...")
            if generate_minutes:
                self.update_minutes("議事録を準備中です...")
            
            # 音声ファイルを読み込む
            try:
                if not os.path.exists(filepath):
                    error_msg = f"ファイルが見つかりません: {filepath}\n"
                    error_msg += f"現在の作業ディレクトリ: {os.getcwd()}\n"
                    error_msg += f"ファイルの親ディレクトリ: {os.path.dirname(filepath)}\n"
                    error_msg += f"親ディレクトリの存在: {os.path.exists(os.path.dirname(filepath))}\n"
                    error_msg += f"親ディレクトリの内容: {os.listdir(os.path.dirname(filepath)) if os.path.exists(os.path.dirname(filepath)) else 'N/A'}"
                    print(f"\n6. ファイル存在エラー:")
                    print(error_msg)
                    self.update_result(error_msg, True)
                    self.finish_processing()
                    return
                
                print(f"\n6. 音声ファイル読み込み開始")
                audio, format_name = load_audio_file(filepath)
                print(f"音声ファイルの読み込みに成功しました")
            except Exception as e:
                error_msg = f"音声ファイルの読み込みに失敗しました:\n"
                error_msg += f"エラーの種類: {type(e).__name__}\n"
                error_msg += f"エラーメッセージ: {str(e)}\n"
                error_msg += f"現在の作業ディレクトリ: {os.getcwd()}\n"
                error_msg += f"ファイルパス: {filepath}\n"
                error_msg += f"システム情報:\n"
                error_msg += f"- OS: {sys.platform}\n"
                error_msg += f"- Python: {sys.version}\n"
                error_msg += f"- 文字コード: {sys.getfilesystemencoding()}"
                print(f"\n6. 音声ファイル読み込みエラー:")
                print(error_msg)
                self.update_result(error_msg, True)
                self.finish_processing()
                return
            
            # 音声の長さを取得（分）
            duration_minutes = len(audio) / (1000 * 60)
            
            # 更新されたステータスを表示
            timestamp_status = "タイムスタンプあり" if with_timestamps else "タイムスタンプなし"
            minutes_status = "議事録生成あり" if generate_minutes else "議事録生成なし"
            
            status_message = f"音声の長さ: {duration_minutes:.1f}分、処理を開始します..."
            self.update_status(status_message)
            
            result_info = (f"音声ファイル: {os.path.basename(filepath)}\n"
                          f"長さ: {duration_minutes:.1f}分\n"
                          f"言語: {language}\n"
                          f"{timestamp_status}\n"
                          f"{minutes_status}\n"
                          f"モデル: {model}\n\n"
                          f"処理を開始します...\n")
            
            self.update_result(result_info)
            
            if generate_minutes:
                self.update_minutes("文字起こし完了後に議事録を生成します...\n" + result_info)
            
            # 独自のコールバック関数を定義して、リアルタイムで進捗を更新
            segment_results = []
            current_segment = [0, 0]  # [処理中のセグメント番号, 合計セグメント数]
            
            def progress_callback(message):
                # メッセージに基づいて進捗状況を更新
                if "セグメント" in message and "処理中" in message:
                    # セグメント処理開始のメッセージからセグメント番号と合計を抽出
                    import re
                    match = re.search(r"セグメント (\d+)/(\d+)", message)
                    if match:
                        current_segment[0] = int(match.group(1))
                        current_segment[1] = int(match.group(2))
                        
                        # 進捗状況の更新
                        progress_message = f"セグメント {current_segment[0]}/{current_segment[1]} を処理中... ({(current_segment[0]-1)/current_segment[1]*100:.1f}% 完了)"
                        self.update_status(progress_message)
                        
                        # 最新の状態を結果エリアに追加
                        current_text = self.result_text.get(1.0, tk.END)
                        if "セグメント" in current_text and "処理中" in current_text:
                            lines = current_text.split("\n")
                            new_lines = []
                            updated = False
                            for line in lines:
                                if "セグメント" in line and "処理中" in line and not updated:
                                    new_lines.append(progress_message)
                                    updated = True
                                else:
                                    new_lines.append(line)
                            
                            if not updated:
                                new_lines.append(progress_message)
                            
                            self.update_result("\n".join(new_lines))
                        else:
                            self.update_result(current_text + "\n" + progress_message)
                
                # セグメント完了メッセージの処理
                elif "セグメント" in message and "完了しました" in message:
                    # 結果を保存
                    segment_results.append(message)
                    
                    # 進捗状況の更新
                    progress_percent = current_segment[0] / current_segment[1] * 100 if current_segment[1] > 0 else 0
                    self.update_status(f"セグメント {current_segment[0]}/{current_segment[1]} が完了しました。({progress_percent:.1f}% 完了)")
                
                # 議事録生成中のメッセージ
                elif "議事録を生成中" in message:
                    self.update_status("議事録を生成中...")
                    if generate_minutes:
                        self.update_minutes("議事録を生成中です。しばらくお待ちください...")
            
            # 元の標準出力を保存
            original_print = print
            
            # printをオーバーライドして進捗を追跡
            def custom_print(*args, **kwargs):
                message = " ".join(map(str, args))
                # 元のprintで出力
                original_print(*args, **kwargs)
                # コールバックを呼び出して進捗を更新
                self.root.after(0, progress_callback, message)
            
            # Pythonのprint関数を置き換え
            import builtins
            builtins.print = custom_print
            
            try:
                # 文字起こしを実行（議事録生成オプション付き）
                if generate_minutes:
                    transcription, minutes = transcribe_audio(
                        audio, 
                        model_name=model, 
                        language=language, 
                        with_timestamps=with_timestamps,
                        generate_minutes_flag=True
                    )
                    # 議事録を保存
                    self.current_minutes = minutes
                    
                    # 議事録を表示
                    self.root.after(0, self.update_minutes, minutes)
                else:
                    transcription = transcribe_audio(
                        audio, 
                        model_name=model, 
                        language=language, 
                        with_timestamps=with_timestamps
                    )
                
                # 自動保存の処理
                if self.autosave_var.get():
                    # 文字起こし結果の自動保存
                    output_filepath = Path(filepath).with_suffix('.txt')
                    with open(output_filepath, "w", encoding="utf-8") as f:
                        f.write(transcription)
                    
                    output_message = f"\n\n[文字起こし結果をファイルに保存しました: {output_filepath}]"
                    transcription += output_message
                    
                    # 議事録も自動保存
                    if generate_minutes:
                        minutes_filepath = Path(filepath).stem + "_minutes.md"
                        minutes_filepath = Path(filepath).parent / minutes_filepath
                        with open(minutes_filepath, "w", encoding="utf-8") as f:
                            f.write(self.current_minutes)
                        
                        minutes_message = f"\n\n[議事録をファイルに保存しました: {minutes_filepath}]"
                        self.current_minutes += minutes_message
                        self.root.after(0, self.update_minutes, self.current_minutes)
                
                # 文字起こし結果を保存
                self.current_result = transcription
                
                # UIを更新
                self.root.after(0, self.update_result, transcription)
                self.update_status("処理完了")
                
            finally:
                # 元のprint関数を復元
                builtins.print = original_print
        
        except Exception as e:
            error_message = f"エラーが発生しました: {str(e)}"
            self.root.after(0, self.update_result, error_message, True)
            if self.minutes_var.get():
                self.root.after(0, self.update_minutes, "議事録の生成中にエラーが発生しました。", True)
            self.update_status("エラーが発生しました")
        finally:
            # 処理完了後にUIを元に戻す
            self.root.after(0, self.finish_processing)
    
    def update_status(self, message):
        """ステータスメッセージを更新する（スレッドセーフ）"""
        self.root.after(0, lambda: self.status_var.set(message))
    
    def update_result(self, text, is_error=False):
        """UIスレッドで文字起こし結果表示を更新する"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, text)
        
        if is_error:
            self.result_text.tag_configure("error", foreground="red")
            self.result_text.tag_add("error", "1.0", tk.END)
    
    def update_minutes(self, text, is_error=False):
        """UIスレッドで議事録表示を更新する"""
        self.minutes_text.delete(1.0, tk.END)
        self.minutes_text.insert(tk.END, text)
        
        if is_error:
            self.minutes_text.tag_configure("error", foreground="red")
            self.minutes_text.tag_add("error", "1.0", tk.END)
    
    def finish_processing(self):
        """処理完了後にUIを元に戻す"""
        self.processing = False
        self.execute_button.config(state=tk.NORMAL, text="文字起こしを実行")
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        
        # 結果があれば保存ボタンを有効に
        if self.current_result:
            self.save_button.config(state=tk.NORMAL)
        
        # 議事録があれば保存ボタンを有効に
        if self.current_minutes:
            self.save_minutes_button.config(state=tk.NORMAL)
    
    def save_result(self):
        """文字起こし結果をファイルに保存する"""
        if not self.current_result:
            return
        
        # 保存先を選択
        filepath = filedialog.asksaveasfilename(
            title="文字起こし結果を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        
        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(self.current_result)
                messagebox.showinfo("保存完了", f"文字起こし結果を保存しました: {filepath}")
            except Exception as e:
                messagebox.showerror("エラー", f"保存中にエラーが発生しました: {str(e)}")
    
    def save_minutes(self):
        """議事録をファイルに保存する"""
        if not self.current_minutes:
            return
        
        # 保存先を選択
        filepath = filedialog.asksaveasfilename(
            title="議事録を保存",
            defaultextension=".md",
            filetypes=[("マークダウンファイル", "*.md"), ("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        
        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(self.current_minutes)
                messagebox.showinfo("保存完了", f"議事録を保存しました: {filepath}")
            except Exception as e:
                messagebox.showerror("エラー", f"保存中にエラーが発生しました: {str(e)}")

    def toggle_api_visibility(self):
        """APIキーの表示/非表示を切り替える"""
        if self.show_api_var.get():
            self.api_entry.config(show="")
        else:
            self.api_entry.config(show="*")
    
    def save_api_key(self):
        """APIキーを保存する"""
        api_key = self.api_var.get().strip()
        if not api_key:
            messagebox.showerror("エラー", "APIキーが入力されていません。")
            return
        
        # 設定ファイルに保存
        self.config["api_key"] = api_key
        self.save_config()
        
        # 環境変数にも設定
        os.environ["GOOGLE_API_KEY"] = api_key
        
        messagebox.showinfo("成功", "APIキーが保存されました。")
    
    def load_config(self):
        """設定ファイルを読み込む"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except Exception as e:
                print(f"設定ファイルの読み込みエラー: {e}")
        return {}
    
    def save_config(self):
        """設定ファイルを保存する"""
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config, f)
        except Exception as e:
            print(f"設定ファイルの保存エラー: {e}")
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")

def main():
    root = tk.Tk()
    app = TranscribeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 