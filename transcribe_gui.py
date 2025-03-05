import os
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from pathlib import Path
from dotenv import load_dotenv
import threading

# 既存のtranscribe.pyから関数をインポート
from transcribe import load_audio_file, transcribe_audio

# 環境変数をロード
load_dotenv()

class TranscribeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("音声文字起こしツール")
        self.root.geometry("800x1000")  # ウィンドウサイズを大きくする
        self.root.minsize(800, 600)  # 最小サイズも変更
        
        # フォントとスタイルの設定
        self.font_default = ("Yu Gothic UI", 10)
        self.font_heading = ("Yu Gothic UI", 12, "bold")
        
        # メインフレーム
        self.main_frame = tk.Frame(root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        self.title_label = tk.Label(
            self.main_frame, 
            text="Gemini 2.0 Flash 音声文字起こし・議事録作成ツール", 
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
        self.results_paned = ttk.PanedWindow(self.main_frame, orient=tk.VERTICAL)
        self.results_paned.pack(fill=tk.BOTH, expand=True, pady=10)
        
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
            height=10
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 議事録テキストエリア
        self.minutes_text = scrolledtext.ScrolledText(
            self.minutes_frame, 
            wrap=tk.WORD, 
            font=self.font_default,
            height=10
        )
        self.minutes_text.pack(fill=tk.BOTH, expand=True)
        
        # ボタンフレーム
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=(5, 0))
        
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
        
        filepath = filedialog.askopenfilename(
            title="音声ファイルを選択",
            filetypes=filetypes
        )
        
        if filepath:
            self.path_var.set(filepath)
    
    def execute_transcription(self):
        """文字起こし処理を実行する"""
        filepath = self.path_var.get().strip()
        
        if not filepath:
            messagebox.showerror("エラー", "音声ファイルを選択してください。")
            return
        
        if not os.path.exists(filepath):
            messagebox.showerror("エラー", f"ファイルが見つかりません: {filepath}")
            return
        
        if self.processing:
            return
        
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
            model = self.model_var.get()
            language = self.language_var.get()
            with_timestamps = self.timestamp_var.get()
            generate_minutes = self.minutes_var.get()
            
            # ステータス更新
            self.update_status("音声ファイルを読み込み中...")
            self.update_result("音声ファイルを読み込み中です。しばらくお待ちください...")
            if generate_minutes:
                self.update_minutes("議事録を準備中です...")
            
            # 音声ファイルを読み込む
            audio, format_name = load_audio_file(filepath)
            
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

def main():
    root = tk.Tk()
    app = TranscribeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 