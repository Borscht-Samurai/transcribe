import os
import argparse
import base64
from pathlib import Path
from dotenv import load_dotenv
import google.generativeai as genai
from pydub import AudioSegment
import numpy as np
import tempfile
import math
import time

# 環境変数をロード
load_dotenv()

# Google API キーを設定
api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    raise ValueError("GOOGLE_API_KEY が設定されていません。.env ファイルを確認してください。")

genai.configure(api_key=api_key)

# サポートされている音声フォーマット
SUPPORTED_FORMATS = [
    "wav", "flac", "mp3", "ogg", "webm", "mp4", 
    "amr", "3gp", "m4a", "opus", "speex"
]

# Geminiが処理できる最大音声長（分）
MAX_AUDIO_DURATION_MINUTES = 25

# 議事録の雛形
MINUTES_TEMPLATE = """
# 議事録

## 1. 会議情報
- 日時: 
- 場所: 
- 参加者: 
- 議題: 

## 2. 議事内容
### 2.1. 議題1: 
- 背景・目的: 
- 主要論点: 
- 決定事項: 
- 課題/次のステップ: 

## 3. アクションアイテム
- 担当者: 
- 内容: 
- 期限: 

## 4. 次回会議
- 日時: 
- 場所: 
- 予定議題: 

## 5. その他・備考

"""

def load_audio_file(file_path):
    """音声ファイルを読み込み、適切なフォーマットに変換します"""
    file_path = Path(file_path)
    extension = file_path.suffix.lower()[1:]
    
    if extension not in SUPPORTED_FORMATS:
        raise ValueError(f"サポートされていないファイル形式です: {extension}。サポートされる形式: {', '.join(SUPPORTED_FORMATS)}")
    
    try:
        audio = AudioSegment.from_file(file_path)
        return audio, extension
    except Exception as e:
        raise ValueError(f"音声ファイルの読み込みに失敗しました: {str(e)}")

def get_audio_segment_duration_minutes(audio_segment):
    """音声セグメントの長さを分で返します"""
    return len(audio_segment) / (1000 * 60)  # ミリ秒から分に変換

def split_audio_segments(audio_data, max_duration_minutes=MAX_AUDIO_DURATION_MINUTES):
    """長い音声ファイルを指定された長さに分割します"""
    audio_length_ms = len(audio_data)
    segment_length_ms = int(max_duration_minutes * 60 * 1000)
    
    # 分割数を計算
    num_segments = math.ceil(audio_length_ms / segment_length_ms)
    
    # ログに音声全体の長さと分割数を出力
    print(f"音声全体の長さ: {format_timestamp(audio_length_ms)} ({audio_length_ms}ms)")
    print(f"分割数: {num_segments}、各セグメントの最大長: {max_duration_minutes}分 ({segment_length_ms}ms)")
    
    segments = []
    for i in range(num_segments):
        start_ms = i * segment_length_ms
        end_ms = min((i + 1) * segment_length_ms, audio_length_ms)
        segment = audio_data[start_ms:end_ms]
        segments.append((segment, start_ms, end_ms))
        print(f"セグメント {i+1} 作成: {format_timestamp(start_ms)} - {format_timestamp(end_ms)} (長さ: {format_timestamp(end_ms-start_ms)})")
    
    return segments

def format_timestamp(ms):
    """ミリ秒をHH:MM:SS形式に変換します"""
    total_seconds = ms // 1000
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    
    if hours > 0:
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    else:
        return f"{minutes:02d}:{seconds:02d}"

def transcribe_audio_segment(segment, start_ms=0, model_name="gemini-2.0-flash", language="japanese", with_timestamps=False, max_retries=3):
    """単一の音声セグメントを文字起こしします
    
    Args:
        segment: 音声セグメント（AudioSegmentオブジェクト）
        start_ms: セグメントの開始位置（ミリ秒）
        model_name: 使用するGeminiモデル名
        language: 文字起こしする言語
        with_timestamps: タイムスタンプを付けるかどうか
        max_retries: 最大再試行回数
    """
    # 一時ファイルを作成して音声データを保存
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_file:
        # 16-bit PCM WAVに変換して保存
        segment = segment.set_sample_width(2)
        segment.export(temp_file.name, format="wav")
        temp_file_path = temp_file.name
    
    retries = 0
    last_error = None
    
    try:
        while retries <= max_retries:
            try:
                # File APIを使ってファイルをアップロード
                print(f"音声セグメントをアップロード中... (セグメント開始位置: {format_timestamp(start_ms)})")
                uploaded_file = genai.upload_file(temp_file_path)
                
                # Geminiモデルを設定
                model = genai.GenerativeModel(model_name)
                
                # 言語とタイムスタンプの有無に応じて指示を設定
                offset_info = f"このセグメントは全体の {format_timestamp(start_ms)} から始まります。" if start_ms > 0 else ""
                
                if language.lower() == "english":
                    if with_timestamps:
                        prompt = f"Transcribe this audio in English with timestamps. Add a timestamp at the beginning of each sentence or after a significant pause. Format timestamps as [MM:SS] or [HH:MM:SS] for longer audio. Please transcribe without omitting any words. Make sure to transcribe the ENTIRE audio file completely, from beginning to end. {offset_info if offset_info else ''}"
                    else:
                        prompt = f"Transcribe this audio in English. Please transcribe without omitting every word, word for word. Make sure to transcribe the ENTIRE audio file completely, from beginning to end. {offset_info if offset_info else ''}"
                else:  # デフォルトは日本語
                    if with_timestamps:
                        prompt = f"この音声を日本語で文字起こししてください。各文の始まりや、意味のある間の後にタイムスタンプを追加してください。タイムスタンプは[MM:SS]または長い音声の場合は[HH:MM:SS]の形式で追加してください。全ての言葉を省略せず、一言一句漏らさず文字起こしして下さい。必ず音声ファイル全体を最初から最後まで完全に書き起こしてください。{offset_info if offset_info else ''}"
                    else:
                        prompt = f"この音声を日本語で文字起こししてください。全ての言葉を省略せず、一言一句漏らさず文字起こしして下さい。必ず音声ファイル全体を最初から最後まで完全に書き起こしてください。{offset_info if offset_info else ''}"
                
                # 音声ファイルのアップロード結果を使ってコンテンツを生成
                print(f"文字起こし処理中... セグメント開始位置: {format_timestamp(start_ms)} (試行: {retries+1}/{max_retries+1})")
                response = model.generate_content([
                    prompt,
                    uploaded_file
                ])
                
                result_text = response.text
                
                # 結果が短すぎる場合は警告を表示
                segment_length_sec = len(segment) / 1000
                expected_min_chars = segment_length_sec * 1.5  # 1秒あたり最低1.5文字を期待
                
                if len(result_text) < expected_min_chars and retries < max_retries:
                    print(f"警告: 文字起こし結果が予想よりも短いです（{len(result_text)}文字、予想: {int(expected_min_chars)}文字以上）。再試行します...")
                    retries += 1
                    continue
                
                print(f"文字起こし完了: {format_timestamp(start_ms)} から {len(result_text)} 文字を取得しました")
                return result_text
                
            except Exception as e:
                last_error = e
                print(f"エラーが発生しました (試行 {retries+1}/{max_retries+1}): {str(e)}")
                
                if retries < max_retries:
                    retries += 1
                    print(f"{retries}秒後に再試行します...")
                    time.sleep(retries)  # 指数バックオフ
                else:
                    raise Exception(f"最大再試行回数に達しました。最後のエラー: {str(last_error)}")
    
    finally:
        # 一時ファイルを削除
        try:
            os.unlink(temp_file_path)
        except:
            pass

def generate_minutes(transcription, model_name="gemini-2.0-flash"):
    """文字起こし結果から議事録を生成します
    
    Args:
        transcription: 文字起こし結果のテキスト
        model_name: 使用するGeminiモデル名
    
    Returns:
        生成された議事録（マークダウン形式）
    """
    print("議事録を生成中...")
    
    try:
        # Geminiモデルを設定
        model = genai.GenerativeModel(model_name)
        
        # 議事録生成のためのプロンプト
        prompt = f"""
以下の会議の文字起こしから議事録を作成してください。マークダウン形式で出力してください。

会議の文字起こし:
```
{transcription}
```

議事録は以下の形式で作成してください：

# 議事録

## 1. 会議情報
- 日時: （文字起こしから推測できる場合は記入）
- 場所: （文字起こしから推測できる場合は記入）
- 参加者: （文字起こしから特定できる人物を記入）
- 議題: （文字起こしから主要な議題を抽出）

## 2. 議事内容
### 2.1. 議題1: （タイトルを記入）
- 背景・目的: （議題の背景や目的を簡潔に）
- 主要論点: （議論された主な点）
- 決定事項: （決定された内容）
- 課題/次のステップ: （残された課題や次に行うべきこと）

### 2.2. 議題2: （タイトルを記入）
- 背景・目的:
- 主要論点:
- 決定事項:
- 課題/次のステップ:
（必要に応じて議題項目を追加）

## 3. アクションアイテム
- 担当者: （担当者名）
- 内容: （タスクの内容）
- 期限: （期限が明示されていれば記入）
（各議題ごとに具体的なアクションを記入）

## 4. 次回会議
- 日時: （言及されていれば記入）
- 場所: （言及されていれば記入）
- 予定議題: （言及されていれば記入）

## 5. その他・備考
（その他の重要事項や備考）

議事録は簡潔かつ明確に作成し、重要な決定事項や次のステップを確実に含めてください。
文字起こしから情報が特定できない場合は、その項目は空欄にするか「情報なし」と記入してください。
"""
        
        # 文字起こし結果から議事録を生成
        response = model.generate_content(prompt)
        minutes = response.text
        
        print(f"議事録生成完了: {len(minutes)}文字")
        return minutes
        
    except Exception as e:
        error_message = f"議事録の生成中にエラーが発生しました: {str(e)}"
        print(error_message)
        return f"# 議事録生成エラー\n\n{error_message}\n\n## 元の文字起こし\n\n{transcription}"

def transcribe_audio(audio_data, model_name="gemini-2.0-flash", language="japanese", with_timestamps=False, generate_minutes_flag=False):
    """Gemini APIを使用して音声を文字起こしします
    長い音声の場合は自動的に分割して処理します
    
    Args:
        audio_data: 音声データ（AudioSegmentオブジェクト）
        model_name: 使用するGeminiモデル名
        language: 文字起こしする言語（"japanese"または"english"）
        with_timestamps: タイムスタンプを付けるかどうか
        generate_minutes_flag: 議事録も生成するかどうか
    
    Returns:
        文字起こし結果のテキスト、または(文字起こし結果, 議事録)のタプル
    """
    duration_minutes = get_audio_segment_duration_minutes(audio_data)
    
    # 音声の長さがMAX_AUDIO_DURATION_MINUTESより短い場合は分割せずに処理
    if duration_minutes <= MAX_AUDIO_DURATION_MINUTES:
        transcription = transcribe_audio_segment(audio_data, 0, model_name, language, with_timestamps)
    else:
        # 長い音声の場合は分割して処理
        print(f"音声の長さが{duration_minutes:.1f}分のため、{MAX_AUDIO_DURATION_MINUTES}分ごとに分割して処理します")
        segments = split_audio_segments(audio_data, MAX_AUDIO_DURATION_MINUTES)
        
        full_transcription = ""
        
        for i, (segment, start_ms, end_ms) in enumerate(segments):
            print(f"セグメント {i+1}/{len(segments)} を処理中 ({format_timestamp(start_ms)} - {format_timestamp(end_ms)})")
            
            # セグメントごとに文字起こし
            segment_transcription = transcribe_audio_segment(
                segment, start_ms, model_name, language, with_timestamps
            )
            
            # 結果を連結
            if i > 0:
                full_transcription += "\n\n"
            
            # セグメント情報を追加（タイムスタンプありの場合は先頭にセグメント情報を追加）
            if with_timestamps:
                segment_header = f"[{format_timestamp(start_ms)}] セグメント {i+1}/{len(segments)} の文字起こし結果:\n"
                full_transcription += segment_header
            
            full_transcription += segment_transcription
            
            # セグメント処理完了のログ
            print(f"セグメント {i+1}/{len(segments)} の処理が完了しました。現在の文字起こし結果の長さ: {len(full_transcription)}文字")
        
        print(f"全セグメントの処理が完了しました。最終的な文字起こし結果の長さ: {len(full_transcription)}文字")
        transcription = full_transcription
    
    # 議事録を生成するかどうか
    if generate_minutes_flag:
        minutes = generate_minutes(transcription, model_name="gemini-2.0-flash")
        return transcription, minutes
    
    return transcription

def main():
    parser = argparse.ArgumentParser(description="音声ファイルをGemini APIで文字起こしします")
    parser.add_argument("audio_file", help="文字起こしする音声ファイルのパス")
    parser.add_argument("-o", "--output", help="出力テキストファイル（指定しない場合は標準出力）")
    parser.add_argument("-m", "--model", default="gemini-2.0-flash", help="使用するGeminiモデル")
    parser.add_argument("-l", "--language", default="japanese", choices=["japanese", "english"], 
                      help="文字起こしする言語（japanese/english）")
    parser.add_argument("-t", "--timestamps", action="store_true", 
                      help="タイムスタンプを付けて出力する")
    parser.add_argument("--minutes", action="store_true",
                      help="議事録も生成する")
    parser.add_argument("--minutes-output", help="議事録の出力ファイル")
    
    args = parser.parse_args()
    
    try:
        print(f"音声ファイルを読み込み中: {args.audio_file}")
        audio, format_name = load_audio_file(args.audio_file)
        
        print(f"文字起こしを開始します（モデル: {args.model}、言語: {args.language}、タイムスタンプ: {'あり' if args.timestamps else 'なし'}、議事録生成: {'あり' if args.minutes else 'なし'}）...")
        
        if args.minutes:
            transcription, minutes = transcribe_audio(
                audio, 
                model_name=args.model, 
                language=args.language, 
                with_timestamps=args.timestamps,
                generate_minutes_flag=True
            )
            
            # 議事録の出力
            if args.minutes_output:
                with open(args.minutes_output, "w", encoding="utf-8") as f:
                    f.write(minutes)
                print(f"議事録を保存しました: {args.minutes_output}")
            else:
                print("\n===== 議事録 =====")
                print(minutes)
                print("=================")
        else:
            transcription = transcribe_audio(
                audio, 
                model_name=args.model, 
                language=args.language, 
                with_timestamps=args.timestamps
            )
        
        # 文字起こし結果の出力
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(transcription)
            print(f"文字起こし結果を保存しました: {args.output}")
        else:
            print("\n===== 文字起こし結果 =====")
            print(transcription)
            print("========================")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 