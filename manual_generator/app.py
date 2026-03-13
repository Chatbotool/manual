import streamlit as st
import google.generativeai as genai
import cv2
import os
import json
import time
import tempfile
import re
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font

# --- 初期設定 ---
st.set_page_config(page_title="AIマニュアル生成", layout="centered")

# Renderの環境変数からAPIキーを読み込む
API_KEY = os.environ.get("GEMINI_API_KEY")

# --- 画像切り出し関数 ---
def extract_frame(video_path, time_str, output_path):
    """指定された時間のフレームを動画から切り出して保存する"""
    try:
        minutes, seconds = map(float, time_str.split(':'))
        total_seconds = minutes * 60 + seconds
        cap = cv2.VideoCapture(video_path)
        fps = cap.get(cv2.CAP_PROP_FPS)
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(total_seconds * fps))
        ret, frame = cap.read()
        if ret:
            cv2.imwrite(output_path, frame)
        cap.release()
        return ret
    except Exception:
        return False

# --- AI出力からJSON（リスト）を安全に抽出する関数 ---
def extract_json_from_text(text):
    """AIが回答に余計な装飾を付けてもJSON部分だけを抜き出す"""
    match = re.search(r'\[.*\]', text, re.DOTALL)
    if match:
        json_str = match.group(0)
        return json.loads(json_str)
    else:
        raise ValueError("AIの回答からリスト形式のデータを抽出できませんでした。")

# --- 画面UI構成 ---
st.title("🎥 AIマニュアル自動作成ツール")
st.write("動画をアップロードするだけで、画像付きの【Excelマニュアル】を作成します。")

if not API_KEY:
    st.error("システム設定エラー: Renderの環境変数に GEMINI_API_KEY が設定されていません。")
    st.stop()

uploaded_video = st.file_uploader("マニュアル化する動画をアップロード (MP4/MOV)", type=["mp4", "mov"])

if st.button("Excelマニュアルを作成する", type="primary"):
    if not uploaded_video:
        st.warning("動画ファイルを先にアップロードしてください。")
    else:
        with st.spinner("AIが動画を解析し、Excelを作成中...（数分かかります）"):
            try:
                # 準備
                genai.configure(api_key=API_KEY)
                temp_dir = tempfile.mkdtemp()
                
                # 動画を一時保存
                video_path = os.path.join(temp_dir, "temp_video.mp4")
                with open(video_path, "wb") as f:
                    f.write(uploaded_video.read())

                # 動画をGeminiサーバーへ送信
                st.info("AIに動画を送信しています...")
                video_file = genai.upload_file(path=video_path)
                
                while video_file.state.name == "PROCESSING":
                    time.sleep(5)
                    video_file = genai.get_file(video_file.name)

                # AIに指示を出す
                st.info("AIが内容を分析し、マニュアルを執筆しています...")
                model = genai.GenerativeModel('gemini-2.5-flash')
                prompt = """
                この動画を解析して、操作マニュアルを日本語で詳しく作成してください。
                重要な操作ステップを抽出し、以下の純粋なJSONリスト形式のみで出力してください。
                Markdown(```json等)は一切含めないでください。
                [
                    {"time": "00:10", "text": "ログイン画面でユーザー名を入力します。"}
                ]
                """
                response = model.generate_content([prompt, video_file])
                
                # セキュリティ：即座に動画削除
                genai.delete_file(video_file.name)
                
                # JSON抽出
                ai_data = extract_json_from_text(response.text)

                # ==========================================
                # Excel文書の作成スタート
                # ==========================================
                st.info("画像を抽出してExcelファイルに整理しています...")
                
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "業務マニュアル"
                
                # 共通のフォント設定（メイリオ）
                font_meiryo = Font(name='メイリオ')
                font_meiryo_bold = Font(name='メイリオ', bold=True)
                
                # 1行目：タイトル
                ws['A1'] = 'AI自動生成 操作マニュアル'
                ws['A1'].font = Font(name='メイリオ', size=16, bold=True)
                
                # 2行目：ヘッダー（見出し）
                headers = ['STEP / 時間', '操作説明', '画面画像']
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_num, value=header)
                    cell.font = font_meiryo_bold
                    cell.alignment = Alignment(horizontal='center')

                # 列の幅を調整（見やすくするため）
                ws.column_dimensions['A'].width = 15  # STEP列
                ws.column_dimensions['B'].width = 50  # 説明列
                ws.column_dimensions['C'].width = 120 # 画像列（画像拡大に合わせて60→120に変更しました）

                # 3行目からデータを書き込んでいく
                current_row = 3
                for i, step in enumerate(ai_data):
                    time_str = step['time']
                    desc_text = step['text']
                    img_path = os.path.join(temp_dir, f"frame_{i}.jpg")
                    
                    # A列：STEPと時間
                    cell_a = ws.cell(row=current_row, column=1, value=f"STEP {i+1}\n({time_str})")
                    cell_a.font = font_meiryo_bold
                    cell_a.alignment = Alignment(vertical='top', horizontal='center', wrap_text=True)
                    
                    # B列：説明文
                    cell_b = ws.cell(row=current_row, column=2, value=desc_text)
                    cell_b.font = font_meiryo
                    cell_b.alignment = Alignment(vertical='top', wrap_text=True)
                    
                    # C列：画像の貼り付けとサイズ調整
                    if extract_frame(video_path, time_str, img_path):
                        img = ExcelImage(img_path)
                        
                        # 元のサイズを取得して縦横比（アスペクト比）を崩さずに計算
                        original_width = img.width
                        original_height = img.height
                        
                        # 👇 ここが画像の大きさを決める数字です（400→800に変更）
                        target_width = 800
                        if original_width > 0:
                            target_height = int(original_height * (target_width / original_width))
                            img.width = target_width
                            img.height = target_height
                            
                            # 画像の高さに合わせてExcelの行の高さを自動調整
                            # (Excelの行の高さ単位はピクセルではなくポイントなので、約0.75倍して余白を足す)
                            ws.row_dimensions[current_row].height = (target_height * 0.75) + 15
                        else:
                            ws.row_dimensions[current_row].height = 200
                        
                        # C列の該当するセルに画像を配置
                        ws.add_image(img, f'C{current_row}')
                    else:
                        # 画像がない場合
                        ws.row_dimensions[current_row].height = 60
                    
                    current_row += 1
                
                # Excelファイルの保存
                excel_output_path = os.path.join(temp_dir, "AI_Manual_Result.xlsx")
                wb.save(excel_output_path)
                st.success("全行程が完了しました！")

                # ダウンロードボタン
                with open(excel_output_path, "rb") as f:
                    st.download_button(
                        label=" Excelマニュアルをダウンロード",
                        data=f,
                        file_name="Business_Manual.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"処理中にエラーが発生しました: {e}")
