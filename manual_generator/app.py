import streamlit as st
import cv2
import os
import time
import tempfile
import re
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font

# --- 初期設定 ---
st.set_page_config(page_title="動画マニュアル作成ツール", layout="centered")

# --- 画像切り出し関数 ---
def extract_frame(video_path, time_str, output_path):
    """指定された時間のフレームを動画から切り出して保存する"""
    try:
        # 00:00 の形式を秒数に変換
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
    except Exception as e:
        print(f"画像切り出しエラー ({time_str}): {e}")
        return False

# --- 入力テキストを解析する関数 ---
def parse_manual_text(raw_text):
    """
    「00:10 説明文」のようなテキストを行ごとに分解し、時間とテキストのリストにする
    """
    parsed_data = []
    lines = raw_text.strip().split('\n')
    
    # 時間(MM:SS または H:MM:SS) から始まる行を探す正規表現
    pattern = r'^(\d{1,2}:\d{2}(?::\d{2})?)[ \t　]+(.*)'
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        match = re.match(pattern, line)
        if match:
            time_str = match.group(1)
            desc_text = match.group(2).strip()
            parsed_data.append({"time": time_str, "text": desc_text})
            
    return parsed_data

# --- 画面UI構成 ---
st.title("動画マニュアル生成ツール (API不要版)")
st.write("動画とテキストを組み合わせ、画像付きの【Excelマニュアル】を安全・高速に作成します。")

# 1. 動画のアップロード
uploaded_video = st.file_uploader("1. マニュアル化する動画をアップロード (MP4/MOV)", type=["mp4", "mov"])

# 2. テキストの入力
st.write("2. マニュアルのテキストを入力してください")
st.info("入力ルール: 各行の先頭に「時間（分:秒）」を書き、スペースを空けて「説明文」を書いてください。")
default_text = """00:05 ログイン画面でユーザー名を入力します。
00:15 パスワードを入力し、ログインボタンを押します。
01:30 画面右上の「設定」アイコンをクリックします。"""

manual_text = st.text_area("テキスト入力エリア", value=default_text, height=200)

# 3. 実行ボタン
if st.button("Excelマニュアルを作成する", type="primary"):
    if not uploaded_video:
        st.warning("動画ファイルをアップロードしてください。")
    elif not manual_text.strip():
        st.warning("マニュアルのテキストを入力してください。")
    else:
        # テキストの解析
        ai_data = parse_manual_text(manual_text)
        
        if not ai_data:
            st.error("テキストの形式が正しくありません。「00:00 説明文」の形式で入力されているか確認してください。")
        else:
            with st.spinner("動画から画像を切り出し、Excelを作成中..."):
                try:
                    temp_dir = tempfile.mkdtemp()
                    
                    # 動画を一時保存
                    video_path = os.path.join(temp_dir, "temp_video.mp4")
                    with open(video_path, "wb") as f:
                        f.write(uploaded_video.read())

                    # ==========================================
                    # Excel文書の作成スタート
                    # ==========================================
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "業務マニュアル"
                    
                    # 共通のフォント設定（メイリオ）
                    font_meiryo = Font(name='メイリオ')
                    font_meiryo_bold = Font(name='メイリオ', bold=True)
                    
                    # 1行目：タイトル
                    ws['A1'] = '操作マニュアル'
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
                    ws.column_dimensions['C'].width = 120 # 画像列（画像幅800pxに最適化）

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
                            
                            original_width = img.width
                            original_height = img.height
                            
                            # 画像の横幅を800pxに設定（潰れないように比率計算）
                            target_width = 800
                            if original_width > 0:
                                target_height = int(original_height * (target_width / original_width))
                                img.width = target_width
                                img.height = target_height
                                
                                # 画像の高さに合わせてExcelの行の高さを自動調整
                                ws.row_dimensions[current_row].height = (target_height * 0.75) + 15
                            else:
                                ws.row_dimensions[current_row].height = 200
                            
                            # C列の該当するセルに画像を配置
                            ws.add_image(img, f'C{current_row}')
                        else:
                            # 画像が取得できなかった場合
                            ws.row_dimensions[current_row].height = 60
                            ws.cell(row=current_row, column=3, value="(画像取得失敗)").alignment = Alignment(vertical='center', horizontal='center')
                        
                        current_row += 1
                    
                    # Excelファイルの保存
                    excel_output_path = os.path.join(temp_dir, "Manual_Result.xlsx")
                    wb.save(excel_output_path)
                    st.success("マニュアルの作成が完了しました！")

                    # ダウンロードボタン
                    with open(excel_output_path, "rb") as f:
                        st.download_button(
                            label="Excelマニュアルをダウンロード",
                            data=f,
                            file_name="Business_Manual.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"処理中にエラーが発生しました: {e}")
