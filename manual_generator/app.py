import streamlit as st
import cv2
import os
import json
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill

# --- 初期設定 ---
st.set_page_config(page_title="マニュアルExcel生成ツール", layout="centered")

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

# --- 画面UI構成 ---
st.title("📊 マニュアルExcel自動生成ツール")
st.write("他のAIで作った「JSONテキスト」と「動画」を入れるだけで、綺麗なExcelマニュアルを生成します！（APIキー不要・完全無料）")

# 1. 動画のアップロード
uploaded_video = st.file_uploader("① マニュアル化する動画をアップロード (MP4/MOV)", type=["mp4", "mov"])

# 2. JSONテキストの入力エリア
st.write("② ChatGPTなどで作成した【JSONデータ】を以下に貼り付けてください。")
json_input = st.text_area(
    "JSONデータを貼り付け", 
    height=300, 
    placeholder='{\n  "title": "マニュアル名",\n  "sections": [\n    {"heading": "1. はじめに", "steps": [{"time": "00:10", "text": "説明"}]}\n  ]\n}'
)

if st.button("🚀 Excelマニュアルを作成する", type="primary"):
    if not uploaded_video:
        st.warning("動画ファイルをアップロードしてください。")
    elif not json_input.strip():
        st.warning("JSONデータを貼り付けてください。")
    else:
        with st.spinner("Excelファイルを作成中...（数秒で終わります）"):
            try:
                # JSONテキストをPythonのデータに変換（エラーチェック）
                try:
                    ai_data = json.loads(json_input)
                except json.JSONDecodeError:
                    st.error("❌ 貼り付けられたテキストが正しいJSONフォーマットではありません。形式を確認してください。")
                    st.stop()

                # 一時フォルダの準備
                temp_dir = tempfile.mkdtemp()
                video_path = os.path.join(temp_dir, "temp_video.mp4")
                with open(video_path, "wb") as f:
                    f.write(uploaded_video.read())

                # ==========================================
                # Excel文書の作成スタート
                # ==========================================
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "業務マニュアル"
                
                font_meiryo = Font(name='メイリオ')
                font_meiryo_bold = Font(name='メイリオ', bold=True)
                
                # 1行目：全体のタイトル
                ws.merge_cells('A1:C1')
                ws['A1'] = ai_data.get('title', '業務 操作マニュアル')
                ws['A1'].font = Font(name='メイリオ', size=16, bold=True)
                ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                ws.row_dimensions[1].height = 40
                
                # 2行目：全体の概要
                ws.merge_cells('A2:C2')
                ws['A2'] = ai_data.get('description', '')
                ws['A2'].font = font_meiryo
                ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
                ws.row_dimensions[2].height = 80
                
                # 3行目：ヘッダー（見出し）
                headers = ['STEP / 時間', '操作説明', '画面画像']
                header_fill = PatternFill(patternType='solid', fgColor='D9D9D9') # 薄いグレー
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_num, value=header)
                    cell.font = font_meiryo_bold
                    cell.alignment = Alignment(horizontal='center')
                    cell.fill = header_fill

                ws.column_dimensions['A'].width = 15
                ws.column_dimensions['B'].width = 60
                ws.column_dimensions['C'].width = 120

                current_row = 4
                step_counter = 1
                section_fill = PatternFill(patternType='solid', fgColor='4F81BD') # 青色
                
                # 各章（セクション）ごとの処理
                for section in ai_data.get('sections', []):
                    # ① 章のタイトル
                    ws.merge_cells(f'A{current_row}:C{current_row}')
                    time_range = section.get('time_range', '')
                    heading_text = f"{section.get('heading', '')} {f'[{time_range}]' if time_range else ''}"
                    cell = ws.cell(row=current_row, column=1, value=heading_text.strip())
                    cell.font = Font(name='メイリオ', size=14, bold=True, color='FFFFFF')
                    cell.fill = section_fill
                    cell.alignment = Alignment(vertical='center')
                    ws.row_dimensions[current_row].height = 30
                    current_row += 1
                    
                    # ② 章の概要説明
                    summary = section.get('summary', '')
                    if summary:
                        ws.merge_cells(f'A{current_row}:C{current_row}')
                        cell = ws.cell(row=current_row, column=1, value=summary)
                        cell.font = font_meiryo
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                        ws.row_dimensions[current_row].height = 60
                        current_row += 1
                        
                    # ③ 具体的なステップと画像の貼り付け
                    for step in section.get('steps', []):
                        time_str = step.get('time', '')
                        desc_text = step.get('text', '')
                        img_path = os.path.join(temp_dir, f"frame_{current_row}.jpg")
                        
                        cell_a = ws.cell(row=current_row, column=1, value=f"STEP {step_counter}\n({time_str})")
                        cell_a.font = font_meiryo_bold
                        cell_a.alignment = Alignment(vertical='top', horizontal='center', wrap_text=True)
                        
                        cell_b = ws.cell(row=current_row, column=2, value=desc_text)
                        cell_b.font = font_meiryo
                        cell_b.alignment = Alignment(vertical='top', wrap_text=True)
                        
                        # 動画から画像を切り出してExcelに貼る
                        if extract_frame(video_path, time_str, img_path):
                            img = ExcelImage(img_path)
                            target_width = 800
                            if img.width > 0:
                                target_height = int(img.height * (target_width / img.width))
                                img.width = target_width
                                img.height = target_height
                                ws.row_dimensions[current_row].height = (target_height * 0.75) + 15
                            else:
                                ws.row_dimensions[current_row].height = 200
                            ws.add_image(img, f'C{current_row}')
                        else:
                            ws.row_dimensions[current_row].height = 60
                            
                        current_row += 1
                        step_counter += 1
                
                # Excelファイルの保存
                excel_output_path = os.path.join(temp_dir, "Manual_Result.xlsx")
                wb.save(excel_output_path)
                st.success("🎉 Excelマニュアルが完成しました！")

                # ダウンロードボタン
                with open(excel_output_path, "rb") as f:
                    st.download_button(
                        label="📊 Excelマニュアルをダウンロード",
                        data=f,
                        file_name="Business_Manual.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"❌ 処理中にエラーが発生しました: {e}")
