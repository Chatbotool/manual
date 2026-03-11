import streamlit as st
import google.generativeai as genai
import cv2
import os
import json
import time
import tempfile
import re
from docx import Document
from docx.shared import Inches

# --- 初期設定 ---
st.set_page_config(page_title="AIマニュアル生成", layout="centered")

# Renderの環境変数からAPIキーを自動取得
API_KEY = os.environ.get("GEMINI_API_KEY")

# --- 画像切り出し関数 ---
def extract_frame(video_path, time_str, output_path):
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

# --- AI出力からJSON（リスト）を安全に抽出する関数 ---
def extract_json_from_text(text):
    # ```python や ```json などの装飾を無視して [ ] の中身だけを抽出する
    match = re.search(r'\[.*\]', text, re.DOTALL)
    if match:
        json_str = match.group(0)
        return json.loads(json_str)
    else:
        raise ValueError("AIの回答からリスト形式のデータを抽出できませんでした。")

# --- 画面UI ---
st.title("🎥 AIマニュアル自動生成ツール")
st.write("動画をアップロードするだけで、画像付きのWordマニュアルを作成します。")

if not API_KEY:
    st.error("⚠️ システムエラー: 環境変数に GEMINI_API_KEY が設定されていません。管理者に連絡してください。")
    st.stop()

uploaded_video = st.file_uploader("マニュアル化する動画をアップロード (MP4/MOV)", type=["mp4", "mov"])

if st.button("🚀 マニュアルを作成する", type="primary"):
    if not uploaded_video:
        st.error("動画をアップロードしてください。")
    else:
        with st.spinner("AIが動画を解析中です...（1〜3分程度かかります）"):
            try:
                # 準備
                genai.configure(api_key=API_KEY)
                temp_dir = tempfile.mkdtemp()
                
                # 動画を一時保存
                video_path = os.path.join(temp_dir, "temp_video.mp4")
                with open(video_path, "wb") as f:
                    f.write(uploaded_video.read())

                # 動画をGeminiにアップロード
                st.info("動画をAIサーバーへ送信中...")
                video_file = genai.upload_file(path=video_path)
                while video_file.state.name == "PROCESSING":
                    time.sleep(5)
                    video_file = genai.get_file(video_file.name)

                # AIに解析を依頼
                st.info("AIが動画を視聴し、マニュアルを執筆中...")
                model = genai.GenerativeModel('gemini-2.5-flash')
                prompt = """
                この動画から業務マニュアルを作成してください。
                重要な操作をピックアップし、以下のJSON形式のみで出力してください。
                Markdown(```json等)や挨拶は一切不要です。純粋な配列のみを返してください。
                [
                    {"time": "00:10", "text": "ログイン画面で入力します。"}
                ]
                """
                response = model.generate_content([prompt, video_file])
                
                # AIの回答からデータを安全に抽出
                ai_data = extract_json_from_text(response.text)

                # Word作成
                st.info("Wordファイルを生成中...")
                doc = Document()
                doc.add_heading('AI自動生成マニュアル', 0)
                docx_path = os.path.join(temp_dir, "Manual.docx")

                for i, step in enumerate(ai_data):
                    time_str = step['time']
                    img_path = os.path.join(temp_dir, f"step_{i}.jpg")
                    # 画像の切り出しが成功したらWordに書き込む
                    if extract_frame(video_path, time_str, img_path):
                        doc.add_heading(f"STEP {i+1}: {time_str}", level=1)
                        doc.add_paragraph(step['text'])
                        doc.add_picture(img_path, width=Inches(5.5))
                        doc.add_page_break()
                
                doc.save(docx_path)
                st.success("✨ マニュアルが完成しました！")

                # ダウンロードボタン
                with open(docx_path, "rb") as f:
                    st.download_button("📄 Wordファイルをダウンロード", data=f, file_name="AI_Manual.docx")

            except Exception as e:
                st.error(f"❌ エラーが発生しました: {e}")

