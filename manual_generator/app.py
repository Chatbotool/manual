import streamlit as st
import google.generativeai as genai
import cv2
import os
import json
import time
import tempfile
from docx import Document
from docx.shared import Inches

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

# --- 画面UI ---
st.set_page_config(page_title="AIマニュアル生成", layout="centered")
st.title("🎥 AIマニュアル自動生成ツール")
st.write("動画をアップロードするだけで、画像付きのWordマニュアルを作成します。")

# ユーザー入力エリア
api_key = st.text_input("1. Gemini APIキーを入力 (セキュリティのため隠れます)", type="password")
uploaded_video = st.file_uploader("2. 動画をアップロード (10MB〜30MB程度の短い動画を推奨)", type=["mp4", "mov"])

if st.button("🚀 マニュアルを作成する", type="primary"):
    if not api_key:
        st.error("APIキーを入力してください。")
    elif not uploaded_video:
        st.error("動画をアップロードしてください。")
    else:
        with st.spinner("AIが動画を解析中です...（数分かかる場合があります）"):
            try:
                # 準備
                genai.configure(api_key=api_key)
                temp_dir = tempfile.mkdtemp()
                
                # 動画を一時保存
                video_path = os.path.join(temp_dir, "temp_video.mp4")
                with open(video_path, "wb") as f:
                    f.write(uploaded_video.read())

                # 動画をGeminiにアップロード
                video_file = genai.upload_file(path=video_path)
                while video_file.state.name == "PROCESSING":
                    time.sleep(5)
                    video_file = genai.get_file(video_file.name)

                # AIに解析を依頼
                model = genai.GenerativeModel('gemini-1.5-flash')
                prompt = """
                この動画から業務マニュアルを作成してください。
                重要な操作をピックアップし、以下のJSON形式のみで出力してください。
                Markdown(```json等)や挨拶は一切不要です。
                [
                    {"time": "00:10", "text": "ログイン画面で入力します。"}
                ]
                """
                response = model.generate_content([prompt, video_file])
                
                # 結果を整形
                result_text = response.text.strip()
                if result_text.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1

---

### 🚀 フェーズ2：自分のパソコンでテスト起動する

GitHubに上げる前に、自分のパソコンでちゃんと画面が出るかテストします。

1. VS Codeの「ターミナル（黒い画面）」を開きます。
2. 以下のコマンドを入力して、必要な部品をパソコンに入れます。
   `pip install -r requirements.txt`
3. インストールが終わったら、以下のコマンドでアプリを起動します。
   `streamlit run app.py`
4. 自動的にブラウザが立ち上がり、アプリの画面が表示されれば大成功です！

---

### 🌐 フェーズ3：GitHub → Render へ

パソコンで動くことが確認できたら、いよいよ公開です。

1. **GitHubへアップロード (Push)**
   VS Codeの左側にある「ソース管理（枝分かれしたアイコン）」から、コミットしてGitHubのリポジトリに発行（Push）します。
2. **Renderで連携**
   [Render](https://render.com/) にログインし、「New」>「Web Service」を選択。
   あなたのGitHubリポジトリを選び、設定をそのまま進めて「Create Web Service」を押すだけです！

まずは **フェーズ2（自分のパソコンでのテスト起動）** まで進めてみましょう。
VS Codeのターミナルの開き方や、GitHubへのPushのやり方など、どこか分からない部分はありますか？詳しくサポートします！