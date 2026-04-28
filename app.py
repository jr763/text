import streamlit as st
import re

st.set_page_config(page_title="VBA 配置修改器", layout="centered")

st.title("VBA 參數線上修改器")
st.write("請上傳你的 `.bas` 檔案，修改後點擊下載。")

# 檔案上傳
uploaded_file = st.file_uploader("選擇 .bas 檔案", type=["bas", "txt"])

if uploaded_file is not None:
    # 讀取檔案內容
    content = uploaded_file.getvalue().decode("utf-8-sig", errors="ignore")
    lines = content.splitlines()

    # 正則表達式
    pattern = re.compile(r'(.*?\b(?:=|Set|Worksheets\(|Cells\())\s*(["\']?)([^"\', \)]+)(["\']?)(.*\'\s*@Config)')
    
    new_lines = lines.copy()
    found_any = False
    
    st.subheader("⚙️ 參數設定")
    
    # 建立表單讓使用者修改
    with st.form("edit_form"):
        for i, line in enumerate(lines):
            if "@Config" in line:
                match = pattern.search(line)
                if match:
                    found_any = True
                    raw_name = line.split('=')[0].replace('Dim', '').replace('Set', '').strip()
                    if '(' in raw_name: raw_name = raw_name.split('(')[0]
                    
                    # 顯示輸入框
                    label = f"行 {i+1}: {raw_name}"
                    new_val = st.text_input(label, value=match.group(3))
                    
                    # 更新內容
                    new_lines[i] = f"{match.group(1)}{match.group(2)}{new_val}{match.group(4)}{match.group(5)}"

        submitted = st.form_submit_button("產生新檔案")

    if found_any:
        if submitted:
            final_content = "\n".join(new_lines)
            st.success("檔案已產生！請點擊下方按鈕下載。")
            st.download_button(
                label="📥 下載修改後的 .bas 檔案",
                data=final_content,
                file_name=f"Modified_{uploaded_file.name}",
                mime="text/plain"
            )
    else:
        st.warning("檔案中找不到 '@Config' 標記，請檢查 VBA 程式碼。")
