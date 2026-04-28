import streamlit as st
import re

st.set_page_config(page_title="VBA 配置修改器", layout="centered")
st.title("VBA 參數線上修改器")

uploaded_file = st.file_uploader("選擇 .bas 檔案", type=["bas", "txt"])

if uploaded_file is not None:
    # 嘗試多種編碼讀取
    raw_bytes = uploaded_file.getvalue()
    content = ""
    for enc in ["utf-8-sig", "utf-8", "big5", "cp950"]:
        try:
            content = raw_bytes.decode(enc)
            break
        except:
            continue
    
    if not content:
        st.error("無法辨識檔案編碼，請嘗試將檔案存為 UTF-8 格式後再上傳。")
    else:
        lines = content.splitlines()
        # 不分大小寫搜尋 @Config，且容許更多空格
        pattern = re.compile(r'(.*?\b(?:=|Set|Worksheets\(|Cells\())\s*(["\']?)([^"\', \)]+)(["\']?)(.*\'\s*@Config)', re.IGNORECASE)
        
        new_lines = lines.copy()
        found_items = []
        
        for i, line in enumerate(lines):
            if "@config" in line.lower():
                match = pattern.search(line)
                if match:
                    raw_name = line.split('=')[0].replace('Dim', '').replace('Set', '').strip()
                    if '(' in raw_name: raw_name = raw_name.split('(')[0]
                    found_items.append({"idx": i, "name": raw_name, "match": match})

        if found_items:
            st.subheader("⚙️ 參數設定")
            with st.form("edit_form"):
                for item in found_items:
                    m = item["match"]
                    # 建立輸入框
                    label = f"參數: {item['name']}"
                    val = st.text_input(label, value=m.group(3), key=f"input_{item['idx']}")
                    # 即時更新 new_lines
                    new_lines[item['idx']] = f"{m.group(1)}{m.group(2)}{val}{m.group(4)}{m.group(5)}"
                
                if st.form_submit_button("產生新檔案"):
                    final_content = "\n".join(new_lines)
                    st.success("檔案已產生！")
                    st.download_button(
                        label="📥 下載修改後的 .bas 檔案",
                        data=final_content.encode('utf-8-sig'), # 加上 BOM 確保 Excel 匯入不亂碼
                        file_name=f"Mod_{uploaded_file.name}",
                        mime="text/plain"
                    )
        else:
            st.warning("檔案中找不到 '@Config' 標記。請確認程式碼行末有加上 ' @Config")
