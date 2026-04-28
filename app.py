import streamlit as st
import re

st.set_page_config(page_title="VBA 修改器", layout="centered")
st.title("VBA 參數線上修改器")

uploaded_file = st.file_uploader("選擇 .bas 檔案", type=["bas", "txt"])

if uploaded_file is not None:
    raw_bytes = uploaded_file.getvalue()
    content = ""
    # 嘗試多種編碼，包含繁體中文最常見的 big5
    for enc in ["utf-8-sig", "utf-8", "big5", "cp950"]:
        try:
            content = raw_bytes.decode(enc)
            break
        except:
            continue
    
    if not content:
        st.error("檔案編碼讀取失敗")
    else:
        lines = content.splitlines()
        found_items = []
        
        # 最強放寬版正則：只要有 = 且後面有 @Config 就抓
        # 抓取：1.等號前內容, 2.左引號, 3.數值, 4.右引號, 5.剩餘部分
        pattern = re.compile(r'^(.*=\s*)(["\']?)([^"\';\n]*)(["\']?)(.*\'\s*@Config.*)$', re.IGNORECASE)
        
        for i, line in enumerate(lines):
            match = pattern.search(line)
            if match:
                # 取得變數名稱
                display_name = match.group(1).split('=')[0].replace('Dim', '').replace('Set', '').strip()
                found_items.append({
                    "idx": i, 
                    "name": display_name, 
                    "prefix": match.group(1),
                    "q1": match.group(2),
                    "val": match.group(3),
                    "q2": match.group(4),
                    "suffix": match.group(5)
                })

        if found_items:
            st.success(f"偵測到 {len(found_items)} 個可修改參數！")
            new_values = {}
            with st.form("edit_form"):
                for item in found_items:
                    new_values[item["idx"]] = st.text_input(f"修改 {item['name']}", value=item["val"])
                
                if st.form_submit_button("產生新檔案並下載"):
                    for idx, val in new_values.items():
                        it = next(x for x in found_items if x["idx"] == idx)
                        lines[idx] = f"{it['prefix']}{it['q1']}{val}{it['q2']}{it['suffix']}"
                    
                    final_txt = "\n".join(lines)
                    st.download_button(
                        label="點我下載 .bas 檔案",
                        data=final_txt.encode('utf-8-sig'),
                        file_name=f"Mod_{uploaded_file.name}",
                        mime="text/plain"
                    )
        else:
            st.warning("還是找不到標記。請確認檔案內容是否有 ' @Config' ")
            st.code("範例：Dim x = 10 ' @Config") # 顯示範例給使用者對照
