import streamlit as st
import re

st.set_page_config(page_title="VBA 修改器", layout="centered")
st.title("🚀 VBA 參數配置工具")

uploaded_file = st.file_uploader("選擇 .bas 檔案", type=["bas", "txt"])

if uploaded_file is not None:
    raw_bytes = uploaded_file.getvalue()
    content = ""
    for enc in ["utf-8-sig", "utf-8", "big5", "cp950"]:
        try:
            content = raw_bytes.decode(enc)
            break
        except:
            continue
    
    if content:
        lines = content.splitlines()
        found_items = []
        
        # 強化版 Regex：支援 @Config:中文說明
        pattern = re.compile(r'^(.*=\s*)(["\']?)([^"\';\n]*)(["\']?)(.*\'\s*@Config[:：]?\s*(.*))$', re.IGNORECASE)
        
        for i, line in enumerate(lines):
            match = pattern.search(line)
            if match:
                var_name = match.group(1).split('=')[0].replace('Dim', '').replace('Set', '').strip()
                desc = match.group(6).strip() # 抓取冒號後的中文說明
                
                found_items.append({
                    "idx": i, 
                    "display": desc if desc else var_name, # 有說明就顯示說明，沒有就顯示變數名
                    "prefix": match.group(1),
                    "q1": match.group(2),
                    "val": match.group(3),
                    "q2": match.group(4),
                    "suffix": match.group(5)
                })

        if found_items:
            st.info(f"💡 貼心提醒：在 VBA 標記 '@Config:您的說明' 即可顯示中文")
            new_values = {}
            with st.form("edit_form"):
                # 使用 columns 讓介面更美觀
                for item in found_items:
                    new_values[item["idx"]] = st.text_input(f"📍 {item['display']}", value=item["val"])
                
                if st.form_submit_button("✅ 儲存修改並下載檔案"):
                    for idx, val in new_values.items():
                        it = next(x for x in found_items if x["idx"] == idx)
                        lines[idx] = f"{it['prefix']}{it['q1']}{val}{it['q2']}{it['suffix']}"
                    
                    final_txt = "\n".join(lines)
                    st.download_button(
                        label="📥 點我下載修改後的 .bas 檔案",
                        data=final_txt.encode('utf-8-sig'),
                        file_name=f"New_{uploaded_file.name}",
                        mime="text/plain"
                    )
        else:
            st.warning("檔案中找不到 '@Config' 標記。")
