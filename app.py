import streamlit as st
import re

st.set_page_config(page_title="VBA 修改器", layout="centered")
st.title("🚀 VBA 參數配置工具")

# 上傳區
uploaded_file = st.file_uploader("選擇 .bas 檔案", type=["bas", "txt"])

if uploaded_file is not None:
    # 讀取檔案
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
        pattern = re.compile(r'^(.*=\s*)(["\']?)([^"\';\n]*)(["\']?)(.*\'\s*@Config[:：]?\s*(.*))$', re.IGNORECASE)
        
        for i, line in enumerate(lines):
            match = pattern.search(line)
            if match:
                var_name = match.group(1).split('=')[0].replace('Dim', '').replace('Set', '').strip()
                desc = match.group(6).strip()
                found_items.append({
                    "idx": i, "display": desc if desc else var_name,
                    "prefix": match.group(1), "q1": match.group(2),
                    "val": match.group(3), "q2": match.group(4), "suffix": match.group(5)
                })

        if found_items:
            st.divider()
            
            # --- 設定區 ---
            st.subheader("✍️ 修改參數與檔名")
            
            with st.form("edit_form"):
                # 1. 檔名輸入框 (放在最上面最明顯的地方)
                default_name = uploaded_file.name.replace(".bas", "")
                new_filename = st.text_input("💾 下載檔名 (不需輸入副檔名)", value=f"New_{default_name}")
                
                st.write("---")
                
                # 2. 參數修改框
                new_values = {}
                for item in found_items:
                    new_values[item["idx"]] = st.text_input(f"📍 {item['display']}", value=item["val"])
                
                submitted = st.form_submit_button("✅ 確認修改內容")

            # --- 下載區 ---
            if submitted or 'final_txt' in st.session_state:
                # 重新組合內容
                for idx, val in new_values.items():
                    it = next(x for x in found_items if x["idx"] == idx)
                    lines[idx] = f"{it['prefix']}{it['q1']}{val}{it['q2']}{it['suffix']}"
                
                final_txt = "\n".join(lines)
                st.session_state.final_txt = final_txt
                
                # 處理檔名：確保最後一定是 .bas
                # 處理檔名：確保最後一定是 .bas
                save_name = new_filename.strip()
                if not save_name.lower().endswith('.bas'):
                    save_name += ".bas"
                
                st.success(f"🎉 內容已更新！檔名將存為：{save_name}")
                
                # --- 關鍵修改區：改為 cp950 (Big5) 避免 VBA 亂碼 ---
                try:
                    download_data = final_txt.encode('cp950', errors='replace')
                except:
                    download_data = final_txt.encode('utf-8-sig') # 萬一有特殊字元才退回 utf-8
                
                st.download_button(
                    label=f"📥 立即點我下載檔案",
                    data=download_data,
                    file_name=save_name,
                    mime="text/plain"
                )
        else:
            st.warning("檔案中找不到 '@Config' 標記。")
