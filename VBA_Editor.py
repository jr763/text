import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import windnd
import re
import os

def handle_drop(files):
    # 支援多個檔案拖入，目前取第一個
    file_path = files[0].decode('gbk') if isinstance(files[0], bytes) else files[0]
    if not file_path.lower().endswith('.bas'):
        messagebox.showerror("錯誤", "這不是 .bas 檔案！")
        return
    analyze_and_edit(file_path)

def analyze_and_edit(file_path):
    # 嘗試不同編碼讀取 VBA 檔案 (常見為 UTF-8 或 Big5)
    lines = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except:
        try:
            with open(file_path, 'r', encoding='big5') as f:
                lines = f.readlines()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取檔案編碼：{e}")
            return

    # 正則表達式：捕捉包含 @Config 的行
    # 群組 1: 前綴, 2: 左引號, 3: 目前值, 4: 右引號, 5: 後綴(含@Config)
    pattern = re.compile(r'(.+?=\s*)(["\']?)([^"\';\n]*)(["\']?)(.*\'\s*@Config)')
    
    found_items = []
    for i, line in enumerate(lines):
        if "@Config" in line:
            match = pattern.search(line)
            if match:
                # 取得變數名稱作為視窗顯示文字
                display_name = line.split('=')[0].replace('Dim', '').replace('Set', '').strip()
                found_items.append({
                    'index': i,
                    'prefix': match.group(1),
                    'quote_start': match.group(2),
                    'current_value': match.group(3),
                    'quote_end': match.group(4),
                    'suffix': match.group(5),
                    'display_name': display_name
                })

    if not found_items:
        messagebox.showwarning("提示", "檔案內找不到結尾有 ' @Config' 的標記！\n請先在 VBA 程式碼欲修改處加上標記。")
        return

    # --- 建立 UI 修改視窗 ---
    edit_win = tk.Toplevel(root)
    edit_win.title(f"修改中：{os.path.basename(file_path)}")
    edit_win.geometry("550x450")
    
    # 頂部說明
    header = tk.Label(edit_win, text="偵測到以下標記參數，請修改後存檔：", font=("Microsoft JhengHei", 10, "bold"), pady=10)
    header.pack()

    # 內容捲動區
    canvas = tk.Canvas(edit_win)
    scrollbar = ttk.Scrollbar(edit_win, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas, padding=10)
    
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    entries = []
    for item in found_items:
        row = ttk.Frame(scroll_frame)
        row.pack(fill='x', pady=5)
        
        lbl = ttk.Label(row, text=f"{item['display_name']} :", width=25)
        lbl.pack(side='left')
        
        ent = ttk.Entry(row, width=35)
        ent.insert(0, item['current_value'])
        ent.pack(side='left', padx=5)
        entries.append((item, ent))

    def save_action():
        # 更新記憶體中的行內容
        for item, ent in entries:
            new_val = ent.get()
            # 重組 VBA 行：前綴 + 引號 + 新值 + 引號 + 後綴
            new_line = f"{item['prefix']}{item['quote_start']}{new_val}{item['quote_end']}{item['suffix']}\n"
            lines[item['index']] = new_line
        
        # 選擇存檔路徑
        save_path = filedialog.asksaveasfilename(
            title="儲存修改後的 VBA 檔案",
            defaultextension=".bas",
            initialfile=f"Mod_{os.path.basename(file_path)}",
            filetypes=[("VBA Files", "*.bas"), ("Text Files", "*.txt")]
        )
        
        if save_path:
            try:
                # 統一以帶簽名的 UTF-8 儲存，確保 Excel 匯入不亂碼
                with open(save_path, 'w', encoding='utf-8-sig') as f:
                    f.writelines(lines)
                messagebox.showinfo("成功", f"檔案已儲存：\n{save_path}")
                edit_win.destroy()
            except Exception as e:
                messagebox.showerror("儲存失敗", str(e))

    # 底部按鈕
    btn_frame = ttk.Frame(edit_win, padding=10)
    btn_frame.pack(fill='x')
    ttk.Button(btn_frame, text="執行修改並另存新檔", command=save_action).pack(pady=5)

# --- 主程式進入點 ---
root = tk.Tk()
root.title("VBA 精準修改器 v1.0")
root.geometry("400x250")

# 簡單的導引介面
main_label = tk.Label(root, text="\n【操作步驟】\n\n1. 在 .bas 關鍵行末加 ' @Config\n2. 將檔案拖曳到此視窗\n3. 在彈出視窗修改內容並存檔", 
                      font=("Microsoft JhengHei", 11), justify="left")
main_label.pack(expand=True)

# 註冊拖放功能
windnd.hook_dropfiles(root, func=handle_drop)

root.mainloop()
