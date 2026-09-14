import os
import re
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

def extract_number(filename):
    numbers = re.findall(r'\d+', filename)
    return int(numbers[-1]) if numbers else 0

# ==========================================
# 智能读取器：专治各种“伪装的”工业机 Excel
# ==========================================
def smart_read_excel(file_path, nrows=None):
    try:
        return pd.read_excel(file_path, nrows=nrows)
    except Exception:
        pass

    encodings = ['utf-8', 'gbk', 'gb18030', 'utf-16']
    separators = ['\t', ',']
    
    for enc in encodings:
        for sep in separators:
            try:
                df = pd.read_csv(file_path, sep=sep, encoding=enc, nrows=nrows, on_bad_lines='skip')
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
                
    try:
        tables = pd.read_html(file_path, encoding='utf-8')
        if tables:
            df = tables[0]
            if nrows is not None:
                df = df.head(nrows)
            return df
    except Exception:
        pass

    raise ValueError("无法解析该 Excel 文件，可能是未知的编码或损坏的文件格式。")


# ==========================================
# 国际化语言字典 (i18n)
# ==========================================
LANG = {
    'cn': {
        'title': "动态CT扫描数据整理工具",
        'btn_lang': "🌐 English",
        'frame_path': "1. 文件路径设置",
        'lbl_tif': "TIF 主文件夹:",
        'lbl_excel': "力学 Excel 文件:",
        'lbl_output': "结果保存路径:",
        'btn_browse': "浏览...",
        'frame_params': "2. 扫描参数 (自动识别)",
        'lbl_mode': "检测到的模式:",
        'val_mode_wait': "等待选择文件夹...",
        'lbl_scans': "总扫描次数(n):",
        'lbl_tifs': "单次扫描TIF数(x):",
        'lbl_offset': "时间差 (分钟):",
        'lbl_offset_tip': "(注：CT机时间 减去 力学机时间)",
        'frame_cols': "3. 变量设置 (从Excel自动读取)",
        'lbl_time_col': "时间列 (基准):",
        'lbl_vars': "目标变量:",
        'btn_add_var': "➕ 添加变量",
        'btn_del_var': "➖ 删除变量",
        'btn_start': "开 始 处 理",
        'lbl_log': "处理日志:",
        'msg_warn_path': "请先选择所有必要的文件和文件夹路径！",
        'msg_warn_mode': "未检测到有效的扫描数据，请检查TIF文件夹结构！",
        'msg_warn_headers': "请确保已选择时间列和至少一个目标变量！",
        'msg_success': "数据处理完成！包含原始数据与结果的多Sheet表格已生成。",
        'log_read_excel': "正在智能解析 Excel 表头 (防乱码)...",
        'log_analyze': "正在分析文件夹...",
        'log_mode_1': "-> 自动识别为: Type 1 (主文件夹共享边界模式)",
        'log_mode_2': "-> 自动识别为: Type 2 (子文件夹独立Scan模式)",
        'log_err_mode': "警告: 未能识别出支持的扫描结构，请检查文件夹！",
        'sheet_raw': "原始数据",
        'sheet_result': "处理结果"
    },
    'en': {
        'title': "Dynamic CT Data Processor v5.2",
        'btn_lang': "🌐 中文",
        'frame_path': "1. File Paths",
        'lbl_tif': "TIF Main Folder:",
        'lbl_excel': "Mechanics Excel:",
        'lbl_output': "Output Save Path:",
        'btn_browse': "Browse...",
        'frame_params': "2. Scan Parameters (Auto-detected)",
        'lbl_mode': "Detected Mode:",
        'val_mode_wait': "Waiting for folder selection...",
        'lbl_scans': "Total Scans (n):",
        'lbl_tifs': "TIFs per Scan (x):",
        'lbl_offset': "Time Offset (min):",
        'lbl_offset_tip': "(CT Time minus Mechanics Time)",
        'frame_cols': "3. Variables (Auto-loaded from Excel)",
        'lbl_time_col': "Time Column (Ref):",
        'lbl_vars': "Target Variables:",
        'btn_add_var': "➕ Add Variable",
        'btn_del_var': "➖ Remove Variable",
        'btn_start': "S T A R T",
        'lbl_log': "Processing Log:",
        'msg_warn_path': "Please select all required paths!",
        'msg_warn_mode': "No valid scan data detected. Check folder structure!",
        'msg_warn_headers': "Please select a time column and at least one variable!",
        'msg_success': "Processing complete! Multi-sheet Excel generated.",
        'log_read_excel': "Smart reading Excel headers...",
        'log_analyze': "Analyzing folder...",
        'log_mode_1': "-> Auto-detected: Type 1 (Shared Boundary Mode)",
        'log_mode_2': "-> Auto-detected: Type 2 (Independent Subfolder Mode)",
        'log_err_mode': "Warning: Unrecognized folder structure!",
        'sheet_raw': "Raw Data",
        'sheet_result': "Processed Results"
    }
}

class CTDataProcessorApp:
    def __init__(self, root):
        self.root = root
        self.lang = 'cn'
        self.root.title(LANG[self.lang]['title'])
        self.root.geometry("720x820")
        self.root.resizable(False, False)
        
        self.detected_mode = tk.StringVar(value="")
        self.tif_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        self.num_scans = tk.IntVar(value=0)
        self.tifs_per_scan = tk.IntVar(value=0)
        self.time_offset = tk.DoubleVar(value=10.0)
        
        self.excel_headers = []
        self.time_col_var = tk.StringVar()
        self.target_vars = [] 
        self.var_comboboxes = [] 
        
        self.create_widgets()
        self.update_ui_text()

    def create_widgets(self):
        # ==================== 0. 顶部工具栏 (语言切换) ====================
        top_bar = ttk.Frame(self.root)
        top_bar.pack(fill="x", padx=10, pady=(10, 0))
        
        self.btn_lang = ttk.Button(top_bar, command=self.toggle_lang)
        self.btn_lang.pack(side="right")

        # ==================== 1. 文件路径设置区 ====================
        self.frame_path = ttk.LabelFrame(self.root, padding=10)
        self.frame_path.pack(fill="x", padx=10, pady=5)

        self.lbl_tif = ttk.Label(self.frame_path)
        self.lbl_tif.grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(self.frame_path, textvariable=self.tif_folder, width=55).grid(row=0, column=1, padx=5)
        self.btn_browse_tif = ttk.Button(self.frame_path, command=self.browse_tif)
        self.btn_browse_tif.grid(row=0, column=2)

        self.lbl_excel = ttk.Label(self.frame_path)
        self.lbl_excel.grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(self.frame_path, textvariable=self.excel_path, width=55).grid(row=1, column=1, padx=5)
        self.btn_browse_excel = ttk.Button(self.frame_path, command=self.browse_excel)
        self.btn_browse_excel.grid(row=1, column=2)

        self.lbl_output = ttk.Label(self.frame_path)
        self.lbl_output.grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(self.frame_path, textvariable=self.output_path, width=55).grid(row=2, column=1, padx=5)
        self.btn_browse_out = ttk.Button(self.frame_path, command=self.browse_output)
        self.btn_browse_out.grid(row=2, column=2)

        # ==================== 2. 参数设置区 ====================
        self.frame_params = ttk.LabelFrame(self.root, padding=10)
        self.frame_params.pack(fill="x", padx=10, pady=5)

        self.lbl_mode = ttk.Label(self.frame_params, font=('Arial', 9, 'bold'))
        self.lbl_mode.grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.val_mode = ttk.Label(self.frame_params, foreground="blue")
        self.val_mode.grid(row=0, column=1, columnspan=3, sticky="w", pady=(0, 10))

        self.lbl_scans = ttk.Label(self.frame_params)
        self.lbl_scans.grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(self.frame_params, textvariable=self.num_scans, width=15, state='readonly').grid(row=1, column=1, sticky="w")

        self.lbl_tifs = ttk.Label(self.frame_params)
        self.lbl_tifs.grid(row=1, column=2, sticky="w", padx=(20,0))
        ttk.Entry(self.frame_params, textvariable=self.tifs_per_scan, width=15, state='readonly').grid(row=1, column=3, sticky="w")

        self.lbl_offset = ttk.Label(self.frame_params)
        self.lbl_offset.grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(self.frame_params, textvariable=self.time_offset, width=15).grid(row=2, column=1, sticky="w")
        self.lbl_offset_tip = ttk.Label(self.frame_params, foreground="gray")
        self.lbl_offset_tip.grid(row=2, column=2, columnspan=2, sticky="w", padx=5)

        # ==================== 3. 动态表头设置区 ====================
        self.frame_cols = ttk.LabelFrame(self.root, padding=10)
        self.frame_cols.pack(fill="x", padx=10, pady=5)

        self.lbl_time_col = ttk.Label(self.frame_cols)
        self.lbl_time_col.grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.cb_time = ttk.Combobox(self.frame_cols, textvariable=self.time_col_var, state='readonly', width=30)
        self.cb_time.grid(row=0, column=1, sticky="w", pady=(0, 10))

        self.lbl_vars = ttk.Label(self.frame_cols)
        self.lbl_vars.grid(row=1, column=0, sticky="nw")
        
        self.vars_container = ttk.Frame(self.frame_cols)
        self.vars_container.grid(row=1, column=1, sticky="w")
        
        # 将添加和删除按钮放在一个小 Frame 里，垂直排列
        self.btn_action_frame = ttk.Frame(self.frame_cols)
        self.btn_action_frame.grid(row=1, column=2, sticky="nw", padx=10)

        self.btn_add_var = ttk.Button(self.btn_action_frame, command=self.add_variable_dropdown)
        self.btn_add_var.pack(fill="x", pady=(0, 5))
        
        self.btn_del_var = ttk.Button(self.btn_action_frame, command=self.remove_variable_dropdown)
        self.btn_del_var.pack(fill="x")

        # 默认添加三个下拉框 (时间、力、位移的默认占位)
        for _ in range(3):
            self.add_variable_dropdown()

        # ==================== 4. 操作与日志区 ====================
        self.btn_start = ttk.Button(self.root, command=self.start_processing_thread)
        self.btn_start.pack(pady=15, ipadx=20, ipady=5)

        self.lbl_log = ttk.Label(self.root)
        self.lbl_log.pack(anchor="w", padx=10)
        
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.txt_log = tk.Text(log_frame, height=10, state='disabled', bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)
        
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def toggle_lang(self):
        self.lang = 'en' if self.lang == 'cn' else 'cn'
        self.update_ui_text()

    def update_ui_text(self):
        t = LANG[self.lang]
        self.root.title(t['title'])
        self.btn_lang.config(text=t['btn_lang'])
        
        self.frame_path.config(text=t['frame_path'])
        self.lbl_tif.config(text=t['lbl_tif'])
        self.lbl_excel.config(text=t['lbl_excel'])
        self.lbl_output.config(text=t['lbl_output'])
        self.btn_browse_tif.config(text=t['btn_browse'])
        self.btn_browse_excel.config(text=t['btn_browse'])
        self.btn_browse_out.config(text=t['btn_browse'])
        
        self.frame_params.config(text=t['frame_params'])
        self.lbl_mode.config(text=t['lbl_mode'])
        if not self.detected_mode.get():
            self.val_mode.config(text=t['val_mode_wait'])
        self.lbl_scans.config(text=t['lbl_scans'])
        self.lbl_tifs.config(text=t['lbl_tifs'])
        self.lbl_offset.config(text=t['lbl_offset'])
        self.lbl_offset_tip.config(text=t['lbl_offset_tip'])
        
        self.frame_cols.config(text=t['frame_cols'])
        self.lbl_time_col.config(text=t['lbl_time_col'])
        self.lbl_vars.config(text=t['lbl_vars'])
        self.btn_add_var.config(text=t['btn_add_var'])
        self.btn_del_var.config(text=t['btn_del_var'])
        
        self.btn_start.config(text=t['btn_start'])
        self.lbl_log.config(text=t['lbl_log'])

    def add_variable_dropdown(self):
        var = tk.StringVar()
        self.target_vars.append(var)
        
        row_idx = len(self.target_vars) - 1
        cb = ttk.Combobox(self.vars_container, textvariable=var, state='readonly', width=30)
        cb.grid(row=row_idx, column=0, pady=2)
        self.var_comboboxes.append(cb)
        
        if self.excel_headers:
            cb['values'] = self.excel_headers

    def remove_variable_dropdown(self):
        # 至少保留 1 个变量输入框
        if len(self.var_comboboxes) > 1:
            self.target_vars.pop()
            cb = self.var_comboboxes.pop()
            cb.destroy()

    def load_excel_headers(self, file_path):
        self.log(LANG[self.lang]['log_read_excel'])
        try:
            df = smart_read_excel(file_path, nrows=0)
            self.excel_headers = list(df.columns)
            
            self.cb_time['values'] = self.excel_headers
            for cb in self.var_comboboxes:
                cb['values'] = self.excel_headers
                
            self.log(f"-> 成功加载 {len(self.excel_headers)} 个表头")
        except Exception as e:
            self.log(f"[Error] 读取 Excel 失败: {str(e)}")

    def analyze_tif_folder(self, folder):
        self.log(LANG[self.lang]['log_analyze'])
        try:
            recon_path = os.path.join(folder, 'recon')
            if os.path.exists(recon_path):
                self.detected_mode.set("Type1")
                self.val_mode.config(text=LANG[self.lang]['log_mode_1'], foreground="green")
                self.log(LANG[self.lang]['log_mode_1'])
                
                raw_files = [f for f in os.listdir(folder) if f.lower().endswith(('.tif', '.tiff'))]
                all_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                m = len(all_files)
                subdirs = [d for d in os.listdir(recon_path) if os.path.isdir(os.path.join(recon_path, d))]
                n = len(subdirs)
                if n > 0:
                    x = (m - 1) // n + 1
                    self.num_scans.set(n)
                    self.tifs_per_scan.set(x)
                return

            subfolders = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d)) and re.search(r'\d+', d)]
            if subfolders:
                self.detected_mode.set("Type2")
                self.val_mode.config(text=LANG[self.lang]['log_mode_2'], foreground="green")
                self.log(LANG[self.lang]['log_mode_2'])
                
                n = len(subfolders)
                first_sub = os.path.join(folder, subfolders[0])
                raw_files = [f for f in os.listdir(first_sub) if f.lower().endswith(('.tif', '.tiff'))]
                valid_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                x = len(valid_files)
                
                self.num_scans.set(n)
                self.tifs_per_scan.set(x)
                return

            self.detected_mode.set("")
            self.val_mode.config(text=LANG[self.lang]['log_err_mode'], foreground="red")
            self.log(LANG[self.lang]['log_err_mode'])
            self.num_scans.set(0)
            self.tifs_per_scan.set(0)

        except Exception as e:
            self.log(f"[Error]: {str(e)}")

    def browse_tif(self):
        folder = filedialog.askdirectory()
        if folder: 
            self.tif_folder.set(folder)
            self.analyze_tif_folder(folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file: 
            self.excel_path.set(file)
            self.load_excel_headers(file)

    def browse_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file: self.output_path.set(file)

    def log(self, message):
        def append():
            self.txt_log.config(state='normal')
            self.txt_log.insert(tk.END, message + "\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state='disabled')
        self.root.after(0, append)

    def start_processing_thread(self):
        t = LANG[self.lang]
        if not self.tif_folder.get() or not self.excel_path.get() or not self.output_path.get():
            messagebox.showwarning("Warning", t['msg_warn_path'])
            return
            
        if not self.detected_mode.get() or self.num_scans.get() <= 0:
            messagebox.showwarning("Warning", t['msg_warn_mode'])
            return
            
        selected_vars = [var.get() for var in self.target_vars if var.get().strip()]
        if not self.time_col_var.get() or not selected_vars:
            messagebox.showwarning("Warning", t['msg_warn_headers'])
            return
            
        self.btn_start.config(state='disabled')
        self.txt_log.config(state='normal')
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.config(state='disabled')
        
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()

    def process_data(self):
        try:
            mode = self.detected_mode.get()
            self.log(f">>> Start processing ({mode})...")
            
            self.log("Loading raw Excel data (Smart Mode)...")
            df_mech = smart_read_excel(self.excel_path.get())
            
            time_col = self.time_col_var.get()
            selected_vars = [var.get() for var in self.target_vars if var.get().strip()]
            
            df_mech[time_col] = pd.to_datetime(df_mech[time_col])
            
            scan_records = []
            offset_mins = self.time_offset.get()
            tif_dir = self.tif_folder.get()
            
            if mode == "Type1":
                raw_files = [f for f in os.listdir(tif_dir) if f.lower().endswith(('.tif', '.tiff'))]
                all_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                all_files.sort(key=extract_number)
                
                current_start_idx = 0
                for rank in range(1, self.num_scans.get() + 1):
                    count = self.tifs_per_scan.get() if rank == 1 else self.tifs_per_scan.get() - 1
                    end_idx = current_start_idx + count
                    actual_end_idx = min(end_idx, len(all_files))
                    current_scan_files = all_files[current_start_idx:actual_end_idx]
                    
                    if not current_scan_files: break
                        
                    t1_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(tif_dir, current_scan_files[0])))
                    t2_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(tif_dir, current_scan_files[-1])))
                    
                    t1_target, t2_target = t1_ct - timedelta(minutes=offset_mins), t2_ct - timedelta(minutes=offset_mins)
                    if t1_target > t2_target: t1_target, t2_target = t2_target, t1_target
                        
                    scan_records.append({
                        'Scan Rank': rank, 'Folder Name': 'N/A (Type1)',
                        'CT Start': t1_ct, 'CT End': t2_ct,
                        'Target Start': t1_target, 'Target End': t2_target
                    })
                    current_start_idx = actual_end_idx - 1

            elif mode == "Type2":
                subfolders = [d for d in os.listdir(tif_dir) if os.path.isdir(os.path.join(tif_dir, d)) and re.search(r'\d+', d)]
                subfolders.sort(key=extract_number)
                
                for d in subfolders:
                    rank = extract_number(d)
                    sub_path = os.path.join(tif_dir, d)
                    raw_files = [f for f in os.listdir(sub_path) if f.lower().endswith(('.tif', '.tiff'))]
                    valid_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                    if not valid_files: continue
                    valid_files.sort(key=extract_number)
                    
                    t1_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(sub_path, valid_files[0])))
                    t2_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(sub_path, valid_files[-1])))
                    
                    t1_target, t2_target = t1_ct - timedelta(minutes=offset_mins), t2_ct - timedelta(minutes=offset_mins)
                    if t1_target > t2_target: t1_target, t2_target = t2_target, t1_target
                        
                    scan_records.append({
                        'Scan Rank': rank, 'Folder Name': d,
                        'CT Start': t1_ct, 'CT End': t2_ct,
                        'Target Start': t1_target, 'Target End': t2_target
                    })

            self.log("Calculating averages...")
            final_results = []
            
            for record in scan_records:
                t1, t2 = record['Target Start'], record['Target End']
                mask = (df_mech[time_col] >= t1) & (df_mech[time_col] <= t2)
                df_filtered = df_mech.loc[mask]
                
                result_row = {
                    'Scan Rank': record['Scan Rank'],
                    'Folder Name': record['Folder Name'],
                    'CT Start Time': record['CT Start'],
                    'CT End Time': record['CT End'],
                    'Mech Target Start': t1,
                    'Mech Target End': t2,
                }
                
                if not df_filtered.empty:
                    result_row['Data Points Averaged'] = len(df_filtered)
                    for var_name in selected_vars:
                        result_row[f'Average {var_name}'] = df_filtered[var_name].mean()
                else:
                    mid_time = t1 + (t2 - t1) / 2
                    closest_idx = (df_mech[time_col] - mid_time).abs().idxmin()
                    closest_row = df_mech.loc[closest_idx]
                    
                    result_row['Data Points Averaged'] = 1
                    for var_name in selected_vars:
                        result_row[f'Average {var_name}'] = closest_row[var_name]
                    self.log(f"Warning: Scan {record['Scan Rank']} used nearest single point.")

                final_results.append(result_row)

            self.log("Exporting to multi-sheet Excel...")
            df_result = pd.DataFrame(final_results)
            
            t = LANG[self.lang]
            with pd.ExcelWriter(self.output_path.get(), engine='openpyxl') as writer:
                df_mech.to_excel(writer, sheet_name=t['sheet_raw'], index=False)
                df_result.to_excel(writer, sheet_name=t['sheet_result'], index=False)
            
            self.log(f">>> Done! Saved to:\n{self.output_path.get()}")
            messagebox.showinfo("Success", t['msg_success'])

        except Exception as e:
            self.log(f"[Error]: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.root.after(0, lambda: self.btn_start.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = CTDataProcessorApp(root)
    root.mainloop()
