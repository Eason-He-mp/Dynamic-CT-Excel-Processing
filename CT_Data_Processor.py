import os
import re
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

def extract_number(filename):
    """从文件名中提取末尾的数字用于正确排序"""
    numbers = re.findall(r'\d+', filename)
    return int(numbers[-1]) if numbers else 0

class CTDataProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("动态CT扫描数据整理工具 v4.1 (双模式完整版)")
        self.root.geometry("680x760")
        self.root.resizable(False, False)
        
        # 变量定义
        self.scan_type = tk.StringVar(value="Type1") 
        
        self.tif_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        self.num_scans = tk.IntVar(value=0)
        self.tifs_per_scan = tk.IntVar(value=0)
        self.time_offset = tk.DoubleVar(value=10.0)
        
        self.col_time = tk.StringVar(value="Time")
        self.col_force = tk.StringVar(value="Force")
        self.col_disp = tk.StringVar(value="Displacement")
        
        self.create_widgets()

    def create_widgets(self):
        # ==================== 0. 模式选择区 ====================
        frame_mode = ttk.LabelFrame(self.root, text="扫描模式选择", padding=10)
        frame_mode.pack(fill="x", padx=10, pady=5)
        
        ttk.Radiobutton(frame_mode, text="Type 1 (Dynamic)", variable=self.scan_type, value="Type1", command=self.on_mode_change).pack(side="left", padx=20)
        ttk.Radiobutton(frame_mode, text="Type 2 (Timelapse)", variable=self.scan_type, value="Type2", command=self.on_mode_change).pack(side="left", padx=20)

        # ==================== 1. 文件路径设置区 ====================
        frame_path = ttk.LabelFrame(self.root, text="文件路径设置", padding=10)
        frame_path.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_path, text="TIF 主文件夹:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.tif_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_tif).grid(row=0, column=2)

        ttk.Label(frame_path, text="力学 Excel 文件:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.excel_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_excel).grid(row=1, column=2)

        ttk.Label(frame_path, text="结果保存路径:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_output).grid(row=2, column=2)

        # ==================== 2. 参数设置区 ====================
        frame_params = ttk.LabelFrame(self.root, text="扫描与时间参数 (选择TIF文件夹后自动计算)", padding=10)
        frame_params.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_params, text="总扫描次数(n):").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_scans = ttk.Entry(frame_params, textvariable=self.num_scans, width=15, state='readonly')
        self.entry_scans.grid(row=0, column=1, sticky="w")

        ttk.Label(frame_params, text="单次扫描TIF数(x):").grid(row=0, column=2, sticky="w", padx=(20,0))
        self.entry_tifs = ttk.Entry(frame_params, textvariable=self.tifs_per_scan, width=15, state='readonly')
        self.entry_tifs.grid(row=0, column=3, sticky="w")

        ttk.Label(frame_params, text="时间差 (分钟):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_params, textvariable=self.time_offset, width=15).grid(row=1, column=1, sticky="w")
        ttk.Label(frame_params, text="(注：CT机时间 减去 力学机时间)", foreground="gray").grid(row=1, column=2, columnspan=2, sticky="w", padx=5)

        # ==================== 3. 表头设置区 ====================
        frame_cols = ttk.LabelFrame(self.root, text="Excel 表头名称设置", padding=10)
        frame_cols.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_cols, text="时间列名:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_cols, textvariable=self.col_time, width=15).grid(row=0, column=1, padx=5)
        
        ttk.Label(frame_cols, text="力列名:").grid(row=0, column=2, sticky="w", padx=(10,0))
        ttk.Entry(frame_cols, textvariable=self.col_force, width=15).grid(row=0, column=3, padx=5)
        
        ttk.Label(frame_cols, text="位移列名:").grid(row=0, column=4, sticky="w", padx=(10,0))
        ttk.Entry(frame_cols, textvariable=self.col_disp, width=15).grid(row=0, column=5, padx=5)

        # ==================== 4. 操作与日志区 ====================
        self.btn_start = ttk.Button(self.root, text="开 始 处 理", command=self.start_processing_thread)
        self.btn_start.pack(pady=15, ipadx=20, ipady=5)

        ttk.Label(self.root, text="处理日志:").pack(anchor="w", padx=10)
        
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.txt_log = tk.Text(log_frame, height=10, state='disabled', bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)
        
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    # --- 交互事件 ---
    def on_mode_change(self):
        """当工程师切换模式时，自动重新分析当前文件夹"""
        mode = self.scan_type.get()
        self.log(f"\n[*] 已切换至 {mode} 模式")
        current_folder = self.tif_folder.get()
        if current_folder and os.path.exists(current_folder):
            self.analyze_tif_folder(current_folder)

    # --- 核心：自动分析 TIF 文件夹 ---
    def analyze_tif_folder(self, folder):
        mode = self.scan_type.get()
        self.log(f"正在分析文件夹 ({mode} 模式): {folder}")
        
        try:
            if mode == "Type1":
                # Type 1 的分析逻辑 (主文件夹, recon子文件夹)
                raw_files = [f for f in os.listdir(folder) if f.lower().endswith(('.tif', '.tiff'))]
                all_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                
                m = len(all_files)
                if m == 0:
                    self.log("警告: 未在主文件夹中找到有效的实际 CT 图像！")
                    return

                recon_path = os.path.join(folder, 'recon')
                if not os.path.exists(recon_path):
                    self.log("警告: 未找到 'recon' 子文件夹，无法自动计算扫描参数。")
                    return
                    
                subdirs = [d for d in os.listdir(recon_path) if os.path.isdir(os.path.join(recon_path, d))]
                n = len(subdirs)
                
                if n == 0: return
                x = (m - 1) // n + 1
                
                self.num_scans.set(n)
                self.tifs_per_scan.set(x)
                self.log(f"-> 自动解析成功！检测到扫描次数(n)={n}, 首次扫描数(x)={x}")

            elif mode == "Type2":
                # Type 2 的分析逻辑 (寻找 nnnn_xx 子文件夹)
                # 正则表达式 r'^\d{4}' 用于匹配以4个数字开头的文件夹
                subfolders = [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d)) and re.match(r'^\d{4}', d)]
                n = len(subfolders)
                
                if n == 0:
                    self.log("警告: 未在主路径下找到符合 'nnnn_xx' 命名规则的子文件夹！")
                    self.num_scans.set(0)
                    self.tifs_per_scan.set(0)
                    return

                # 读取第一个子文件夹，计算单次扫描的有效 TIF 数
                first_sub = os.path.join(folder, subfolders[0])
                raw_files = [f for f in os.listdir(first_sub) if f.lower().endswith(('.tif', '.tiff'))]
                valid_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
                x = len(valid_files)

                self.num_scans.set(n)
                self.tifs_per_scan.set(x)
                
                self.log(f"-> 自动解析成功！")
                self.log(f"-> 检测到符合规则的子文件夹 (Scan数) = {n}")
                self.log(f"-> 检测到单个 Scan 内有效 TIF 数 = {x}")

        except Exception as e:
            self.log(f"自动分析文件夹时出错: {str(e)}")

    def browse_tif(self):
        folder = filedialog.askdirectory(title="选择TIF图像主文件夹")
        if folder: 
            self.tif_folder.set(folder)
            self.analyze_tif_folder(folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(title="选择力学Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file: self.excel_path.set(file)

    def browse_output(self):
        file = filedialog.asksaveasfilename(title="保存结果", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file: self.output_path.set(file)

    def log(self, message):
        def append():
            self.txt_log.config(state='normal')
            self.txt_log.insert(tk.END, message + "\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state='disabled')
        self.root.after(0, append)

    def start_processing_thread(self):
        if not self.tif_folder.get() or not self.excel_path.get() or not self.output_path.get():
            messagebox.showwarning("提示", "请先选择所有必要的文件和文件夹路径！")
            return
            
        if self.num_scans.get() <= 0:
            messagebox.showwarning("提示", "未检测到有效的扫描数据，请检查TIF文件夹结构！")
            return
            
        self.btn_start.config(state='disabled')
        self.txt_log.config(state='normal')
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.config(state='disabled')
        
        thread = threading.Thread(target=self.process_data_router)
        thread.daemon = True
        thread.start()

    def process_data_router(self):
        try:
            mode = self.scan_type.get()
            self.log(f">>> 开始执行 {mode} 模式处理任务...\n")
            
            if mode == "Type1":
                self.process_type1()
            elif mode == "Type2":
                self.process_type2()
                
        except Exception as e:
            self.log(f"[错误] 发生异常: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
        finally:
            self.root.after(0, lambda: self.btn_start.config(state='normal'))

    # ==========================================
    # Type 2 处理逻辑 (子文件夹独立模式)
    # ==========================================
    def process_type2(self):
        # 1. 读取Excel
        self.log("正在读取力学Excel数据...")
        df_mech = pd.read_excel(self.excel_path.get())
        
        time_col = self.col_time.get()
        force_col = self.col_force.get()
        disp_col = self.col_disp.get()
        
        for col in [time_col, force_col, disp_col]:
            if col not in df_mech.columns:
                raise ValueError(f"Excel中找不到表头: '{col}'，请检查设置。")
        
        df_mech[time_col] = pd.to_datetime(df_mech[time_col])
        
        # 2. 寻找并排序子文件夹
        tif_dir = self.tif_folder.get()
        subfolders = [d for d in os.listdir(tif_dir) if os.path.isdir(os.path.join(tif_dir, d)) and re.match(r'^\d{4}', d)]
        
        if not subfolders:
            raise FileNotFoundError("未找到符合 'nnnn_xx' 命名规则的子文件夹！")
            
        # 根据前4位数字进行排序
        subfolders.sort(key=lambda d: int(re.match(r'^(\d{4})', d).group(1)))
        
        # 3. 遍历子文件夹提取时间
        scan_records = []
        offset_mins = self.time_offset.get()
        
        for d in subfolders:
            # 提取前4位数字作为 Scan Rank
            rank = int(re.match(r'^(\d{4})', d).group(1))
            sub_path = os.path.join(tif_dir, d)
            
            # 读取该子文件夹下的所有TIF，并排除 di 和 io
            raw_files = [f for f in os.listdir(sub_path) if f.lower().endswith(('.tif', '.tiff'))]
            valid_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
            
            if not valid_files:
                self.log(f"警告: 子文件夹 {d} 中没有有效的 TIF 文件，已跳过。")
                continue
                
            valid_files.sort(key=extract_number)
            
            # 提取独立的第一帧和最后一帧
            first_file = valid_files[0]
            last_file = valid_files[-1]
            
            time1_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(sub_path, first_file)))
            time2_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(sub_path, last_file)))
            
            target_time1 = time1_ct - timedelta(minutes=offset_mins)
            target_time2 = time2_ct - timedelta(minutes=offset_mins)
            
            if target_time1 > target_time2:
                target_time1, target_time2 = target_time2, target_time1
                
            scan_records.append({
                'Scan Rank': rank,
                'Folder Name': d,  # 记录子文件夹名字方便核对
                'CT Start Time': time1_ct,
                'CT End Time': time2_ct,
                'Target Start Time': target_time1,
                'Target End Time': target_time2
            })
            
            self.log(f"Scan {rank} ({d}) 提取完成 | 包含 {len(valid_files)} 张有效图")

        # 4. 匹配数据
        self.log("\n正在计算时间段内的力和位移平均值...")
        final_results = []
        
        for record in scan_records:
            t1 = record['Target Start Time']
            t2 = record['Target End Time']
            
            mask = (df_mech[time_col] >= t1) & (df_mech[time_col] <= t2)
            df_filtered = df_mech.loc[mask]
            
            if not df_filtered.empty:
                avg_force = df_filtered[force_col].mean()
                avg_disp = df_filtered[disp_col].mean()
                matched_points = len(df_filtered)
            else:
                mid_time = t1 + (t2 - t1) / 2
                time_diff = (df_mech[time_col] - mid_time).abs()
                closest_idx = time_diff.idxmin()
                closest_row = df_mech.loc[closest_idx]
                
                avg_force = closest_row[force_col]
                avg_disp = closest_row[disp_col]
                matched_points = 1
                self.log(f"警告: Scan {record['Scan Rank']} 时间段内无数据点，已使用最近单点代替。")

            final_results.append({
                'Scan Rank': record['Scan Rank'],
                'Folder Name': record['Folder Name'], # Type 2 专属：输出对应的文件夹名
                'CT Start Time (PC2)': record['CT Start Time'],
                'CT End Time (PC2)': record['CT End Time'],
                'Mech Target Start (PC1)': t1,
                'Mech Target End (PC1)': t2,
                'Data Points Averaged': matched_points,
                'Average Force': avg_force,
                'Average Displacement': avg_disp
            })

        # 5. 导出结果
        self.log("正在生成结果Excel...")
        df_result = pd.DataFrame(final_results)
        df_result.to_excel(self.output_path.get(), index=False)
        
        self.log(f">>> 处理完成！结果已保存至:\n{self.output_path.get()}")
        messagebox.showinfo("成功", "数据处理完成！")

    # ==========================================
    # Type 1 处理逻辑 (主文件夹共享边界模式)
    # ==========================================
    def process_type1(self):
        self.log("正在读取力学Excel数据...")
        df_mech = pd.read_excel(self.excel_path.get())
        
        time_col = self.col_time.get()
        force_col = self.col_force.get()
        disp_col = self.col_disp.get()
        
        for col in [time_col, force_col, disp_col]:
            if col not in df_mech.columns:
                raise ValueError(f"Excel中找不到表头: '{col}'，请检查设置。")
        
        df_mech[time_col] = pd.to_datetime(df_mech[time_col])
        
        self.log("正在读取并排序有效TIF文件...")
        tif_dir = self.tif_folder.get()
        raw_files = [f for f in os.listdir(tif_dir) if f.lower().endswith(('.tif', '.tiff'))]
        all_files = [f for f in raw_files if not f.lower().startswith(('di', 'io'))]
        
        all_files.sort(key=extract_number)
        
        scan_records = []
        num_scans = self.num_scans.get()
        tifs_per_scan = self.tifs_per_scan.get()
        offset_mins = self.time_offset.get()
        
        current_start_idx = 0
        
        for rank in range(1, num_scans + 1):
            count = tifs_per_scan if rank == 1 else tifs_per_scan - 1
            end_idx = current_start_idx + count
            
            actual_end_idx = min(end_idx, len(all_files))
            current_scan_files = all_files[current_start_idx:actual_end_idx]
            
            if not current_scan_files: break
                
            first_file = current_scan_files[0]
            last_file = current_scan_files[-1]
            
            time1_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(tif_dir, first_file)))
            time2_ct = datetime.fromtimestamp(os.path.getmtime(os.path.join(tif_dir, last_file)))
            
            target_time1 = time1_ct - timedelta(minutes=offset_mins)
            target_time2 = time2_ct - timedelta(minutes=offset_mins)
            
            if target_time1 > target_time2:
                target_time1, target_time2 = target_time2, target_time1
                
            scan_records.append({
                'Scan Rank': rank,
                'CT Start Time': time1_ct,
                'CT End Time': time2_ct,
                'Target Start Time': target_time1,
                'Target End Time': target_time2
            })
            
            self.log(f"Scan {rank} 提取完成 | 包含 {len(current_scan_files)} 张图")
            current_start_idx = actual_end_idx - 1
            if current_start_idx >= len(all_files) - 1: break

        self.log("\n正在计算时间段内的力和位移平均值...")
        final_results = []
        
        for record in scan_records:
            t1 = record['Target Start Time']
            t2 = record['Target End Time']
            
            mask = (df_mech[time_col] >= t1) & (df_mech[time_col] <= t2)
            df_filtered = df_mech.loc[mask]
            
            if not df_filtered.empty:
                avg_force = df_filtered[force_col].mean()
                avg_disp = df_filtered[disp_col].mean()
                matched_points = len(df_filtered)
            else:
                mid_time = t1 + (t2 - t1) / 2
                time_diff = (df_mech[time_col] - mid_time).abs()
                closest_idx = time_diff.idxmin()
                closest_row = df_mech.loc[closest_idx]
                
                avg_force = closest_row[force_col]
                avg_disp = closest_row[disp_col]
                matched_points = 1
                self.log(f"警告: Scan {record['Scan Rank']} 时间段内无数据点，已使用最近单点代替。")

            final_results.append({
                'Scan Rank': record['Scan Rank'],
                'CT Start Time (PC2)': record['CT Start Time'],
                'CT End Time (PC2)': record['CT End Time'],
                'Mech Target Start (PC1)': t1,
                'Mech Target End (PC1)': t2,
                'Data Points Averaged': matched_points,
                'Average Force': avg_force,
                'Average Displacement': avg_disp
            })

        self.log("正在生成结果Excel...")
        df_result = pd.DataFrame(final_results)
        df_result.to_excel(self.output_path.get(), index=False)
        
        self.log(f">>> 处理完成！结果已保存至:\n{self.output_path.get()}")
        messagebox.showinfo("成功", "数据处理完成！")

if __name__ == "__main__":
    root = tk.Tk()
    app = CTDataProcessorApp(root)
    root.mainloop()
