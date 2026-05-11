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
        self.root.title("动态CT扫描数据整理工具 v1.0")
        self.root.geometry("650x700")
        self.root.resizable(False, False)
        
        # 变量定义
        self.tif_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        self.num_scans = tk.IntVar(value=5)
        self.tifs_per_scan = tk.IntVar(value=1200)
        self.time_offset = tk.DoubleVar(value=10.0)
        
        self.col_time = tk.StringVar(value="Time")
        self.col_force = tk.StringVar(value="Force")
        self.col_disp = tk.StringVar(value="Displacement")
        
        self.create_widgets()

    def create_widgets(self):
        # ==================== 1. 文件路径设置区 ====================
        frame_path = ttk.LabelFrame(self.root, text="文件路径设置", padding=10)
        frame_path.pack(fill="x", padx=10, pady=5)

        # TIF 文件夹
        ttk.Label(frame_path, text="TIF 图像文件夹:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.tif_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_tif).grid(row=0, column=2)

        # Excel 路径
        ttk.Label(frame_path, text="力学 Excel 文件:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.excel_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_excel).grid(row=1, column=2)

        # 输出 路径
        ttk.Label(frame_path, text="结果保存路径:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frame_path, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(frame_path, text="浏览...", command=self.browse_output).grid(row=2, column=2)

        # ==================== 2. 参数设置区 ====================
        frame_params = ttk.LabelFrame(self.root, text="扫描与时间参数", padding=10)
        frame_params.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_params, text="总扫描次数:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frame_params, textvariable=self.num_scans, width=15).grid(row=0, column=1, sticky="w")

        ttk.Label(frame_params, text="每次扫描TIF数量:").grid(row=0, column=2, sticky="w", padx=(20,0))
        ttk.Entry(frame_params, textvariable=self.tifs_per_scan, width=15).grid(row=0, column=3, sticky="w")

        ttk.Label(frame_params, text="时间差 (分钟):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_params, textvariable=self.time_offset, width=15).grid(row=1, column=1, sticky="w")
        ttk.Label(frame_params, text="(注：CT机时间 减去 力学机时间，如CT快10分钟则填10)", foreground="gray").grid(row=1, column=2, columnspan=2, sticky="w", padx=5)

        # ==================== 3. 表头设置区 ====================
        frame_cols = ttk.LabelFrame(self.root, text="Excel 表头名称设置 (请与原始Excel保持一致)", padding=10)
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
        
        # 日志文本框带滚动条
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.txt_log = tk.Text(log_frame, height=10, state='disabled', bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)
        
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    # --- 浏览文件对话框 ---
    def browse_tif(self):
        folder = filedialog.askdirectory(title="选择TIF图像文件夹")
        if folder: self.tif_folder.set(folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(title="选择力学Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file: self.excel_path.set(file)

    def browse_output(self):
        file = filedialog.asksaveasfilename(title="保存结果", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file: self.output_path.set(file)

    # --- 日志输出 ---
    def log(self, message):
        """线程安全的日志输出"""
        def append():
            self.txt_log.config(state='normal')
            self.txt_log.insert(tk.END, message + "\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state='disabled')
        self.root.after(0, append)

    # --- 多线程启动 ---
    def start_processing_thread(self):
        # 基本校验
        if not self.tif_folder.get() or not self.excel_path.get() or not self.output_path.get():
            messagebox.showwarning("提示", "请先选择所有必要的文件和文件夹路径！")
            return
            
        self.btn_start.config(state='disabled')
        self.txt_log.config(state='normal')
        self.txt_log.delete(1.0, tk.END)  # 清空日志
        self.txt_log.config(state='disabled')
        
        # 启动后台线程处理数据，防止GUI卡死
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()

    # --- 核心处理逻辑 ---
    def process_data(self):
        try:
            self.log(">>> 开始处理任务...")
            
            # 1. 读取Excel
            self.log("正在读取力学Excel数据...")
            df_mech = pd.read_excel(self.excel_path.get())
            
            time_col = self.col_time.get()
            force_col = self.col_force.get()
            disp_col = self.col_disp.get()
            
            # 检查表头是否存在
            for col in [time_col, force_col, disp_col]:
                if col not in df_mech.columns:
                    raise ValueError(f"Excel中找不到表头: '{col}'，请检查设置。")
            
            df_mech[time_col] = pd.to_datetime(df_mech[time_col])
            
            # 2. 读取并排序TIF
            self.log("正在读取并排序TIF文件...")
            tif_dir = self.tif_folder.get()
            all_files = [f for f in os.listdir(tif_dir) if f.lower().endswith(('.tif', '.tiff'))]
            
            if not all_files:
                raise FileNotFoundError("未在指定的文件夹中找到TIF文件！")

            all_files.sort(key=extract_number)
            self.log(f"共找到 {len(all_files)} 张TIF图片。")
            
            # 3. 分组并提取时间
            scan_records = []
            num_scans = self.num_scans.get()
            tifs_per_scan = self.tifs_per_scan.get()
            offset_mins = self.time_offset.get()
            
            expected_total = num_scans * tifs_per_scan
            if len(all_files) < expected_total:
                self.log(f"警告：TIF文件总数({len(all_files)})少于预期({expected_total})")

            for rank in range(1, num_scans + 1):
                start_idx = (rank - 1) * tifs_per_scan
                end_idx = min(rank * tifs_per_scan, len(all_files))
                
                current_scan_files = all_files[start_idx:end_idx]
                if not current_scan_files:
                    break
                    
                mid_idx = len(current_scan_files) // 2
                representative_file = current_scan_files[mid_idx]
                file_path = os.path.join(tif_dir, representative_file)
                
                # 获取CT时间
                timestamp = os.path.getmtime(file_path)
                ct_time = datetime.fromtimestamp(timestamp)
                
                # 时间微调 (减去差值)
                target_mech_time = ct_time - timedelta(minutes=offset_mins)
                
                scan_records.append({
                    'Scan Rank': rank,
                    'CT Time': ct_time,
                    'Target Mech Time': target_mech_time
                })
                self.log(f"Scan {rank} | CT时间: {ct_time.strftime('%H:%M:%S')} -> 对应力学时间: {target_mech_time.strftime('%H:%M:%S')}")

            # 4. 匹配数据
            self.log("正在匹配时间和力学数据...")
            final_results = []
            
            for record in scan_records:
                target_time = record['Target Mech Time']
                
                # 计算绝对差值并找到最接近的行
                time_diff = (df_mech[time_col] - target_time).abs()
                closest_idx = time_diff.idxmin()
                closest_row = df_mech.loc[closest_idx]
                
                final_results.append({
                    'Scan Rank': record['Scan Rank'],
                    'CT Time (PC2)': record['CT Time'],
                    'Matched Mech Time (PC1)': closest_row[time_col],
                    'Force': closest_row[force_col],
                    'Displacement': closest_row[disp_col]
                })

            # 5. 导出结果
            self.log("正在生成结果Excel...")
            df_result = pd.DataFrame(final_results)
            df_result.to_excel(self.output_path.get(), index=False)
            
            self.log(f">>> 处理完成！结果已保存至:\n{self.output_path.get()}")
            messagebox.showinfo("成功", "数据处理完成！")

        except Exception as e:
            self.log(f"[错误] 发生异常: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
        finally:
            # 恢复按钮状态
            self.root.after(0, lambda: self.btn_start.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = CTDataProcessorApp(root)
    root.mainloop()
