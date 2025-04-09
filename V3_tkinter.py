import os
import pandas as pd
import pdfplumber
import re
import time
from datetime import timedelta
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
from tkinter import messagebox
import queue

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF表格轉換工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 設定風格
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#ccc")
        self.style.configure("TLabel", padding=6)
        self.style.configure("TFrame", padding=10)
        
        # 建立訊息佇列
        self.message_queue = queue.Queue()
        
        # 建立主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 建立資料夾選擇框架
        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        # 輸入資料夾選擇
        input_frame = ttk.Frame(folder_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="輸入資料夾:").pack(side=tk.LEFT)
        self.input_folder_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_folder_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(input_frame, text="瀏覽...", command=self.browse_input_folder).pack(side=tk.LEFT)
        
        # 輸出資料夾選擇
        output_frame = ttk.Frame(folder_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="輸出資料夾:").pack(side=tk.LEFT)
        self.output_folder_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_folder_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="瀏覽...", command=self.browse_output_folder).pack(side=tk.LEFT)
        
        # 設定框架
        settings_frame = ttk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(settings_frame, text="每處理筆數記錄:").pack(side=tk.LEFT)
        self.record_interval_var = tk.StringVar(value="10000")
        ttk.Entry(settings_frame, textvariable=self.record_interval_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # 進度條
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.progress_frame, text="處理進度:").pack(side=tk.LEFT)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, length=100, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.progress_label = ttk.Label(self.progress_frame, text="0%")
        self.progress_label.pack(side=tk.LEFT)
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.start_button = ttk.Button(button_frame, text="開始轉換", command=self.start_conversion)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="停止轉換", command=self.stop_conversion, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # 日誌框架
        log_frame = ttk.LabelFrame(main_frame, text="轉換日誌")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 狀態列
        self.status_var = tk.StringVar(value="就緒")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 轉換執行緒
        self.conversion_thread = None
        self.stop_event = threading.Event()
        
        # 定期更新UI
        self.update_ui()
    
    def browse_input_folder(self):
        folder = filedialog.askdirectory(title="選擇包含PDF檔案的資料夾")
        if folder:
            self.input_folder_var.set(folder)
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="選擇CSV輸出資料夾")
        if folder:
            self.output_folder_var.set(folder)
    
    def log_message(self, message):
        self.message_queue.put(message)
    
    def update_ui(self):
        # 處理佇列中的訊息
        try:
            while True:
                message = self.message_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        
        # 每100毫秒更新一次UI
        self.root.after(100, self.update_ui)
    
    def start_conversion(self):
        input_folder = self.input_folder_var.get()
        output_folder = self.output_folder_var.get()
        
        if not input_folder or not output_folder:
            messagebox.showerror("錯誤", "請選擇輸入和輸出資料夾")
            return
        
        try:
            record_interval = int(self.record_interval_var.get())
        except ValueError:
            messagebox.showerror("錯誤", "記錄間隔必須是整數")
            return
        
        # 重設停止事件
        self.stop_event.clear()
        
        # 禁用開始按鈕，啟用停止按鈕
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # 清空日誌
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 更新狀態
        self.status_var.set("轉換中...")
        
        # 啟動轉換執行緒
        self.conversion_thread = threading.Thread(
            target=self.batch_process_pdfs,
            args=(input_folder, output_folder, record_interval)
        )
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def stop_conversion(self):
        if self.conversion_thread and self.conversion_thread.is_alive():
            self.stop_event.set()
            self.log_message("正在停止轉換，請稍候...")
            self.status_var.set("正在停止...")
            self.stop_button.config(state=tk.DISABLED)
    
    def extract_table_data_from_pdf(self, pdf_path, output_folder, record_interval):
        """從PDF檔案中提取表格數據並轉換為CSV格式"""
        # 記錄開始時間
        start_time = time.time()

        # 確保輸出資料夾存在
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # 獲取檔案名稱（不含副檔名）
        file_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_csv = os.path.join(output_folder, f"{file_name}.csv")

        # 初始化數據列表和計數器
        all_data = []
        record_count = 0

        try:
            # 打開PDF檔案
            with pdfplumber.open(pdf_path) as pdf:
                # 檢查是否要停止
                if self.stop_event.is_set():
                    return False, 0, 0

                # 定義判斷目標表格的函數
                def is_target_table(table):
                    """檢查是否為目標表格"""
                    if not table or len(table) < 2:
                        return False

                    # 檢查表格結構 - 看是否有特定的列標題特徵
                    if len(table[0]) >= 7:  # 確保表格至少有7列
                        header_row = [str(cell).strip() if cell else "" for cell in table[0]]
                        header_text = ' '.join(header_row)

                        # 檢查表頭是否包含這些關鍵欄位
                        key_columns = ["地段名稱", "地號", "面積", "繪製或檢討變更"]
                        matches = sum(key in header_text for key in key_columns)

                        if matches >= 2:
                            return True

                    # 檢查資料行 - 看是否包含類似地號格式的資料
                    if len(table) >= 2:
                        for row in table[1:3]:  # 檢查前幾行資料
                            if row and len(row) >= 6:
                                # 檢查是否包含類似地號的格式 (如 0001-0000)
                                for cell in row:
                                    if cell and isinstance(cell, str) and re.search(r'\d{4}-\d{4}', cell):
                                        return True

                    return False

                # 定義處理面積的函數
                def process_area(area_str):
                    """處理面積字串，移除空格並保留數字和小數點"""
                    if not area_str or not isinstance(area_str, str):
                        return ""

                    # 移除所有空格
                    area_str = area_str.replace(" ", "")

                    # 嘗試保留數字、小數點和逗號
                    area_str = ''.join(c for c in area_str if c.isdigit() or c in ['.', ','])

                    return area_str

                # 尋找真正的表格開始頁
                start_page = 0
                found_table_page = False

                # 遍歷所有頁面尋找目標表格
                for i, page in enumerate(pdf.pages):
                    # 檢查是否要停止
                    if self.stop_event.is_set():
                        return False, 0, 0

                    tables = page.extract_tables()

                    # 檢查每個表格是否為目標表格
                    for table in tables:
                        if is_target_table(table):
                            start_page = i
                            found_table_page = True
                            self.log_message(f"在第 {i+1} 頁找到目標表格")
                            break
                        
                    if found_table_page:
                        break
                    
                    # 檢查關鍵字
                    text = page.extract_text() or ""
                    if "第 1 頁" in text or "公開展覽草案" in text or "地籍資料版本" in text:
                        start_page = i
                        self.log_message(f"在第 {i+1} 頁找到關鍵字")
                        if not found_table_page:
                            self.log_message("但未找到目標表格，將從此頁開始處理")
                        break
                    
                # 如果找不到表格開始頁，預設從第1頁開始
                if not found_table_page and start_page == 0:
                    start_page = 0
                    self.log_message(f"警告：無法找到目標表格，將從第1頁開始處理")

                self.log_message(f"開始處理 {file_name} 從第 {start_page + 1} 頁開始...")

                # 從表格開始頁開始提取數據
                for i in range(start_page, len(pdf.pages)):
                    # 檢查是否要停止
                    if self.stop_event.is_set():
                        return False, record_count, time.time() - start_time

                    page = pdf.pages[i]

                    # 提取頁面文本，用於識別地區
                    page_text = page.extract_text() or ""

                    # 嘗試識別地區（台中市、高雄市等）
                    city_match = re.search(r'([\u4e00-\u9fff]{2,3}市)', page_text)
                    city = city_match.group(1) if city_match else "未知城市"

                    # 嘗試識別區域（三重區、北區等）
                    district_match = re.search(r'([\u4e00-\u9fff]{2,3}區)', page_text)
                    district = district_match.group(1) if district_match else "未知區域"

                    # 提取表格
                    tables = page.extract_tables()

                    for table in tables:
                        # 檢查表格是否有效且為目標表格
                        if is_target_table(table) or (table and len(table) >= 2 and any("地段名稱" in str(cell) or "地號" in str(cell) or "繪製或檢討變更" in str(cell) for cell in table[0] if cell)):
                            # 處理數據行
                            for row in table[1:]:  # 跳過標題行
                                # 檢查是否要停止
                                if self.stop_event.is_set():
                                    return False, record_count, time.time() - start_time

                                if not row or not any(cell for cell in row if cell):
                                    continue
                                
                                # 確保行有足夠的列
                                while len(row) < 12:
                                    row.append("")

                                # 檢查是否為有效數據行（包含地號或地段名稱）
                                if any(row) and (
                                    (isinstance(row[5], str) and re.search(r'\d+-\d+', row[5])) or  # 地號格式檢查
                                    (isinstance(row[4], str) and len(row[4]) > 0)  # 地段名稱檢查
                                ):
                                    # 處理面積欄位 - 移除空格
                                    area = process_area(row[6]) if len(row) > 6 else ""

                                    # 提取數據，處理不同地區的格式差異
                                    data = {
                                        "直轄市縣市名稱": row[0] if row[0] else city,
                                        "鄉鎮市區": row[1] if row[1] else district,
                                        "地政事務所代碼": row[2] if len(row) > 2 and row[2] else "",
                                        "地段代碼": row[3] if len(row) > 3 and row[3] else "",
                                        "地段名稱": row[4] if len(row) > 4 and row[4] else "",
                                        "地號": row[5] if len(row) > 5 and row[5] else "",
                                        "面積": area,  # 使用處理後的面積
                                        "繪製或檢討變更前_國土功能分區及其分類": row[7] if len(row) > 7 and row[7] else "",
                                        "繪製或檢討變更前_使用地編定類別": row[8] if len(row) > 8 and row[8] else "",
                                        "繪製或檢討變更後_國土功能分區及其分類": row[9] if len(row) > 9 and row[9] else "",
                                        "繪製或檢討變更後_使用地編定類別": row[10] if len(row) > 10 and row[10] else "",
                                        "備註": row[11] if len(row) > 11 and row[11] else ""
                                    }
                                    all_data.append(data)
                                    record_count += 1

                                    # 每處理指定筆數記錄一次
                                    if record_count % record_interval == 0:
                                        elapsed_time = time.time() - start_time
                                        self.log_message(f"已處理 {record_count} 筆資料，耗時: {timedelta(seconds=int(elapsed_time))}，"
                                                f"平均每筆: {elapsed_time/record_count:.4f} 秒")

            # 計算總處理時間
            total_time = time.time() - start_time

            # 創建DataFrame並保存為CSV
            if all_data and not self.stop_event.is_set():
                df = pd.DataFrame(all_data)
                df.to_csv(output_csv, index=False, encoding='utf-8-sig')
                self.log_message(f"成功將 {pdf_path} 轉換為 {output_csv}，共提取 {len(all_data)} 筆資料")
                self.log_message(f"總處理時間: {timedelta(seconds=int(total_time))}，平均每筆: {total_time/len(all_data):.4f} 秒")

                # 寫入處理統計信息到日誌檔案
                log_file = os.path.join(output_folder, "processing_log.txt")
                with open(log_file, "a", encoding="utf-8") as f:
                    f.write(f"{file_name}: 提取 {len(all_data)} 筆資料，"
                            f"處理時間: {timedelta(seconds=int(total_time))}，"
                            f"平均每筆: {total_time/len(all_data):.4f} 秒\n")

                return True, len(all_data), total_time
            else:
                if self.stop_event.is_set():
                    self.log_message(f"處理 {pdf_path} 被使用者中斷")
                else:
                    self.log_message(f"警告: 無法從 {pdf_path} 提取表格數據")
                return False, record_count, total_time

        except Exception as e:
            total_time = time.time() - start_time
            self.log_message(f"處理 {pdf_path} 時發生錯誤: {str(e)}")
            return False, record_count, total_time


    
    def batch_process_pdfs(self, input_folder, output_folder, record_interval):
        """批次處理資料夾中的所有PDF檔案"""
        # 確保輸出資料夾存在
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # 獲取所有PDF檔案
        pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            self.log_message(f"在 {input_folder} 中找不到PDF檔案")
            self.status_var.set("未找到PDF檔案")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            return
        
        self.log_message(f"找到 {len(pdf_files)} 個PDF檔案，開始處理...")
        
        # 初始化總計數據
        total_start_time = time.time()
        total_records = 0
        success_count = 0
        file_stats = []
        
        # 設定進度條最大值
        self.progress_var.set(0)
        
        # 批次處理所有PDF檔案
        for i, pdf_file in enumerate(pdf_files):
            # 檢查是否要停止
            if self.stop_event.is_set():
                break
            
            # 更新進度條
            progress_percent = (i / len(pdf_files)) * 100
            self.progress_var.set(progress_percent)
            self.progress_label.config(text=f"{progress_percent:.1f}%")
            
            pdf_path = os.path.join(input_folder, pdf_file)
            self.log_message(f"正在處理 ({i+1}/{len(pdf_files)}): {pdf_file}")
            
            success, records, process_time = self.extract_table_data_from_pdf(
                pdf_path, output_folder, record_interval
            )
            
            if success:
                success_count += 1
                total_records += records
            
            # 記錄每個檔案的統計信息
            file_stats.append({
                "檔案名稱": pdf_file,
                "成功處理": success,
                "記錄數": records,
                "處理時間(秒)": process_time,
                "平均每筆時間(秒)": process_time / records if records > 0 else 0
            })
        
        # 計算總處理時間
        total_time = time.time() - total_start_time
        
        # 生成總結報告
        self.log_message("\n" + "="*50)
        
        if self.stop_event.is_set():
            self.log_message("處理被使用者中斷!")
        else:
            self.log_message(f"處理完成! 成功轉換 {success_count}/{len(pdf_files)} 個檔案")
        
        self.log_message(f"總共提取 {total_records} 筆資料")
        self.log_message(f"總處理時間: {timedelta(seconds=int(total_time))}")
        if total_records > 0:
            self.log_message(f"平均每筆資料處理時間: {total_time/total_records:.4f} 秒")
        self.log_message("="*50)
        
        # 將檔案統計信息保存為CSV
        stats_df = pd.DataFrame(file_stats)
        stats_csv = os.path.join(output_folder, "processing_statistics.csv")
        stats_df.to_csv(stats_csv, index=False, encoding='utf-8-sig')
        self.log_message(f"處理統計信息已保存至: {stats_csv}")
        
        # 更新進度條到100%
        self.progress_var.set(100)
        self.progress_label.config(text="100%")
        
        # 更新狀態
        if self.stop_event.is_set():
            self.status_var.set("處理已中斷")
        else:
            self.status_var.set("處理完成")
        
        # 啟用開始按鈕，禁用停止按鈕
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

# 啟動應用程式
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
