import pandas as pd
import io
from ftplib import FTP
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# 萬陽配置
CONFIG = {
    "ftp_host": "192.168.1.240",
    "ftp_port": 8821,
    "ftp_username": "wanyung1",
    "ftp_password": "Wan26026601$",
    "ftp_target_folder": "/ZY_MA_Recoder",
    "excel_file": "ZY_MA_Recoder.xlsx",
}


def update_excel_structure():
    print("開始更新萬陽 Excel 結構...")
    try:
        # 1. 連接 FTP
        ftp = FTP()
        ftp.connect(CONFIG["ftp_host"], CONFIG["ftp_port"])
        ftp.login(CONFIG["ftp_username"], CONFIG["ftp_password"])
        ftp.cwd(CONFIG["ftp_target_folder"])
        print("FTP 連接成功")

        # 2. 下載 Excel
        buffer = io.BytesIO()
        try:
            ftp.retrbinary(f"RETR {CONFIG['excel_file']}", buffer.write)
            buffer.seek(0)
            df = pd.read_excel(buffer)
            print(f"成功下載 Excel，共 {len(df)} 筆資料")
        except Exception as e:
            print(f"下載失敗或檔案不存在: {e}")
            return

        # 3. 備份 Excel
        backup_filename = (
            f"{CONFIG['excel_file'].split('.')[0]}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        backup_buffer = io.BytesIO()
        df.to_excel(backup_buffer, index=False)
        backup_buffer.seek(0)
        ftp.storbinary(f"STOR {backup_filename}", backup_buffer)
        print(f"已備份原始檔案為: {backup_filename}")

        # 4. 更新欄位結構
        # 確保有 ID 欄位
        if "ID" not in df.columns:
            print("新增 ID 欄位...")
            # 為舊資料產生 ID
            # 假設舊資料有 '報修時間'，格式為 YYYY-MM-DD HH:MM:SS
            # 如果沒有，則使用當前日期
            new_ids = []
            date_counts = {}

            for index, row in df.iterrows():
                try:
                    report_time = pd.to_datetime(row["報修時間"])
                    date_str = report_time.strftime("%Y%m%d")
                except:
                    date_str = datetime.now().strftime("%Y%m%d")

                if date_str not in date_counts:
                    date_counts[date_str] = 0
                date_counts[date_str] += 1
                
                new_id = f"W{date_str}-{date_counts[date_str]:03d}"
                new_ids.append(new_id)
            
            df.insert(0, "ID", new_ids)
        
        # 確保有處理方式和處理時間欄位
        if "處理方式" not in df.columns:
            print("新增 處理方式 欄位...")
            df["處理方式"] = ""
        
        if "處理時間" not in df.columns:
            print("新增 處理時間 欄位...")
            df["處理時間"] = ""

        # 5. 上傳更新後的 Excel
        output_buffer = io.BytesIO()
        df.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        ftp.storbinary(f"STOR {CONFIG['excel_file']}", output_buffer)
        
        ftp.quit()
        print("Excel 結構更新完成！")
        messagebox.showinfo("成功", f"萬陽 Excel 結構更新完成！\n已備份為 {backup_filename}")

    except Exception as e:
        print(f"發生錯誤: {e}")
        messagebox.showerror("錯誤", f"更新失敗: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() # 隱藏主視窗
    update_excel_structure()
