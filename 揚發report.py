#!/usr/bin/env python
# coding: utf-8

# In[4]:


import os
import io
from ftplib import FTP
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import smtplib
from email.mime.text import MIMEText

# === NAS 與檔案設定 ===
CONFIG = {
    "ftp_host": "192.168.0.147",
    "ftp_port": 8821,
    "ftp_user": "service",
    "ftp_pass": "Zanyao0915",
    "ftp_folder": "/MA_recoder",
    "main_excel": "ZY_MA_Recoder.xlsx",
    "sender_email": "Zanyao.Service@msa.hinet.net",
    "smtp_password": "Zanyao0915$",
    "smtp_server": "msr.hinet.net",
    "smtp_port": 587,
    "extra_recipients": [
        "zanyao0925@gmail.com",
        "adrian@yang-fa.com.tw",
        "sheep255174@gmail.com",
        "Zanyao.Service@msa.hinet.net",
        "hanna@yang-fa.com.tw",
    ],
}


def download_excel_from_nas(filename):
    ftp = FTP()
    ftp.connect(CONFIG["ftp_host"], CONFIG["ftp_port"])
    ftp.login(CONFIG["ftp_user"], CONFIG["ftp_pass"])
    ftp.cwd(CONFIG["ftp_folder"])

    buf = io.BytesIO()
    ftp.retrbinary(f"RETR {filename}", buf.write)
    buf.seek(0)
    ftp.quit()
    return pd.read_excel(buf)


def upload_excel_to_nas(df, filename):
    ftp = FTP()
    ftp.connect(CONFIG["ftp_host"], CONFIG["ftp_port"])
    ftp.login(CONFIG["ftp_user"], CONFIG["ftp_pass"])
    ftp.cwd(CONFIG["ftp_folder"])

    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    ftp.storbinary(f"STOR {filename}", buf)
    ftp.quit()


def send_email(to_email, staff_id, description, solution):
    sender = CONFIG["sender_email"]
    subject = "報修完成處理"
    body = f"""您好：

您的報修案件已由工程師完成處理

工號：{staff_id}
問題描述：{description}
處理方式：{solution}

感謝您的耐心等候！
贊耀資訊"""

    msg = MIMEText(body, "plain", "utf-8")
    recipients = [to_email] + CONFIG["extra_recipients"]
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject  # ← 主旨要加在這裡

    try:
        with smtplib.SMTP(
            CONFIG["smtp_server"], CONFIG["smtp_port"], local_hostname="localhost"
        ) as server:
            server.starttls()
            server.login(CONFIG["sender_email"], CONFIG["smtp_password"])
            server.sendmail(sender, recipients, msg.as_string())
    except Exception as e:
        messagebox.showerror("錯誤", f"無法寄送通知信：{e}")


def engineer_ui_nas():
    def search_report():
        nonlocal matched_row, df_main

        date_part = entry_date.get().strip()
        num_part = entry_number.get().strip()

        if not date_part or not num_part:
            messagebox.showerror("錯誤", "請完整輸入日期與編號")
            return

        report_id = f"R{date_part}-{num_part.zfill(3)}"

        if not report_id:
            messagebox.showerror("錯誤", "請輸入報修單 ID")
            return
        try:
            df_main = download_excel_from_nas(CONFIG["main_excel"])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法下載主檔：{e}")
            return

        matched = df_main[df_main["ID"] == report_id]
        if matched.empty:
            messagebox.showinfo("查無資料", "找不到該報修單 ID。")
            return

        matched_row = matched.iloc[0]
        label_staff.config(text=matched_row["工號"])
        label_email.config(text=matched_row["信箱"])
        text_desc.config(state="normal")
        text_desc.delete("1.0", tk.END)
        text_desc.insert(tk.END, matched_row["問題描述"])
        text_desc.config(state="disabled")

    def submit_solution():
        nonlocal matched_row, df_main
        if matched_row is None:
            messagebox.showerror("錯誤", "請先查詢報修單 ID")
            return

        solution = text_solution.get("1.0", tk.END).strip()
        if not solution:
            messagebox.showerror("錯誤", "請輸入處理方式")
            return

        try:
            # 重新下載最新主檔，避免衝突
            df_main = download_excel_from_nas(CONFIG["main_excel"])

            index = df_main[df_main["ID"] == matched_row["ID"]].index
            if not index.empty:
                # 確保欄位存在，且型別是字串
                if "處理方式" not in df_main.columns:
                    df_main["處理方式"] = ""
                if "處理時間" not in df_main.columns:
                    df_main["處理時間"] = ""

                df_main["處理方式"] = df_main["處理方式"].astype(str)
                df_main["處理時間"] = df_main["處理時間"].astype(str)

                # 更新資料
                df_main.at[index[0], "處理方式"] = solution
                df_main.at[index[0], "處理時間"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                upload_excel_to_nas(df_main, CONFIG["main_excel"])

            else:
                messagebox.showerror("錯誤", "無法在主檔中找到該報修單 ID。")
                return
        except Exception as e:
            messagebox.showerror("錯誤", f"無法更新主檔：{e}")
            return

        try:
            send_email(
                matched_row["信箱"],
                matched_row["工號"],
                matched_row["問題描述"],
                solution,
            )
        except Exception as e:
            print(f"Email 發送失敗：{e}")

        messagebox.showinfo("完成", "處理結果已寫入主檔並通知客戶")

        # 清空欄位
        entry_date.delete(0, tk.END)
        entry_number.delete(0, tk.END)
        label_staff.config(text="")
        label_email.config(text="")
        text_desc.config(state="normal")
        text_desc.delete("1.0", tk.END)
        text_desc.config(state="disabled")
        text_solution.delete("1.0", tk.END)
        matched_row = None

    matched_row = None
    df_main = None

    window = tk.Tk()
    window.title("工程師報修處理系統(揚發)")
    window.geometry("500x650")

    tk.Label(window, text="工程師報修處理系統(揚發)", font=("Arial", 16, "bold")).pack(
        pady=10
    )

    tk.Label(window, text="輸入報修單 ID：", font=("Arial", 12)).pack(pady=10)

    frame_id = tk.Frame(window)
    frame_id.pack()

    tk.Label(frame_id, text="R", font=("Arial", 12)).grid(row=0, column=0)
    entry_date = tk.Entry(frame_id, font=("Arial", 12), width=10)
    entry_date.grid(row=0, column=1)

    tk.Label(frame_id, text=" - ", font=("Arial", 12)).grid(row=0, column=2)
    entry_number = tk.Entry(frame_id, font=("Arial", 12), width=5)
    entry_number.grid(row=0, column=3)

    tk.Button(window, text="搜尋", command=search_report, font=("Arial", 12)).pack(
        pady=5
    )

    frame_info = tk.Frame(window)
    frame_info.pack(pady=10)
    tk.Label(frame_info, text="工號：", font=("Arial", 12)).grid(
        row=0, column=0, sticky="e"
    )
    label_staff = tk.Label(frame_info, text="", font=("Arial", 12))
    label_staff.grid(row=0, column=1, sticky="w")
    tk.Label(frame_info, text="信箱：", font=("Arial", 12)).grid(
        row=1, column=0, sticky="e"
    )
    label_email = tk.Label(frame_info, text="", font=("Arial", 12))
    label_email.grid(row=1, column=1, sticky="w")

    tk.Label(window, text="問題描述：", font=("Arial", 12)).pack()
    text_desc = tk.Text(window, width=40, height=5, font=("Arial", 12), state="normal")
    text_desc.pack(pady=5)

    tk.Label(window, text="處理方式：", font=("Arial", 12)).pack()
    text_solution = tk.Text(window, width=40, height=7, font=("Arial", 12))
    text_solution.pack(pady=5)

    tk.Button(
        window, text="送出處理結果", font=("Arial", 12, "bold"), command=submit_solution
    ).pack(pady=10)

    footer_label = tk.Label(
        window, text="Copyright by ZY-Info V1.1", font=("Arial", 12)
    )
    footer_label.pack(side=tk.BOTTOM, anchor="e", padx=10, pady=10)

    window.mainloop()


if __name__ == "__main__":
    engineer_ui_nas()
