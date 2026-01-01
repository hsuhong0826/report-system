#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import io
from io import BytesIO
import uuid
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from datetime import datetime
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import sys
import time
from PIL import Image, ImageTk
import pyautogui
from screeninfo import get_monitors
from ftplib import FTP
import mss
import mss.tools

# 配置參數
CONFIG = {
    "name": "揚發",
    "prefix": "R",
    "system_name": "揚發_資訊管家",
    "record_folder": r"C:\\贊耀報修紀錄_揚發",
    "excel_file": "ZY_MA_Recoder.xlsx",
    "recipient_emails": [
        "zanyao0925@gmail.com",
        "adrian@yang-fa.com.tw",
        "sheep255174@gmail.com",
        "Zanyao.Service@msa.hinet.net",
        "hanna@yang-fa.com.tw",
    ],
    "ftp_host": "192.168.0.215",
    "ftp_port": 8821,
    "ftp_username": "service",
    "ftp_password": "Zanyao0915",
    "ftp_target_folder": "/MA_recoder",
    "sender_email": "Zanyao.Service@msa.hinet.net",
    "smtp_password": "Zanyao0915$",
}

attachment_paths = []
preview_images = []

root = None
name_entry = None
office_var = None
staff_id_entry = None
email_entry = None
phone_entry = None
anydesk_entry = None
description_text = None
attachment_label = None

select_preview_label = None
capture_preview_label = None
attachment_preview_frame = None
screenshot_preview_frame = None


def clean_text(text):
    try:
        return text.encode("utf-8", errors="replace").decode("utf-8").strip()
    except Exception as e:
        raise ValueError(f"文字處理錯誤: {text}, 錯誤: {e}")


def add_image_preview(image_path, which="select"):
    if not os.path.exists(image_path):
        return

    img = Image.open(image_path)
    img.thumbnail((100, 100))
    tk_img = ImageTk.PhotoImage(img)
    preview_images.append(tk_img)

    if which == "select":
        select_preview_label.config(text="")
        lbl = tk.Label(attachment_preview_frame, image=tk_img)
        lbl.pack(side=tk.LEFT, padx=5, pady=5)
    else:
        capture_preview_label.config(text="")
        lbl = tk.Label(screenshot_preview_frame, image=tk_img)
        lbl.pack(side=tk.LEFT, padx=5, pady=5)


def reset_fields():
    global attachment_paths
    name_entry.delete(0, tk.END)
    staff_id_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    anydesk_entry.delete(0, tk.END)
    description_text.delete("1.0", tk.END)

    office_var.set("第一辦公室")

    attachment_paths.clear()
    attachment_label.config(text="未選擇附件", font=("標楷體", 10, "bold"))

    for widget in attachment_preview_frame.winfo_children():
        widget.destroy()

    for widget in screenshot_preview_frame.winfo_children():
        widget.destroy()

    select_preview_label.config(text="")
    capture_preview_label.config(text="")


def select_image():
    global attachment_paths
    file_path = filedialog.askopenfilename(
        title="選擇圖片文件",
        filetypes=[("圖片文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp"), ("所有文件", "*.*")],
    )
    if file_path:
        attachment_paths.append(file_path)
        attachment_label.config(
            text=f"附件數量: {len(attachment_paths)}", font=("標楷體", 12, "bold")
        )
        add_image_preview(file_path, which="select")


def capture_screenshot():
    root.iconify()

    monitors = get_monitors()

    min_x = min(m.x for m in monitors)
    min_y = min(m.y for m in monitors)
    max_x = max(m.x + m.width for m in monitors)
    max_y = max(m.y + m.height for m in monitors)

    total_width = max_x - min_x
    total_height = max_y - min_y

    overlay = tk.Toplevel()
    overlay.geometry(f"{total_width}x{total_height}+{min_x}+{min_y}")
    overlay.attributes("-alpha", 0.3)
    overlay.overrideredirect(True)
    overlay.attributes("-topmost", True)
    overlay.config(bg="black")
    overlay.config(cursor="tcross")

    selection_canvas = tk.Canvas(overlay, bg="black", highlightthickness=0)
    selection_canvas.pack(fill="both", expand=True)

    start_x, start_y = None, None
    rect_id = None

    def on_mouse_down(event):
        nonlocal start_x, start_y, rect_id
        start_x, start_y = event.x, event.y
        rect_id = selection_canvas.create_rectangle(
            start_x, start_y, start_x, start_y, outline="blue", width=5, fill="blue"
        )

    def on_mouse_move(event):
        nonlocal rect_id
        if rect_id:
            selection_canvas.coords(rect_id, start_x, start_y, event.x, event.y)

    def on_mouse_up(event):
        nonlocal rect_id
        if rect_id:
            x2, y2 = event.x, event.y
            x1 = min(start_x, x2)
            y1 = min(start_y, y2)
            x2 = max(start_x, x2)
            y2 = max(start_y, y2)
            width = x2 - x1
            height = y2 - y1

            overlay.destroy()

            real_x1 = x1 + min_x
            real_y1 = y1 + min_y

            if not os.path.exists(CONFIG["record_folder"]):
                os.makedirs(CONFIG["record_folder"])

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = os.path.join(
                CONFIG["record_folder"], f"screenshot_{timestamp}.png"
            )

            with mss.mss() as sct:
                monitor = {
                    "top": real_y1,
                    "left": real_x1,
                    "width": width,
                    "height": height,
                }
                screenshot = sct.grab(monitor)
                img = Image.frombytes("RGB", screenshot.size, screenshot.rgb)
                img.save(save_path)

            attachment_paths.append(save_path)
            attachment_label.config(
                text=f"附件數量: {len(attachment_paths)}", font=("標楷體", 12, "bold")
            )
            add_image_preview(save_path, which="capture")

            root.deiconify()
            messagebox.showinfo("成功", "已完成截圖並加入附件！")

    selection_canvas.bind("<ButtonPress-1>", on_mouse_down)
    selection_canvas.bind("<B1-Motion>", on_mouse_move)
    selection_canvas.bind("<ButtonRelease-1>", on_mouse_up)


def generate_new_report_id():
    today = datetime.now()
    today_str = today.strftime("%Y%m%d")
    default_id = f"{CONFIG['prefix']}{today_str}-001"

    try:
        ftp = FTP()
        ftp.connect(CONFIG["ftp_host"], CONFIG["ftp_port"])
        ftp.login(CONFIG["ftp_username"], CONFIG["ftp_password"])
        ftp.cwd(CONFIG["ftp_target_folder"])
        ftp.encoding = "utf-8"

        buffer = io.BytesIO()
        ftp.retrbinary(f"RETR {CONFIG['excel_file']}", buffer.write)
        buffer.seek(0)
        df = pd.read_excel(buffer)

        if "報修時間" not in df.columns:
            return default_id

        df["報修時間"] = pd.to_datetime(df["報修時間"], errors="coerce")
        today_count = (df["報修時間"].dt.date == today.date()).sum()

        return f"{CONFIG['prefix']}{today_str}-{today_count + 1:03d}"

    except:
        return default_id


def send_customer_email(customer_email):
    sender_email = CONFIG["sender_email"]
    subject = "報修成功通知"
    body = "親愛的客戶：\n\n感謝您提交報修申請，我們已收到您的需求\n將盡快安排維修服務！\n\n敬祝\n順安！\n\n贊耀資訊"

    msg = MIMEText(body, "plain", "utf-8")
    msg["From"] = sender_email
    msg["To"] = customer_email
    msg["Subject"] = subject

    try:
        with smtplib.SMTP("msr.hinet.net", 587, local_hostname="localhost") as server:
            server.starttls()
            server.login(sender_email, CONFIG["smtp_password"])
            server.sendmail(sender_email, customer_email, msg.as_string())
    except:
        pass


def generate_email_body(name, staff_id, email, phone, anydesk, description, office):
    return (
        f"姓名: {name}\n"
        f"辦公室: {office}\n"
        f"工號: {staff_id}\n"
        f"信箱: {email}\n"
        f"電話(分機): {phone}\n"
        f"Anydesk號碼: {anydesk}\n"
        f"問題描述: {description}"
    )


def send_email(
    name,
    staff_id,
    email,
    phone,
    anydesk,
    description,
    attachment_paths,
    report_id,
    office,
):
    sender_email = CONFIG["sender_email"]
    recipient_emails = CONFIG["recipient_emails"]
    recipient_email_str = ", ".join(recipient_emails)

    subject = f"{CONFIG['name']}_電腦報修單"
    body = (
        f"ID: {report_id}\n"
        f"姓名: {name}\n"
        f"辦公室: {office}\n"
        f"工號: {staff_id}\n"
        f"信箱: {email}\n"
        f"電話(分機): {phone}\n"
        f"Anydesk號碼: {anydesk}\n"
        f"問題描述: {description}"
    )

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email_str
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    for file in attachment_paths:
        with open(file, "rb") as img_file:
            img_part = MIMEBase("image", "jpeg")
            img_part.set_payload(img_file.read())
        encoders.encode_base64(img_part)
        img_part.add_header(
            "Content-Disposition", f"attachment; filename={os.path.basename(file)}"
        )
        msg.attach(img_part)

    try:
        with smtplib.SMTP("msr.hinet.net", 587, local_hostname="localhost") as server:
            server.starttls()
            server.login(sender_email, CONFIG["smtp_password"])
            server.sendmail(sender_email, recipient_emails, msg.as_string())
    except:
        pass


def save_report_to_excel(name, staff_id, email, phone, anydesk, description):
    folder_path = CONFIG["record_folder"]
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    filename = CONFIG["excel_file"]
    file_path = os.path.join(folder_path, filename)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    data = [
        {
            "姓名": name,
            "工號": staff_id,
            "信箱": email,
            "電話(分機)": phone,
            "Anydesk號碼": anydesk,
            "問題描述": description,
            "報修時間": timestamp,
        }
    ]
    df_new = pd.DataFrame(data)

    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new

    df_combined.to_excel(file_path, index=False)
    return os.path.abspath(file_path)


def upload_excel_to_ftp(single_record: dict):
    try:
        ftp = FTP()
        ftp.connect(CONFIG["ftp_host"], CONFIG["ftp_port"])
        ftp.login(CONFIG["ftp_username"], CONFIG["ftp_password"])
        ftp.encoding = "utf-8"
        ftp.cwd(CONFIG["ftp_target_folder"])

        filename = CONFIG["excel_file"]

        remote_buffer = io.BytesIO()
        try:
            ftp.retrbinary(f"RETR {filename}", remote_buffer.write)
            remote_buffer.seek(0)
            df_existing = pd.read_excel(remote_buffer)
        except:
            df_existing = None

        df_new = pd.DataFrame([single_record])

        if df_existing is not None:
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        output_buffer = io.BytesIO()
        df_combined.to_excel(output_buffer, index=False)
        output_buffer.seek(0)

        ftp.storbinary(f"STOR {filename}", output_buffer)
        ftp.quit()

    except Exception as e:
        raise RuntimeError(f"FTP 上傳失敗: {e}")


def submit_report():
    try:
        name = clean_text(name_entry.get() or "未填寫")
        office = clean_text(office_var.get())
        staff_id = clean_text(staff_id_entry.get() or "未填寫")
        email = clean_text(email_entry.get() or "未填寫")
        phone = clean_text(phone_entry.get() or "未填寫")
        anydesk = clean_text(anydesk_entry.get() or "未填寫")
        description = clean_text(
            description_text.get("1.0", tk.END).strip() or "未填寫"
        )

        if not email or email == "未填寫":
            messagebox.showerror("錯誤", "信箱為必填欄位")
            return

    except ValueError as e:
        messagebox.showerror("錯誤", f"清理文字發生問題: {e}")
        return

    report_id = generate_new_report_id()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    record = {
        "ID": report_id,
        "姓名": name,
        "辦公室": office,
        "工號": staff_id,
        "信箱": email,
        "電話(分機)": phone,
        "Anydesk號碼": anydesk,
        "問題描述": description,
        "報修時間": timestamp,
        "處理方式": "",
        "處理時間": "",
    }

    local_excel_path = save_report_to_excel(
        name, staff_id, email, phone, anydesk, description
    )

    record.pop("辦公室", None)

    try:
        upload_excel_to_ftp(record)
    except Exception as e:
        messagebox.showerror("錯誤", f"NAS 上傳失敗: {e}")

    send_email(
        name,
        staff_id,
        email,
        phone,
        anydesk,
        description,
        attachment_paths,
        report_id,
        office,
    )

    if email != "未填寫":
        send_customer_email(email)

    reset_fields()
    messagebox.showinfo("成功", "報修單提交成功！")


def main():
    global root
    global name_entry, staff_id_entry, email_entry, phone_entry, anydesk_entry
    global description_text, attachment_label
    global select_preview_label, capture_preview_label
    global attachment_preview_frame, screenshot_preview_frame
    global office_var

    root = tk.Tk()
    root.title(CONFIG["system_name"])

    root.resizable(False, False)
    root.minsize(600, 650)

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.maxsize(screen_width, screen_height)

    tk.Label(root, text=CONFIG["system_name"], font=("標楷體", 16, "bold")).grid(
        row=0, column=0, columnspan=2, pady=10
    )

    tk.Label(root, text="姓名：", font=("標楷體", 12)).grid(
        row=1, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    name_entry = tk.Entry(root, width=40, font=("標楷體", 12))
    name_entry.grid(row=1, column=1, padx=(10, 50), pady=5)

    tk.Label(root, text="辦公室：", font=("標楷體", 12)).grid(
        row=2, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )

    office_var = tk.StringVar()
    office_combobox = ttk.Combobox(
        root,
        textvariable=office_var,
        values=["第一辦公室", "第二辦公室", "第三辦公室", "倉庫", "其他"],
        font=("標楷體", 12),
        state="readonly",
        width=38,
    )
    office_combobox.grid(row=2, column=1, padx=(10, 50), pady=5, sticky=tk.W)
    office_combobox.current(0)

    tk.Label(root, text="工號：", font=("標楷體", 12)).grid(
        row=3, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    staff_id_entry = tk.Entry(root, width=40, font=("標楷體", 12))
    staff_id_entry.grid(row=3, column=1, padx=(10, 50), pady=5)

    tk.Label(root, text="信箱：", font=("標楷體", 12)).grid(
        row=4, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    email_entry = tk.Entry(root, width=40, font=("標楷體", 12))
    email_entry.grid(row=4, column=1, padx=(10, 50), pady=5)

    tk.Label(root, text="電話(分機)：", font=("標楷體", 12)).grid(
        row=5, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    phone_entry = tk.Entry(root, width=40, font=("標楷體", 12))
    phone_entry.grid(row=5, column=1, padx=(10, 50), pady=5)

    tk.Label(root, text="Anydesk號碼：", font=("標楷體", 12)).grid(
        row=6, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    anydesk_entry = tk.Entry(root, width=40, font=("標楷體", 12))
    anydesk_entry.grid(row=6, column=1, padx=(10, 50), pady=5)

    tk.Label(root, text="問題描述：", font=("標楷體", 12)).grid(
        row=7, column=0, padx=(50, 10), pady=5, sticky=tk.W
    )
    description_text = tk.Text(root, width=40, height=10, font=("標楷體", 12))
    description_text.grid(row=7, column=1, padx=(10, 50), pady=5)

    tk.Button(root, text="選擇附件", command=select_image, font=("標楷體", 12)).grid(
        row=8, column=0, pady=5, padx=(50, 10), sticky=tk.W
    )

    select_preview_label = tk.Label(root, text="無附件預覽", font=("標楷體", 12))
    select_preview_label.grid(row=8, column=1, padx=(10, 10), pady=5, sticky=tk.W)

    attachment_preview_frame = tk.Frame(root)
    attachment_preview_frame.grid(row=8, column=1, padx=0, pady=10, sticky=tk.W)

    tk.Button(
        root, text="畫面截圖", command=capture_screenshot, font=("標楷體", 12)
    ).grid(row=9, column=0, pady=5, padx=(50, 10), sticky=tk.W)

    capture_preview_label = tk.Label(root, text="無截圖預覽", font=("標楷體", 12))
    capture_preview_label.grid(row=9, column=1, padx=(10, 10), pady=5, sticky=tk.W)

    screenshot_preview_frame = tk.Frame(root)
    screenshot_preview_frame.grid(row=9, column=1, padx=0, pady=10, sticky=tk.W)

    attachment_label = tk.Label(
        root, text="未選擇附件", font=("標楷體", 12), anchor="w"
    )
    attachment_label.grid(row=11, column=0, padx=50, pady=5, sticky=tk.W)

    submit_button = tk.Button(
        root, text="提交報修", command=submit_report, font=("標楷體", 14, "bold")
    )
    submit_button.grid(row=12, column=0, columnspan=2, pady=(20, 20))

    footer_label = tk.Label(
        root, text="Copyright by ZY-Info V1.6", font=("標楷體", 11, "bold")
    )
    footer_label.grid(row=13, column=1, sticky=tk.E, padx=10, pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
