import cv2
from pyzbar.pyzbar import decode
import tkinter as tk
from tkinter import filedialog, messagebox
import qrcode
from PIL import Image, ImageTk
import pyperclip
import pandas as pd
import openpyxl
import os

def read_qr_code(image_path):
    image = cv2.imread(image_path)
    if image is None:
        messagebox.showerror("錯誤", "無法讀取圖片，請檢查路徑是否正確。")
        return
    
    decoded_objects = decode(image)
    if not decoded_objects:
        messagebox.showinfo("結果", "未偵測到 QR Code。")
        return
    
    qr_data_list = [obj.data.decode("utf-8") for obj in decoded_objects]
    qr_data_text = "\n".join(qr_data_list)
    
    def copy_to_clipboard():
        pyperclip.copy(qr_data_text)
        messagebox.showinfo("已複製", "QR Code 內容已複製到剪貼簿！")
    
    result_window = tk.Toplevel(root)
    result_window.title("QR Code 內容")
    result_window.geometry("400x300")
    
    text_box = tk.Text(result_window, wrap="word", font=("Arial", 12))
    text_box.insert("1.0", qr_data_text)
    text_box.config(state="disabled")
    text_box.pack(expand=True, fill="both", padx=10, pady=10)
    
    copy_button = tk.Button(result_window, text="複製內容", command=copy_to_clipboard)
    copy_button.pack(pady=5)

def select_image():
    file_path = filedialog.askopenfilename(title="選擇 QR Code 圖片", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
    if file_path:
        read_qr_code(file_path)

def generate_qrcode():

    url = url_entry.get()
    if not url.strip():
        messagebox.showerror("錯誤", "URL 不能為空！")
        return

    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        qr_img.thumbnail((200, 200))
        qr_img_tk = ImageTk.PhotoImage(qr_img)
        qr_label.config(image=qr_img_tk)
        qr_label.image = qr_img_tk
        qr_label.qr_img = qr_img
    except Exception as e:
        messagebox.showerror("錯誤", f"無法生成 QR Code: {e}")

def generate_muti_qr_code(url, filename):
    """生成 QR Code 並儲存為圖片."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.thumbnail((200, 200))
    qr_label.qr_img = img
    img.save(filename)
    return filename  # 回傳檔案名稱

def save_qrcode():
    if hasattr(qr_label, "qr_img"):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("All Files", "*.*")],
        )
        if file_path:
            try:
                qr_label.qr_img.save(file_path)
                messagebox.showinfo("成功", f"QR Code 已儲存至 {file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法儲存 QR Code: {e}")
    else:
        messagebox.showwarning("警告", "請先生成 QR Code！")

def generate_example_excel():
    """生成範例 Excel 檔案."""
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not filename:
        return

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['QRCODE圖檔檔名', '網址'])
        ws.append(['example1', 'www.example.com'])
        ws.append(['example2', 'www.google.com'])
        wb.save(filename)
        messagebox.showinfo("成功", f"範例 Excel 檔案已儲存至 {filename}")
    except Exception as e:
        messagebox.showerror("錯誤", f"生成範例 Excel 檔案時發生錯誤：{e}")

def generate_bulk_qr_codes():
    """從 CSV 或 Excel 檔案批次生成 QR Code."""
    filename = filedialog.askopenfilename(filetypes=[ ("Excel files", "*.xlsx;*.xls"),("CSV files", "*.csv")])
    if not filename:
        return

    try:
        if filename.endswith('.csv'):
            df = pd.read_csv(filename)
        else:  # Excel
            df = pd.read_excel(filename)

        # 確保 DataFrame 包含 'QRCODE圖檔檔名' 和 '網址' 欄位
        if 'QRCODE圖檔檔名' not in df.columns or '網址' not in df.columns:
            messagebox.showerror("錯誤", "CSV/Excel 檔案必須包含 'QRCODE圖檔檔名' 和 '網址' 欄位。")
            return

        output_dir = filedialog.askdirectory()  # 讓使用者選擇輸出目錄
        if not output_dir:
            return  # 使用者取消選擇目錄

        success_count = 0
        error_count = 0

        for index, row in df.iterrows():
            qr_filename = row['QRCODE圖檔檔名']
            url = row['網址']

            # 確保檔案名稱包含副檔名
            if not qr_filename.lower().endswith('.png'):
                qr_filename += '.png'

            filepath = os.path.join(output_dir, qr_filename)

            try:
                generate_muti_qr_code(url, filepath)
                success_count += 1
            except Exception as e:
                error_count += 1
                print(f"生成 {qr_filename} 的 QR Code 時發生錯誤：{e}")

        message = f"成功生成 {success_count} 個 QR Code。"
        if error_count > 0:
            message += f"\n生成 {error_count} 個 QR Code 時發生錯誤，詳情請查看控制台輸出。"

        messagebox.showinfo("批次生成完成", message)

    except FileNotFoundError:
        messagebox.showerror("錯誤", "找不到指定的 CSV/Excel 檔案！")
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取或處理 CSV/Excel 檔案時發生錯誤：{e}")


root = tk.Tk()
root.title("QR Code 產生與掃描工具")
root.resizable(False, False)

title_label = tk.Label(root, text="QR Code 產生與掃描工具", font=("Arial", 16))
title_label.pack(pady=10)

url_label = tk.Label(root, text="輸入 URL:", font=("Arial", 12))
url_label.pack(pady=5)

url_entry = tk.Entry(root, font=("Arial", 12), width=40)
url_entry.pack(pady=5)

generate_button = tk.Button(root, text="生成 QR Code", font=("Arial", 12), command=generate_qrcode)
generate_button.pack(pady=10)

save_button = tk.Button(root, text="儲存 QR Code", font=("Arial", 12), command=save_qrcode)
save_button.pack(pady=10)

generate_muti_button = tk.Button(root, text="匯入EXCEL 生成 QR Code", font=("Arial", 12), command=generate_bulk_qr_codes)
generate_muti_button.pack(pady=10)

qr_label = tk.Label(root)
qr_label.pack(pady=20)

scan_button = tk.Button(root, text="選擇圖片掃描 QR Code", font=("Arial", 12), command=select_image)
scan_button.pack(pady=10)

generate_exfile_button = tk.Button(root,text="建立範例Excel檔",font=("Arial",12),command=generate_example_excel)
generate_exfile_button.pack(pady=10)

root.mainloop()
