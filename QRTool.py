import cv2
import tkinter as tk
from tkinter import filedialog, messagebox
import qrcode
from PIL import Image, ImageTk
import pyperclip
import pandas as pd
import openpyxl
import os
import numpy as np
from PIL import ImageGrab
import ctypes
from screeninfo import get_monitors

def get_virtual_screen_bounds():
    monitors = get_monitors()
    x_coords = [m.x for m in monitors]
    y_coords = [m.y for m in monitors]
    widths  = [m.x + m.width for m in monitors]
    heights = [m.y + m.height for m in monitors]
    return min(x_coords), min(y_coords), max(widths), max(heights)

def get_scaling_factor():
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    hdc = user32.GetDC(0)
    logical_dpi = gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
    user32.ReleaseDC(0, hdc)
    return logical_dpi / 96  # 96 是 100% 時的 DPI

def screenshot_and_detect_qrcode():
    selection = {'x1': 0, 'y1': 0, 'x2': 0, 'y2': 0}
    x0, y0, x1, y1 = get_virtual_screen_bounds()
    screen_width = x1 - x0
    screen_height = y1 - y0

    def on_mouse_down(event):
        selection['x1'] = event.x_root
        selection['y1'] = event.y_root

    def on_mouse_drag(event):
        canvas.delete("selection")
        canvas.create_rectangle(
            selection['x1'] - x0,
            selection['y1'] - y0,
            event.x_root - x0,
            event.y_root - y0,
            outline='red',
            width=2,
            tag="selection"
        )

    def on_mouse_up(event):
        selection['x2'] = event.x_root
        selection['y2'] = event.y_root
        top.destroy()

    top = tk.Toplevel(root)
    top.geometry(f"{screen_width}x{screen_height}+{x0}+{y0}")
    top.attributes("-alpha", 0.3)
    top.attributes("-topmost", True)
    top.overrideredirect(True)
    top.config(bg='black')

    canvas = tk.Canvas(top, cursor="cross", bg='black', highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    canvas.bind("<ButtonPress-1>", on_mouse_down)
    canvas.bind("<B1-Motion>", on_mouse_drag)
    canvas.bind("<ButtonRelease-1>", on_mouse_up)

    top.wait_window()

    x1 = min(selection['x1'], selection['x2'])
    y1 = min(selection['y1'], selection['y2'])
    x2 = max(selection['x1'], selection['x2'])
    y2 = max(selection['y1'], selection['y2'])

    if x2 - x1 == 0 or y2 - y1 == 0:
        messagebox.showinfo("提示", "沒有選取任何區域。")
        return

    screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
    cv_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

    detector = cv2.QRCodeDetector()
    retval, decoded_info, points, _ = detector.detectAndDecodeMulti(cv_image)

    if not retval or not decoded_info:
        messagebox.showinfo("結果", "選取區域中未偵測到 QR Code。")
        return

    def copy_to_clipboard():
        pyperclip.copy(retval)
        messagebox.showinfo("已複製", "QR Code 內容已複製到剪貼簿！")

    result_window = tk.Toplevel(root)
    result_window.title("掃描結果 - 多個 QR Code")
    result_window.geometry("1000x200")

    text_box = tk.Text(result_window, wrap="word", font=("Arial", 12))
    text_box.pack(expand=True, fill="both", padx=10, pady=10)

    for idx, qr_data in enumerate(decoded_info):
        if qr_data:
            text_box.insert("end", f"[{idx + 1}] {qr_data}\n")

    text_box.config(state="disabled")

    def copy_all_to_clipboard():
        pyperclip.copy("\n".join(filter(None, decoded_info)))
        messagebox.showinfo("已複製", "所有 QR Code 內容已複製到剪貼簿！")

    copy_button = tk.Button(result_window, text="複製全部內容", command=copy_all_to_clipboard)
    copy_button.pack(pady=5)

def read_qr_code(image_path):
    try:
        with open(image_path, "rb") as f:
            image_bytes = np.asarray(bytearray(f.read()), dtype=np.uint8)
            image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
    except Exception as e:
        messagebox.showerror("錯誤", f"無法讀取圖片：{e}")
        return

    detector = cv2.QRCodeDetector()
    data, bbox, _ = detector.detectAndDecode(image)

    if not data:
        messagebox.showinfo("結果", "未偵測到 QR Code。")
        return

    def copy_to_clipboard():
        pyperclip.copy(data)
        messagebox.showinfo("已複製", "QR Code 內容已複製到剪貼簿！")

    result_window = tk.Toplevel(root)
    result_window.title("QR Code 內容")
    result_window.geometry("400x300")

    text_box = tk.Text(result_window, wrap="word", font=("Arial", 12))
    text_box.insert("1.0", data)
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

screen_button = tk.Button(root, text="選取螢幕範圍掃描 QR Code", font=("Arial", 12), command=screenshot_and_detect_qrcode)
screen_button.pack(pady=10)

generate_exfile_button = tk.Button(root,text="建立範例Excel檔",font=("Arial",12),command=generate_example_excel)
generate_exfile_button.pack(pady=10)

root.mainloop()
