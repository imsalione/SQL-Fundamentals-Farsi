import tkinter as tk
from tkinter import messagebox

def show_message():
    messagebox.showwarning("پیام", "سلام صالح!")

# ایجاد پنجره اصلی
root = tk.Tk()
root.title("نمایش پیام")

# تنظیم سایز پنجره
window_width = 600
window_height = 400
root.geometry(f"{window_width}x{window_height}")

# بدست آوردن ابعاد صفحه نمایش
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# محاسبه مختصات x و y برای نمایش پنجره در وسط صفحه
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

# تنظیم مختصات پنجره
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# ایجاد دکمه و افزودن به پنجره
button = tk.Button(root, text="کلیک کن!", command=show_message)
button.pack(pady=40)

# اجرای حلقه اصلی برنامه
root.mainloop()