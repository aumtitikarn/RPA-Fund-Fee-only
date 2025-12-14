#!/usr/local/bin/python3
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox
from mergepdf import MergePDFPage
from doal import DaolPage
from scbam import SCBExtractorPage
from eastspring import EastspringPage
from assetfund import AssetFundPage

root = ttk.Window(themename="litera")
root.title("📘 Fund Fee only by NongAumzaap")
root.geometry("900x600")
root.resizable(False, False)

style = ttk.Style()
style.configure(".", font=("Kanit", 10))

# -------------------- NAVIGATION BAR --------------------
nav_frame = ttk.Frame(root, bootstyle="dark")
nav_frame.pack(fill="x")

pages = {}
current_page = None

def switch_page(page_name):
    global current_page
    if current_page:
        pages[current_page].pack_forget()
    pages[page_name].pack(fill="both", expand=True, padx=20, pady=20)
    current_page = page_name

def nav_button(parent, text, page):
    btn = ttk.Button(parent, text=text, bootstyle="secondary-outline",
                     command=lambda: switch_page(page))
    btn.pack(side="left", padx=5, pady=5)

ttk.Label(nav_frame, text="📘 Fund Fee only", font=("Kanit Semibold", 13),
          foreground="white", background="#343a40").pack(side="left", padx=10, pady=6)

nav_button(nav_frame, "Home", "home")
nav_button(nav_frame, "Merge PDF", "merge")
nav_button(nav_frame, "DAOL Extractor", "daol")
nav_button(nav_frame, "SCBAM", "scbam")
nav_button(nav_frame, "Eastspring", "eastspring")
nav_button(nav_frame, "Asset Fund", "assetfund")

# -------------------- PAGE: HOME --------------------
home = ttk.Frame(root)
pages["home"] = home

ttk.Label(home, text="🎯 Welcome to Document Tools Suite", font=("Kanit Semibold", 20)).pack(pady=40)
ttk.Label(home, text="รวมฟังก์ชันจัดการไฟล์ PDF และระบบสกัดข้อมูลกองทุน DAOL ไว้ในโปรแกรมเดียว", font=("Kanit", 12)).pack(pady=5)
ttk.Label(home, text="เลือกเมนูด้านบนเพื่อเริ่มใช้งาน", font=("Kanit", 11, "italic"), foreground="#6c757d").pack(pady=20)

# -------------------- PAGE: IMPORTED --------------------
pages["merge"] = MergePDFPage(root)
pages["daol"] = DaolPage(root)
pages["scbam"] = SCBExtractorPage(root)
pages["eastspring"] = EastspringPage(root)
pages["assetfund"] = AssetFundPage(root)

# -------------------- INITIAL PAGE --------------------
switch_page("home")
root.mainloop()
