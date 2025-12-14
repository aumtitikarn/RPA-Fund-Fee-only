import os
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PyPDF2 import PdfMerger, PdfReader
import subprocess
import platform

class MergePDFPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(padding=20)

        self.merge_path_var = ttk.StringVar()
        self.password_var = ttk.StringVar()

        # ========== HEADER ==========
        ttk.Label(self, text="📄 Merge PDF Files",
                  font=("Kanit Semibold", 22),
                  bootstyle="info").pack(pady=(10, 5))
        ttk.Label(self, text="รวมไฟล์ PDF ทั้งหมดในโฟลเดอร์ให้เป็นไฟล์เดียวอย่างรวดเร็วและปลอดภัย",
                  font=("Kanit", 11), foreground="#c7d0d9").pack(pady=(0, 20))

        # ========== CARD FRAME ==========
        card = ttk.Frame(self, padding=25)
        card.pack(pady=10, padx=50, fill="x")

        # --- Folder Input ---
        folder_label = ttk.Label(card, text="📁 โฟลเดอร์ไฟล์ PDF:", font=("Kanit", 11))
        folder_label.grid(row=0, column=0, sticky="w", pady=5)

        folder_frame = ttk.Frame(card)
        folder_frame.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        card.columnconfigure(1, weight=1)

        ttk.Entry(folder_frame, textvariable=self.merge_path_var,
                  bootstyle="info", font=("Kanit", 10)).pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(folder_frame, text="Browse...",
                   bootstyle="secondary-outline", command=self.browse_folder).pack(side="right")

        # --- Password Input ---
        ttk.Label(card, text="🔐 รหัสผ่าน (ถ้ามี):", font=("Kanit", 11)).grid(row=1, column=0, sticky="w", pady=(15, 5))
        ttk.Entry(card, textvariable=self.password_var, show="*",
                  bootstyle="info", font=("Kanit", 10)).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(15, 5))

        # ========== PROGRESS ==========
        self.progress = ttk.Progressbar(self, orient="horizontal",
                                        mode="determinate", length=600,
                                        bootstyle="info-striped")
        self.progress.pack(pady=(20, 8))
        self.status_label = ttk.Label(self, text="พร้อมทำงาน", font=("Kanit", 10))
        self.status_label.pack(pady=(0, 15))

        # ========== BUTTONS ==========
        btn_frame = ttk.Frame(self)
        btn_frame.pack()

        self.merge_btn = ttk.Button(btn_frame, text="✨ รวมไฟล์ PDF ✨",
                                    bootstyle="success-outline",
                                    width=22, command=self.start_merge)
        self.merge_btn.pack(side="left", padx=10)

        self.open_btn = ttk.Button(btn_frame, text="เปิดโฟลเดอร์ 📂",
                                   bootstyle="secondary", width=18,
                                   command=self.open_folder, state="disabled")
        self.open_btn.pack(side="left", padx=10)

        ttk.Label(self, text="© 2025 NongAumzaap", font=("Kanit", 9),
                  foreground="#7c8a97").pack(side="bottom", pady=10)

    # ========== FUNCTION ==========
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.merge_path_var.set(folder_selected)

    def start_merge(self):
        threading.Thread(target=self.merge_pdfs, daemon=True).start()

    def merge_pdfs(self):
        folder_path = self.merge_path_var.get()
        password = self.password_var.get().strip()
        self.open_btn.configure(state="disabled")

        if not folder_path:
            messagebox.showwarning("กรุณาเลือกโฟลเดอร์", "กรุณาเลือกโฟลเดอร์ที่มีไฟล์ PDF")
            return

        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        pdf_files.sort()
        if not pdf_files:
            messagebox.showwarning("ไม่พบไฟล์ PDF", "โฟลเดอร์นี้ไม่มีไฟล์ PDF")
            return

        merger = PdfMerger()
        total = len(pdf_files)

        try:
            for i, filename in enumerate(pdf_files, start=1):
                filepath = os.path.join(folder_path, filename)
                reader = PdfReader(filepath)

                if reader.is_encrypted:
                    if not password:
                        messagebox.showwarning("ไฟล์ถูกเข้ารหัส",
                                               f"'{filename}' ต้องการรหัสผ่าน!\nกรุณาใส่รหัสผ่านแล้วลองใหม่.")
                        return
                    result = reader.decrypt(password)
                    if result == 0:
                        messagebox.showerror("รหัสผ่านไม่ถูกต้อง",
                                             f"ไม่สามารถปลดล็อก '{filename}' ได้ (รหัสไม่ถูกต้อง)")
                        return

                merger.append(reader)
                self.progress["value"] = (i / total) * 100
                self.status_label.config(text=f"กำลังรวมไฟล์ {i}/{total} ...")
                self.update_idletasks()

            output_path = os.path.join(folder_path, "merged.pdf")
            merger.write(output_path)
            merger.close()

            self.status_label.config(text="✅ รวมไฟล์เสร็จสิ้น!")
            messagebox.showinfo("สำเร็จ!",
                                f"รวมไฟล์เรียบร้อยแล้ว\n\nบันทึกไว้ที่:\n{output_path}")
            self.open_btn.configure(state="normal")
            self.output_path = output_path
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", str(e))
        finally:
            self.progress["value"] = 0

    def open_folder(self):
        if hasattr(self, "output_path"):
            folder = os.path.dirname(self.output_path)
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
