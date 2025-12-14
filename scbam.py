import os
import re
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pdfplumber
import openpyxl
from openpyxl.styles import Font
from PIL import Image
import pytesseract
import platform

# 🔧 ตั้งค่า OCR สำหรับ cross-platform
if platform.system() == "Windows":
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
elif platform.system() == "Darwin":  # macOS
    # Check common macOS locations
    possible_paths = ["/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract", "/usr/bin/tesseract"]
    for path in possible_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break
    # If not found, pytesseract will try to use 'tesseract' from PATH

class SCBExtractorPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pdf_path = ttk.StringVar()
        self.password = ttk.StringVar()

        # ---------------- HEADER ----------------
        ttk.Label(self, text="🏦 SCB Fund Statement Extractor",
                  font=("Kanit Semibold", 20), bootstyle="info").pack(pady=(15, 5))
        ttk.Label(self, text="ดึงข้อมูลจาก Statement ของ SCB Asset Management (หลายหน้าในไฟล์เดียว)",
                  font=("Kanit", 11), foreground="#6c757d").pack(pady=(0, 20))

        # ---------------- INPUT CARD ----------------
        card = ttk.Frame(self, padding=20)
        card.pack(padx=40, pady=10, fill="x")

        ttk.Label(card, text="📄 เลือกไฟล์ PDF:", font=("Kanit", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(card, textvariable=self.pdf_path, width=60, bootstyle="info").grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(card, text="Browse...", bootstyle="secondary-outline",
                   command=self.browse_pdf).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(card, text="🔐 รหัสผ่าน PDF (ถ้ามี):", font=("Kanit", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(card, textvariable=self.password, width=60, show="*", bootstyle="info").grid(row=1, column=1, padx=5, pady=5)

        # ---------------- PROGRESS ----------------
        self.progress_bar = ttk.Progressbar(self, length=600, mode="determinate", bootstyle="info-striped")
        self.progress_bar.pack(pady=(25, 10))
        self.status_label = ttk.Label(self, text="พร้อมทำงาน", font=("Kanit", 10))
        self.status_label.pack(pady=5)

        # ---------------- BUTTON ----------------
        ttk.Button(self, text="เริ่มประมวลผล", bootstyle="success", width=20,
                   command=lambda: threading.Thread(target=self.run_extract, daemon=True).start()).pack(pady=10)
        ttk.Label(self, text="© 2025 NongAumzaap", font=("Kanit", 8), foreground="#888").pack(side="bottom", pady=5)

    # ---------------- FUNCTIONS ----------------
    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path.set(path)

    def run_extract(self):
        pdf_path = self.pdf_path.get()
        password = self.password.get().strip()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกไฟล์ PDF ก่อน")
            return

        try:
            with pdfplumber.open(pdf_path, password=password if password else None) as pdf:
                total_pages = len(pdf.pages)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "SCB Data"
                headers = ["ลำดับ", "เลขที่", "วันที่", "Unitholder No.", "ชื่อกองทุน", "Fee", "VAT", "total fee"]
                ws.append(headers)
                for col in range(1, len(headers)+1):
                    ws.cell(row=1, column=col).font = Font(bold=True)

                self.progress_bar["maximum"] = total_pages
                self.progress_bar["value"] = 0

                for i, page in enumerate(pdf.pages, start=1):
                    self.status_label.config(text=f"📑 กำลังอ่านหน้า {i}/{total_pages}")
                    self.update_idletasks()

                    text = page.extract_text() or ""

                    # OCR fallback
                    if not text or "Fund" not in text:
                        img = page.to_image(resolution=300).original

                        # OCR PREPROCESSING
                        import cv2, numpy as np
                        gray = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
                        gray = cv2.medianBlur(gray, 3)
                        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

                        custom_config = r"--oem 3 --psm 6 -c preserve_interword_spaces=1"

                        text = pytesseract.image_to_string(
                            thresh,
                            lang="eng+tha",
                            config=custom_config
                        )

                    # 🟣 PRINT RAW TEXT ก่อน extract
                    print("\n\n================ RAW TEXT PAGE", i, "================")
                    print(text)
                    print("====================================================\n\n")

                    # ส่งเข้า extract function
                    data = self.extract_info(text)

                    ws.append([
                        i,
                        data.get("เลขที่", ""),
                        data.get("วันที่", ""),
                        data.get("Unitholder No.", ""),
                        data.get("ชื่อกองทุน", ""),
                        data.get("Fee", ""),
                        data.get("VAT", ""),
                        data.get("total fee", "")
                    ])
                    self.progress_bar["value"] = i
                    self.update_idletasks()

                # ✅ ใช้ชื่อไฟล์ตรงตามที่ต้องการ
                output_path = os.path.join(os.path.dirname(pdf_path), "TaxInvoiceSCBAM.xlsx")
                wb.save(output_path)
                self.status_label.config(text="✅ เสร็จสิ้น")
                messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลเรียบร้อย:\n{output_path}")

        except Exception as e:
            if "incorrect password" in str(e).lower():
                messagebox.showerror("รหัสผ่านไม่ถูกต้อง", "ไม่สามารถเปิดไฟล์ได้เนื่องจากรหัสผ่านไม่ถูกต้อง ❌")
            else:
                messagebox.showerror("ข้อผิดพลาด", str(e))
            self.status_label.config(text="❌ เกิดข้อผิดพลาด")

    def extract_info(self, text: str):
        data = {
            "เลขที่": "",
            "วันที่": "",
            "Unitholder No.": "",
            "ชื่อกองทุน": "",
            "Fee": "",
            "VAT": "",
            "total fee": "",
        }

        # ---------- วันที่: ใช้ dd/mm/yyyy ตัวแรก ----------
        m = re.search(r"\b([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})\b", text)
        if m:
            data["วันที่"] = m.group(1)

        # เตรียม lines ไว้ใช้ต่อ
        lines = [l.strip() for l in text.splitlines() if l.strip()]

        # ---------- จับคู่ Client No. + Unitholder No. บรรทัดเดียวกัน ----------
        # รูปแบบประมาณ: 000-0-1872560-3 .... 0009910902
        pair = re.search(
            r"([0-9OQ]{3}-[0-9]-[0-9]{7}-[0-9]).{0,80}?([0-9]{6,12})",
            text,
            re.S  # ให้ . match ข้ามบรรทัดได้
        )
        if pair:
            unit_raw = pair.group(1)
            client_no = pair.group(2)

            # แก้ OCR เพี้ยน: Q / O -> 0
            unit_norm = (
                unit_raw
                .replace("O", "0")
                .replace("Q", "0")
            )

            data["เลขที่"] = client_no.strip()
            data["Unitholder No."] = unit_norm.strip()
        else:
            # Fallback แยกจับ ถ้าคู่ไม่เจอ
            # Unitholder No. = เลขแบบมีขีด 000-0-2540211-7
            m_unit = re.search(
                r"\b[0-9OQ]{3}-[0-9]-[0-9]{7}-[0-9]\b",
                text
            )
            if m_unit:
                unit_norm = (
                    m_unit.group(0)
                    .replace("O", "0")
                    .replace("Q", "0")
                )
                data["Unitholder No."] = unit_norm.strip()

            # เลขที่ = ตัวเลขล้วน 6–12 หลัก ที่อยู่หลังข้อความประมาณ Xxxxx/Xxxxx 0010106785
            m_client = re.search(
                r"[^0-9\s]{3,}/[^0-9\s]{3,}\s*([0-9]{6,12})",
                text
            )
            if m_client:
                data["เลขที่"] = m_client.group(1).strip()

        # ---------- Fund Code (ชื่อกองทุน) ----------
        # มองหา (SCBUSAA) แล้วตามด้วยบรรทัด Fund Name
        m = re.search(
            r"\(([A-Z0-9]{3,})\)\s*[\r\n]+\s*Fund\s*Name",
            text,
            re.IGNORECASE
        )
        if m:
            data["ชื่อกองทุน"] = m.group(1).strip()

        # ---------- หา block Fee (VAT Excluded) -> Fee (VAT Included หรือ Brokerage Fee) ----------
        # ---------- หา block Fee (VAT Excluded) -> Fee (VAT Included หรือ Brokerage Fee) ----------
        fee_start = None
        fee_end = None
        idx_broker = None

        for idx, line in enumerate(lines):
            # normalize เล็กน้อยให้ทน OCR เพี้ยน
            norm = (
                line
                .replace("Exctuded", "Excluded")
                .replace("Exduded", "Excluded")
            )

            # เริ่ม block จาก Fund Supervisor หรือ Fee (VAT Excluded)
            if fee_start is None and (
                re.search(r"Fund\s+Supervisor", norm, re.IGNORECASE)
                or re.search(r"(Fee\s*\()?(V|W)AT\s*Excluded", norm, re.IGNORECASE)
            ):
                fee_start = idx

            # จบ block ที่ Fee (VAT Included) ถ้ามี
            if re.search(r"(Fee\s*\()?(V|W)AT\s*Included", norm, re.IGNORECASE):
                fee_end = idx

            # เก็บตำแหน่ง Brokerage Fee ไว้ใช้เป็น fallback
            if idx_broker is None and re.search(r"Brokerage\s*Fee", norm, re.IGNORECASE):
                idx_broker = idx

        # ถ้าไม่เจอ Included แต่มี Brokerage Fee → ใช้มันเป็นจุดจบ block
        if fee_start is not None and fee_end is None and idx_broker is not None and idx_broker > fee_start:
            fee_end = idx_broker

        vat_val = None
        all_nums = []

        if fee_start is not None and fee_end is not None and fee_end >= fee_start:
            block_lines = lines[fee_start:fee_end + 1]

            for line in block_lines:
                nums_in_line = re.findall(r"[\d,]+\.\d{2}", line)

                has_vat_hint = ("(7%)" in line) or re.search(r"\bVAT\b", line, re.IGNORECASE)

                # ถ้ามีเลขปกติ (มีทศนิยม)
                if nums_in_line:
                    # ถ้าเป็นบรรทัด VAT ให้ใช้ตัวแรกเป็น VAT
                    if has_vat_hint and vat_val is None:
                        try:
                            vat_val = float(nums_in_line[0].replace(",", ""))
                        except ValueError:
                            pass

                    # เก็บทุกตัวเข้ารวม
                    for n in nums_in_line:
                        try:
                            all_nums.append(float(n.replace(",", "")))
                        except ValueError:
                            continue

                else:
                    # ไม่มีทศนิยมแต่เป็นบรรทัด VAT (เช่น 51688) → แปลงเป็น x/100
                    if has_vat_hint and vat_val is None:
                        m_int = re.search(r"\b(\d{3,7})\b", line)
                        if m_int:
                            try:
                                vat_val = int(m_int.group(1)) / 100.0
                                all_nums.append(vat_val)
                            except ValueError:
                                pass

        # ---------- ตรรกะเลือก Fee / VAT / total fee ----------
        fee_val = None
        total_val = None

        if all_nums:
            total_val = max(all_nums)

        # เคสปกติ: ถ้ามีทั้ง VAT และ Total → คำนวน Fee จากส่วนต่าง
        if vat_val is not None and total_val is not None:
            fee_val = round(total_val - vat_val, 2)

        # ถ้ายังไม่มี VAT แต่มี hint ว่ามีบรรทัด VAT และมีเลขอย่างน้อย 2 ตัว
        if (vat_val is None or fee_val is None) and all_nums:
            # สมมติว่า "ค่าธรรมเนียมก่อน VAT" คือเลขที่น้อยที่สุดใน block
            fee_guess = min(all_nums)
            total_guess = max(all_nums)
            if total_guess > fee_guess:
                fee_val = fee_guess
                vat_val = round(total_guess - fee_guess, 2)
                total_val = total_guess

        # Fallback เดิม: ถ้ามีเลข ≥ 3 ตัวและยังไม่ได้ set อะไร
        if (fee_val is None or total_val is None) and len(all_nums) >= 3:
            fee_val = all_nums[0] if fee_val is None else fee_val
            if vat_val is None:
                vat_val = all_nums[1]
            if total_val is None:
                total_val = all_nums[2]

        # ---------- format กลับเป็น string ----------
        def fmt(x):
            return f"{x:,.2f}"

        if fee_val is not None:
            data["Fee"] = fmt(fee_val)
        if vat_val is not None:
            data["VAT"] = fmt(vat_val)
        if total_val is not None:
            data["total fee"] = fmt(total_val)


        print(">> EXTRACTED DATA:", data)
        return data

