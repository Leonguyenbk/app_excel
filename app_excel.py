import re
import unicodedata
import sys
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =========================
# Regex & Helpers (từ script của bạn, giữ nguyên + các fix đã trao đổi)
# =========================
RE_ID = re.compile(r'\b(\d[\d\s\-]{7,20}\d)\b')
RE_DATE = r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})'

HONORIFIC = re.compile(
    r'^\s*(?:(?:và|va|&)\s*)?'
    r'(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em|mr|mrs|ms|miss)'
    r'\s*[:\.-]?\s*',
    flags=re.IGNORECASE
)

def normalize_id(token: str) -> str:
    digits = re.sub(r'\D', '', token)
    # chấp nhận CMND 9 số & CCCD 12 số; nếu muốn chỉ 12 số thì đổi điều kiện dưới
    return digits if 9 <= len(digits) <= 12 else ''

def vn_fold(s: str) -> str:
    """Hạ chữ + bỏ dấu tiếng Việt để so khớp không phân biệt dấu."""
    if not isinstance(s, str):
        return ''
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = s.replace('Đ', 'D').replace('đ', 'd')
    return s.lower()

def is_address_line(s: str) -> bool:
    return bool(re.search(r'^\s*địa\s*chỉ', s or '', flags=re.IGNORECASE))

def is_household_header(s: str) -> bool:
    """Dòng bắt đầu một 'bìa đỏ' mới: 'Hộ ông:', 'Hộ bà:', 'Hộ ...' """
    return bool(re.search(r'^\s*hộ\s*(?:ông|bà)?\b', s or '', flags=re.IGNORECASE))

def is_person_like(s: str) -> bool:
    """Xem dòng có phải thông tin 1 người không:
    - Có CMND/CCCD/CMT, HOẶC
    - Có 'Sinh năm ...' / 'Sinh ngày ...'
    """
    if not isinstance(s, str):
        return False
    return bool(
        re.search(r'(CMND|CCCD|CMT)\b', s, flags=re.IGNORECASE) or
        re.search(r'\bSinh\s+(?:năm|ngày)\b', s, flags=re.IGNORECASE)
    )

def is_plain_honorific_header(s: str) -> bool:
    """
    Dòng mở đầu 1 bìa mới nếu:
    - Bắt đầu bằng Ông/Bà/Anh/Chị/... (không có 'và|va|&' phía trước)
    - Và có dữ liệu người (CMND/CCCD hoặc 'Sinh ...')
    """
    if not isinstance(s, str):
        return False
    s = s.strip()
    if not s:
        return False
    if re.search(r'^(?:và|va|&)\b', s, flags=re.IGNORECASE):
        return False
    if not re.search(r'^(ông|bà|anh|chị|cô|chú|bác|em)\b', s, flags=re.IGNORECASE):
        return False
    return is_person_like(s)

def parse_person(text: str):
    """Tách thông tin 1 người từ 1 dòng (tên, năm sinh/ngày sinh, số giấy tờ, ngày cấp)."""
    if not isinstance(text, str) or not text.strip():
        return None
    s = text.strip()

    # rửa chuỗi nhẹ để tránh gạch dưới/khoảng trắng đặc biệt
    s = re.sub(r'[_\u00A0]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()

    if is_address_line(s):
        return None
    if not is_person_like(s):
        return None

    # Tên: không bắt buộc có dấu phẩy; kết thúc khi gặp "," | "Sinh" | "CMND/CCCD/CMT" | hết dòng
    m_name = re.search(
        r'^(?:\s*(?:(?:và|va|&)\s*)?'
        r'(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em)\s*[:\.-]?\s*)?'
        r'([A-ZÀ-ỸĐ][^,\d]+?)'
        r'\s*(?=,|\bSinh\b|\bCMND\b|\bCCCD\b|\bCMT\b|$)',
        s, flags=re.IGNORECASE
    )

    name = m_name.group(1).strip() if m_name else ''
    name = HONORIFIC.sub('', name).strip(' ,.-_')

    # Năm sinh | Ngày sinh
    birth_year, dob = '', ''
    m_year = re.search(r'Sinh\s*năm\s*:?[\s]*(\d{4})', s, flags=re.IGNORECASE)
    if m_year:
        birth_year = m_year.group(1)
    else:
        m_dob = re.search(r'Sinh\s*ngày\s*:?[\s]*' + RE_DATE, s, flags=re.IGNORECASE)
        if m_dob:
            dob = m_dob.group(1)

    # Số giấy tờ (ưu tiên có nhãn)
    id_number = ''
    m_id_labeled = re.search(
        r'(?:CMND|CCCD|CMT)\s*(?:số|so)?\s*[: ]*\s*([0-9\-\s]{7,20}\d)',
        s, flags=re.IGNORECASE
    )
    if m_id_labeled:
        id_number = normalize_id(m_id_labeled.group(1))
    if not id_number:
        for tok in RE_ID.findall(s):
            n = normalize_id(tok)
            if n:
                id_number = n
                break

    # Ngày cấp
    issue_date = ''

    # 1) Thử khớp trực tiếp có dấu trước (nhanh nhất)
    m_issue = re.search(
        r'(?:(?:c[âa]p)\s+ng[aàá]y|ng[aàá]y\s+(?:c[âa]p))\s*:?\s*' + RE_DATE,
        s, flags=re.IGNORECASE
    )
    if m_issue:
        issue_date = m_issue.group(1)
    else:
        # 2) Dự phòng: bỏ dấu rồi tìm vị trí 'ngay cap' / 'cap ngay', sau đó lấy DATE phía sau
        sfold = vn_fold(s)
        m_pre = re.search(r'(?:ngay\s+cap|cap\s+ngay)\s*:?\s*', sfold)
        if m_pre:
            start_pos = m_pre.end()
            # lấy date đầu tiên ở bản gốc, bắt đầu sau vị trí tìm được (an toàn với dấu câu/khoảng trắng lạ)
            for m in re.finditer(RE_DATE, s):
                if m.start() >= start_pos - 2:  # nới nhẹ vài ký tự do khác biệt dấu
                    issue_date = m.group(1)
                    break

    if not name or (not id_number and not birth_year and not dob):
        return None

    return {
        'full_name': name,
        'birth_year': birth_year or dob,  # nếu có ngày sinh chi tiết thì để nguyên "dd/mm/yyyy"
        'id_number': id_number,
        'issue_date': issue_date
    }

def record_to_wide(record_people):
    """Chuyển 1 bìa đỏ (list người) thành 1 dict dạng wide."""
    row = {}
    for idx, p in enumerate(record_people, start=1):
        row[f'Tên chủ {idx}']       = p.get('full_name', '')
        row[f'Năm sinh {idx}']      = p.get('birth_year', '')
        row[f'Số giấy tờ {idx}']    = p.get('id_number', '')
        row[f'Ngày cấp {idx}']      = p.get('issue_date', '')
    return row

def group_records(lines):
    """Gom dữ liệu theo 'bìa đỏ'."""
    records = []
    current_people = []

    def flush():
        if current_people:
            records.append(list(current_people))  # copy
            current_people.clear()

    for raw in lines:
        s = (raw or '').strip()
        if not s:
            flush()
            continue

        # 1) Gặp 'Hộ ...' -> luôn mở bìa mới
        if is_household_header(s):
            flush()
            p = parse_person(s)
            if p:
                current_people.append(p)
            continue

        # 2) Gặp 'Ông/Bà ...' (không kèm 'và') -> cũng mở bìa mới
        if is_plain_honorific_header(s):
            flush()
            p = parse_person(s)
            if p:
                current_people.append(p)
            continue

        # 3) Bỏ qua địa chỉ
        if is_address_line(s):
            continue

        # 4) Các dòng còn lại -> thêm người vào record hiện tại (nếu khớp người)
        p = parse_person(s)
        if p:
            current_people.append(p)

    flush()  # kết sổ cuối
    return records

def run_extraction(input_path: Path, output_path: Path, sheet, col_index: int):
    df = pd.read_excel(input_path, sheet_name=sheet, dtype=str)
    lines = df.iloc[:, col_index].fillna('').tolist()

    records = group_records(lines)
    max_people = max((len(r) for r in records), default=0)

    wide_rows = [record_to_wide(rec) for rec in records]
    out_df = pd.DataFrame(wide_rows)

    # Bổ sung cột trống để đều cột
    for i in range(1, max_people + 1):
        for col in [f'Tên chủ {i}', f'Năm sinh {i}', f'Số giấy tờ {i}', f'Ngày cấp {i}']:
            if col not in out_df.columns:
                out_df[col] = ''

    ordered_cols = []
    for i in range(1, max_people + 1):
        ordered_cols += [f'Tên chủ {i}', f'Năm sinh {i}', f'Số giấy tờ {i}', f'Ngày cấp {i}']
    out_df = out_df.reindex(columns=ordered_cols)

    out_df.to_excel(output_path, index=False)

# =========================
# Tkinter GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tách thông tin đồng sử dụng từ Excel")
        self.geometry("680x320")
        self.resizable(False, False)

        self.in_path_var = tk.StringVar()
        self.out_path_var = tk.StringVar()
        self.col_var = tk.StringVar(value="A")
        self.sheet_var = tk.StringVar(value="0")  # tạm thời; sẽ thay bằng combobox

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 6}

        # Input file
        frm1 = ttk.Frame(self)
        frm1.pack(fill='x', **pad)
        ttk.Label(frm1, text="File Excel đầu vào (.xlsx):").pack(side='left')
        ttk.Entry(frm1, textvariable=self.in_path_var, width=60).pack(side='left', padx=8)
        ttk.Button(frm1, text="Chọn...", command=self.choose_input).pack(side='left')

        # Output file
        frm2 = ttk.Frame(self)
        frm2.pack(fill='x', **pad)
        ttk.Label(frm2, text="File Excel đầu ra:").pack(side='left')
        ttk.Entry(frm2, textvariable=self.out_path_var, width=60).pack(side='left', padx=8)
        ttk.Button(frm2, text="Lưu thành...", command=self.choose_output).pack(side='left')

        # Sheet & Column
        frm3 = ttk.Frame(self)
        frm3.pack(fill='x', **pad)

        ttk.Label(frm3, text="Sheet:").pack(side='left')
        self.sheet_cb = ttk.Combobox(frm3, textvariable=self.sheet_var, width=18, state="readonly", values=["0"])
        self.sheet_cb.pack(side='left', padx=8)

        ttk.Label(frm3, text="Cột nguồn (A=0, B=1, ...):").pack(side='left', padx=(20,0))
        self.col_entry = ttk.Entry(frm3, textvariable=self.col_var, width=8)
        self.col_entry.pack(side='left', padx=8)
        ttk.Label(frm3, text="(nhập số index: 0 cho cột A)").pack(side='left')

        # Run
        frm4 = ttk.Frame(self)
        frm4.pack(fill='x', **pad)
        self.run_btn = ttk.Button(frm4, text="Chạy", command=self.on_run)
        self.run_btn.pack(side='left')
        ttk.Button(frm4, text="Thoát", command=self.destroy).pack(side='right')

        # Log
        frm5 = ttk.Frame(self)
        frm5.pack(fill='both', expand=True, **pad)
        ttk.Label(frm5, text="Nhật ký:").pack(anchor='w')
        self.log = tk.Text(frm5, height=8)
        self.log.pack(fill='both', expand=True)

    def log_print(self, s: str):
        self.log.insert('end', s + "\n")
        self.log.see('end')
        self.update_idletasks()

    def choose_input(self):
        p = filedialog.askopenfilename(
            title="Chọn file Excel đầu vào",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not p:
            return
        self.in_path_var.set(p)
        # đọc sheet names
        try:
            xls = pd.ExcelFile(p)
            sheets = xls.sheet_names
            self.sheet_cb['values'] = sheets
            # mặc định chọn sheet đầu
            self.sheet_var.set(sheets[0] if sheets else "0")
            # gợi ý output
            out_guess = str(Path(p).with_name(f"{Path(p).stem}_coowners.xlsx"))
            if not self.out_path_var.get():
                self.out_path_var.set(out_guess)
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Không đọc được danh sách sheet:\n{e}")

    def choose_output(self):
        p = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file kết quả",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            self.out_path_var.set(p)

    def on_run(self):
        in_path = self.in_path_var.get().strip()
        out_path = self.out_path_var.get().strip()
        sheet = self.sheet_var.get().strip()
        col_text = self.col_var.get().strip()

        if not in_path:
            messagebox.showwarning("Thiếu file đầu vào", "Vui lòng chọn file Excel đầu vào.")
            return
        if not out_path:
            messagebox.showwarning("Thiếu file đầu ra", "Vui lòng chọn nơi lưu file kết quả.")
            return
        # sheet: nếu là số, chuyển sang int
        try:
            sheet_idx = int(sheet)
            sheet_arg = sheet_idx
        except:
            sheet_arg = sheet

        # cột: yêu cầu nhập số index (0=A, 1=B, ...)
        try:
            col_index = int(col_text)
        except:
            messagebox.showwarning("Cột nguồn không hợp lệ", "Vui lòng nhập số index: 0 cho cột A, 1 cho B, ...")
            return

        try:
            self.run_btn.config(state='disabled')
            self.log_print("Đang chạy trích xuất...")
            run_extraction(Path(in_path), Path(out_path), sheet_arg, col_index)
            self.log_print(f"Đã xuất: {out_path}")
            messagebox.showinfo("Hoàn tất", f"Đã xuất kết quả:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xử lý", str(e))
        finally:
            self.run_btn.config(state='normal')

if __name__ == "__main__":
    app = App()
    app.mainloop()
