# -*- coding: utf-8 -*-
import re, unicodedata
from typing import List, Dict, Optional
from datetime import datetime
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ==== import converter của bạn (TCVN3/VNI -> Unicode) ====
# Đảm bảo file converter.py (bạn gửi) nằm cùng thư mục.
from converter import Converter
conv = Converter()

# ============ Regex & Chuẩn hoá ============
RE_ID   = re.compile(r'\b(\d[\d\s\-.]{7,24}\d)\b')
RE_DATE = r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})'

HONORIFIC = re.compile(
    r'^\s*(?:(?:và|va|&)\s*)?'
    r'(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em|mr|mrs|ms|miss)'
    r'\s*[:\.-]?\s*',
    flags=re.IGNORECASE
)

def normalize_id(token: str) -> str:
    digits = re.sub(r'\D', '', token or '')
    return digits if 9 <= len(digits) <= 12 else ''

def strip_accents(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.replace("Đ","D").replace("đ","d")

def normalize_text(s: str) -> str:
    if not isinstance(s, str): return ""
    s = s.replace('\u00A0', ' ')  # NBSP -> space
    s = re.sub(r'[_]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# --- chuyển TCVN3/VNI/… sang Unicode trước khi parse ---
def legacy_to_unicode(s: str) -> str:
    if not isinstance(s, str): return ""
    return conv.convert(s, target_charset="UNICODE", source_charset="TCVN3")

# chữ cái (kể cả có dấu), số; dùng để tách token vs. dấu câu/khoảng trắng
_TOKEN_SPLIT = re.compile(r'(\W+)', flags=re.UNICODE)

def fix_mixed_tcvn3_to_unicode(s: str) -> str:
    """
    Chuỗi có thể pha Unicode + TCVN3.
    - Tách theo token, với dấu câu/khoảng trắng được giữ lại.
    - Token nào detect là TCVN3 thì convert -> UNICODE.
    - Token đã là Unicode thì giữ nguyên.
    """
    if not isinstance(s, str):
        return ""
    parts = _TOKEN_SPLIT.split(s)   # giữ delimiter
    out = []
    for tok in parts:
        if not tok:
            continue
        # chỉ thử convert với "từ/cụm" có chữ (tránh convert dấu câu/space)
        if re.search(r'\w', tok, flags=re.UNICODE):
            cs = conv.detectCharset(tok)
            if cs and cs != "UNICODE" and cs == "TCVN3":
                try:
                    tok = conv.convert(tok, target_charset="UNICODE", source_charset="TCVN3")
                except Exception:
                    pass
        out.append(tok)
    return "".join(out)

def to_text(v) -> str:
    if v is None:
        return ""
    s = str(v)
    # 👇 vá pha trộn trước
    s = fix_mixed_tcvn3_to_unicode(s)
    # rồi chuẩn hoá khoảng trắng/nbps
    s = s.replace('\u00A0', ' ')
    s = re.sub(r'[_]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# ============ Nhận diện tiêu đề I/II ============
def is_section_header_i(s: str) -> bool:
    ss = strip_accents(normalize_text(s)).upper()
    return bool(re.match(r'^I\s*[-–—]?\s*NGUOI\s*SU\s*DUNG\s*DAT\b', ss))

def is_section_header_ii(s: str) -> bool:
    ss = strip_accents(normalize_text(s)).upper()
    return bool(re.match(r'^II\s*[-–—]?\s*THUA\s*DAT\b', ss))

def row_has_section_I(df, r) -> bool:
    if r < 0 or r >= len(df): return False
    row = " ".join(to_text(v) for v in df.iloc[r, :].tolist())
    return is_section_header_i(row)

def row_has_section_II(df, r) -> bool:
    if r < 0 or r >= len(df): return False
    row = " ".join(to_text(v) for v in df.iloc[r, :].tolist())
    return is_section_header_ii(row)

# ============ Địa chỉ ============
def is_address_line(s: str) -> bool:
    ss = strip_accents(normalize_text(s)).lower()
    return bool(re.match(r'^(?:dia\s*chi(?:\s*thuong\s*tru)?|d\s*\/\s*c|(?:noi\s*)?thuong\s*tru)\b', ss))

def extract_address(s: str) -> str:
    s_norm = normalize_text(s)
    m_vn = re.search(r'(?:địa\s*ch[ỉi](?:\s*thường\s*trú)?|đ\s*\/\s*c|(?:nơi\s*)?thường\s*trú)\s*[:\-]?\s*(.+)$',
                     s_norm, flags=re.IGNORECASE)
    if m_vn: return m_vn.group(1).strip(" .,_-")
    s_ascii = strip_accents(s_norm).lower()
    m = re.search(r'(?:dia\s*chi(?:\s*thuong\s*tru)?|d\s*\/\s*c|(?:noi\s*)?thuong\s*tru)\s*[:\-]?\s*(.+)$',
                  s_ascii, flags=re.IGNORECASE)
    return m.group(1).strip(" .,_-") if m else ""

def propagate_addresses(people: List[Dict]) -> None:
    addrs = [p.get('address','').strip() for p in people]
    uniq = {a for a in addrs if a}
    if len(uniq) == 1:
        addr = next(iter(uniq))
        for p in people:
            if not p.get('address'): p['address'] = addr

# ============ Người sử dụng ============
def is_person_like(s: str) -> bool:
    return bool(
        re.search(r'(CMND|CCCD|CMT)\b', s, flags=re.IGNORECASE) or
        re.search(r'\bSinh\s+(?:năm|ngày)\b', s, flags=re.IGNORECASE)
    )

def parse_person(text: str) -> Optional[Dict]:
    s = normalize_text(text)
    if not s or is_address_line(s) or not is_person_like(s): return None

    # Tên
    m_name = re.search(
        r'^(?:\s*(?:(?:và|va|&)\s*)?(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em)\s*[:\.-]?\s*)?'
        r'([A-ZÀ-ỸĐ][^,\d]+?)\s*(?=,|\bSinh\b|\bCMND\b|\bCCCD\b|\bCMT\b|$)', s, flags=re.IGNORECASE)
    name = m_name.group(1).strip() if m_name else ''
    name = HONORIFIC.sub('', name).strip(' ,.-_')

    # Năm/Ngày sinh
    birth_year, dob = '', ''
    m_year = re.search(r'Sinh\s*năm\s*:?[\s]*(\d{4})', s, flags=re.IGNORECASE)
    if m_year:
        birth_year = m_year.group(1)
    else:
        m_dob = re.search(r'Sinh\s*ngày\s*:?[\s]*' + RE_DATE, s, flags=re.IGNORECASE)
        if m_dob: dob = m_dob.group(1)

    # CMND/CCCD (cả "CCCD số" và "số CCCD")
    id_number = ''
    m_id_labeled = re.search(
        r'(?:(?:CMND|CCCD|CMT)\s*(?:s[ốo]|so)?|(?:s[ốo]|so)\s*(?:CMND|CCCD|CMT))'
        r'\s*[:\-\.,;]*\s*'
        r'([0-9][0-9\-\s\.]{7,24}\d)',
        s, flags=re.IGNORECASE
    )
    if m_id_labeled:
        id_number = normalize_id(m_id_labeled.group(1))
    if not id_number:
        for tok in RE_ID.findall(s):
            n = normalize_id(tok)
            if n: id_number = n; break

    # Ngày cấp: cả "cấp ngày" & "ngày cấp"
    issue_date = ''
    m_issue = re.search(
        r'(?:(?:cấp|cap)\s*[:,\-]?\s*ngày|ngày\s*[:,\-]?\s*(?:cấp|cap))\s*[:,\-]?\s*' + RE_DATE,
        s, flags=re.IGNORECASE
    )
    if m_issue: issue_date = m_issue.group(1)

    if not name or (not id_number and not birth_year and not dob): return None
    return {
        'full_name': name,
        'birth_year': birth_year or dob,
        'id_number': id_number,
        'issue_date': issue_date,
        'address': ''
    }

# ============ Thửa đất (đọc bảng nhiều cột) ============
def is_date_like(s: str) -> bool:
    s = s.strip()
    if not s: return False
    if re.match(r'^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$', s): return True
    if re.match(r'^\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}$', s): return True
    if s.isdigit() and 20000 <= int(s) <= 60000: return True  # serial excel
    return False

def normalize_date(s: str) -> str:
    s = s.strip()
    if not s: return ""
    if s.isdigit() and 20000 <= int(s) <= 60000:
        base = datetime(1899, 12, 30)
        dt = base + pd.to_timedelta(int(s), unit="D")
        return dt.strftime("%d/%m/%Y")
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return s

# ============ Gom bìa từ DataFrame ============
def group_records_from_df(df: pd.DataFrame) -> List[Dict]:
    n_rows, n_cols = df.shape
    r = 0
    rows: List[Dict] = []
    in_bia = False
    section = None
    current_people: List[Dict] = []
    current_parcels: List[Dict] = []

    def cell(rr, cc):
        if rr < 0 or rr >= n_rows or cc < 0 or cc >= n_cols: return ""
        return to_text(df.iat[rr, cc])

    def flush_bia():
        nonlocal rows, current_people, current_parcels, in_bia
        if not in_bia: return
        if not current_people and not current_parcels: 
            in_bia = False
            return
        propagate_addresses(current_people)
        p1 = current_people[0] if len(current_people) >= 1 else {}
        p2 = current_people[1] if len(current_people) >= 2 else {}
        for thua in (current_parcels or [dict()]):
            rows.append({
                'Ngày tháng vào sổ': thua.get('ngay',''),
                'Số thứ tự thửa đất': thua.get('thua',''),
                'Số thứ tự tờ bản đồ': thua.get('to',''),
                'Diện tích sử dụng riêng (m²)': thua.get('rieng',''),
                'Diện tích sử dụng chung (m²)': thua.get('chung',''),
                'Mục đích sử dụng': thua.get('mucdich',''),
                'Thời hạn sử dụng': thua.get('thoihan',''),
                'Nguồn gốc sử dụng': thua.get('nguongoc',''),
                'Số phát hành GCN': thua.get('soph',''),
                'Số vào sổ': thua.get('sovs',''),
                'Tên chủ 1': p1.get('full_name',''),
                'Năm sinh 1': p1.get('birth_year',''),
                'Số giấy tờ 1': p1.get('id_number',''),
                'Ngày cấp 1': p1.get('issue_date',''),
                'Địa chỉ 1': p1.get('address',''),
                'Tên chủ 2': p2.get('full_name',''),
                'Năm sinh 2': p2.get('birth_year',''),
                'Số giấy tờ 2': p2.get('id_number',''),
                'Ngày cấp 2': p2.get('issue_date',''),
                'Địa chỉ 2': p2.get('address',''),
            })
        current_people.clear()
        current_parcels.clear()
        in_bia = False

    while r < n_rows:
        # Bắt đầu bìa tại dòng có "I - NGƯỜI SỬ DỤNG ĐẤT" (ở bất kỳ ô nào của hàng)
        if row_has_section_I(df, r):
            flush_bia()
            in_bia = True
            section = 'I'
            r += 1
            continue

        if not in_bia:
            r += 1
            continue

        # Đến mục II -> đọc bảng nhiều cột
        if row_has_section_II(df, r):
            section = 'II'
            # tìm header trong 1-4 dòng tiếp theo
            header_r = None
            for k in range(1, 5):
                if r + k >= n_rows: break
                line = " | ".join(strip_accents(cell(r + k, c)).lower() for c in range(n_cols))
                if ("ngay" in line and "vao so" in line) and ("so thu tu" in line):
                    header_r = r + k
                    break
            if header_r is None:
                r += 1
                continue

            # map tên cột -> index
            col_map = {}
            for c in range(n_cols):
                h = strip_accents(cell(header_r, c)).lower()
                if 'ngay' in h and 'vao so' in h: col_map['ngay'] = c
                if 'so thu tu' in h and 'thua dat' in h: col_map['thua'] = c
                if 'so thu tu' in h and 'to ban do' in h: col_map['to'] = c
                if 'rieng' in h: col_map['rieng'] = c
                if 'chung' in h: col_map['chung'] = c
                if 'muc dich' in h: col_map['mucdich'] = c
                if 'thoi han' in h: col_map['thoihan'] = c
                if 'nguon goc' in h: col_map['nguongoc'] = c
                if 'so phat hanh' in h or 'gcn qsdd' in h: col_map['soph'] = c
                if 'so vao so' in h: col_map['sovs'] = c

            # đọc các hàng dữ liệu cho đến khi cột ngày không còn hợp lệ
            rr = header_r + 1
            while rr < n_rows:
                v_ngay_raw = cell(rr, col_map.get('ngay', 0))
                if not is_date_like(v_ngay_raw):
                    break
                ngay = normalize_date(v_ngay_raw)
                def g(key):
                    if key not in col_map: return ''
                    return cell(rr, col_map[key])

                thua = {
                    'ngay'   : ngay,
                    'thua'   : g('thua'),
                    'to'     : g('to'),
                    'rieng'  : g('rieng').replace(',', '.').replace('Không','').replace('không',''),
                    'chung'  : g('chung').replace(',', '.').replace('Không','').replace('không',''),
                    'mucdich': g('mucdich'),
                    'thoihan': g('thoihan'),
                    'nguongoc': g('nguongoc'),
                    'soph'   : g('soph').replace(' ', ''),
                    'sovs'   : g('sovs'),
                }
                current_parcels.append(thua)
                rr += 1

            r = rr
            continue

        # Đang trong mục I: đọc người & địa chỉ (gộp cả hàng cho chắc)
        if section == 'I':
            row_text = " ".join(cell(r, c) for c in range(n_cols))
            if is_address_line(row_text):
                if current_people:
                    addr = extract_address(row_text)
                    if addr and not current_people[-1].get('address'):
                        current_people[-1]['address'] = addr
                        propagate_addresses(current_people)
            else:
                p = parse_person(row_text)
                if p: current_people.append(p)
            r += 1
            continue

        r += 1

    # flush bìa cuối
    flush_bia()
    return rows

# ============ Run + GUI ============
def run_extraction(input_path: Path, output_path: Path, sheet, _col_index_unused: int):
    # ép openpyxl để đảm bảo Unicode
    df = pd.read_excel(input_path, sheet_name=sheet, dtype=object, engine="openpyxl")
    records = group_records_from_df(df)
    out_df = pd.DataFrame(records)

    ordered_cols = [
        'Ngày tháng vào sổ','Số thứ tự thửa đất','Số thứ tự tờ bản đồ',
        'Diện tích sử dụng riêng (m²)','Diện tích sử dụng chung (m²)',
        'Mục đích sử dụng','Thời hạn sử dụng','Nguồn gốc sử dụng',
        'Số phát hành GCN','Số vào sổ',
        'Tên chủ 1','Năm sinh 1','Số giấy tờ 1','Ngày cấp 1','Địa chỉ 1',
        'Tên chủ 2','Năm sinh 2','Số giấy tờ 2','Ngày cấp 2','Địa chỉ 2'
    ]
    for c in ordered_cols:
        if c not in out_df.columns: out_df[c] = ''
    out_df = out_df[ordered_cols]
    out_df.to_excel(output_path, index=False, engine="openpyxl")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Xuất THỬA ĐẤT + Chủ (I -> II) | Tự convert TCVN3/VNI -> Unicode")
        self.geometry("760x380"); self.resizable(False, False)
        self.in_path_var = tk.StringVar(); self.out_path_var = tk.StringVar()
        self.col_var = tk.StringVar(value="0")  # giữ lại cho tương thích, không dùng ở II
        self.sheet_var = tk.StringVar(value="0")
        self._build_ui()

    def _build_ui(self):
        pad = {'padx':10,'pady':6}
        f1=ttk.Frame(self); f1.pack(fill='x', **pad)
        ttk.Label(f1,text="File Excel đầu vào (.xlsx):").pack(side='left')
        ttk.Entry(f1,textvariable=self.in_path_var,width=60).pack(side='left',padx=8)
        ttk.Button(f1,text="Chọn...",command=self.choose_input).pack(side='left')

        f2=ttk.Frame(self); f2.pack(fill='x', **pad)
        ttk.Label(f2,text="File Excel đầu ra:").pack(side='left')
        ttk.Entry(f2,textvariable=self.out_path_var,width=60).pack(side='left',padx=8)
        ttk.Button(f2,text="Lưu thành...",command=self.choose_output).pack(side='left')

        f3=ttk.Frame(self); f3.pack(fill='x', **pad)
        ttk.Label(f3,text="Sheet:").pack(side='left')
        self.sheet_cb=ttk.Combobox(f3,textvariable=self.sheet_var,width=28,state="readonly",values=["0"])
        self.sheet_cb.pack(side='left',padx=8)

        ttk.Label(f3,text="(Tuỳ chọn) Cột nguồn cho mục I (0=A):").pack(side='left',padx=(20,0))
        ttk.Entry(f3,textvariable=self.col_var,width=6).pack(side='left',padx=8)

        f4=ttk.Frame(self); f4.pack(fill='x', **pad)
        self.run_btn=ttk.Button(f4,text="Chạy",command=self.on_run); self.run_btn.pack(side='left')
        ttk.Button(f4,text="Thoát",command=self.destroy).pack(side='right')

        f5=ttk.Frame(self); f5.pack(fill='both',expand=True, **pad)
        ttk.Label(f5,text="Nhật ký:").pack(anchor='w')
        self.log=tk.Text(f5,height=9); self.log.pack(fill='both',expand=True)

    def log_print(self,s:str):
        self.log.insert('end',s+"\n"); self.log.see('end'); self.update_idletasks()

    def choose_input(self):
        p=filedialog.askopenfilename(title="Chọn file Excel đầu vào",
            filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if not p: return
        self.in_path_var.set(p)
        try:
            xls=pd.ExcelFile(p, engine="openpyxl")
            sheets=xls.sheet_names
            self.sheet_cb['values']=sheets
            self.sheet_var.set(sheets[0] if sheets else "0")
            out_guess=str(Path(p).with_name(f"{Path(p).stem}_thuadat_owners.xlsx"))
            if not self.out_path_var.get(): self.out_path_var.set(out_guess)
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Không đọc được danh sách sheet:\n{e}")

    def choose_output(self):
        p=filedialog.asksaveasfilename(title="Chọn nơi lưu file kết quả",
            defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
        if p: self.out_path_var.set(p)

    def on_run(self):
        in_path=self.in_path_var.get().strip()
        out_path=self.out_path_var.get().strip()
        sheet=self.sheet_var.get().strip()
        col_text=self.col_var.get().strip()  # giữ để tương thích
        if not in_path: messagebox.showwarning("Thiếu file đầu vào","Chọn file Excel đầu vào."); return
        if not out_path: messagebox.showwarning("Thiếu file đầu ra","Chọn nơi lưu file kết quả."); return
        try:
            sheet_arg=int(sheet)
        except:
            sheet_arg=sheet
        try:
            _col_index=int(col_text) if col_text else 0
        except:
            messagebox.showwarning("Cột nguồn không hợp lệ","Nhập số index: 0 cho A, 1 cho B, ..."); return

        try:
            self.run_btn.config(state='disabled')
            self.log_print("Đang chạy trích xuất (tự chuyển TCVN3/VNI -> Unicode)...")
            run_extraction(Path(in_path), Path(out_path), sheet_arg, _col_index)
            self.log_print(f"Đã xuất: {out_path}")
            messagebox.showinfo("Hoàn tất", f"Đã xuất:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xử lý", str(e))
        finally:
            self.run_btn.config(state='normal')

if __name__ == "__main__":
    App().mainloop()
