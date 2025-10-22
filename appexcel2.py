import re, unicodedata
from typing import List, Dict, Optional
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ============ Helpers & Regex ============
RE_ID = re.compile(r'\b(\d[\d\s\-.]{7,24}\d)\b')
RE_DATE = r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})'

HONORIFIC = re.compile(
    r'^\s*(?:(?:và|va|&)\s*)?'
    r'(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em|mr|mrs|ms|miss)'
    r'\s*[:\.-]?\s*', flags=re.IGNORECASE
)

def normalize_id(token: str) -> str:
    digits = re.sub(r'\D', '', token)
    return digits if 9 <= len(digits) <= 12 else ''

def strip_accents(s: str) -> str:
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.replace("Đ","D").replace("đ","d")

def normalize_text(s: str) -> str:
    if not isinstance(s, str): return ""
    s = s.replace('\u00A0', ' ')
    s = re.sub(r'[_]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def is_section_header_i(s: str) -> bool:
    s_ascii = strip_accents(normalize_text(s)).upper()
    return bool(re.match(r'^I\s*[-–—]?\s*NGUOI\s*SU\s*DUNG\s*DAT\b', s_ascii))

def is_section_header_ii(s: str) -> bool:
    s_ascii = strip_accents(normalize_text(s)).upper()
    return bool(re.match(r'^II\s*[-–—]?\s*THUA\s*DAT\b', s_ascii))

def is_address_line(s: str) -> bool:
    s_ascii = strip_accents(normalize_text(s)).lower()
    return bool(re.match(r'^(?:dia\s*chi(?:\s*thuong\s*tru)?|d\s*\/\s*c|(?:noi\s*)?thuong\s*tru)\b', s_ascii))

def extract_address(s: str) -> str:
    s_norm = normalize_text(s)
    m_vn = re.search(r'(?:địa\s*ch[ỉi](?:\s*thường\s*trú)?|đ\s*\/\s*c|(?:nơi\s*)?thường\s*trú)\s*[:\-]?\s*(.+)$',
                     s_norm, flags=re.IGNORECASE)
    if m_vn: return m_vn.group(1).strip(" .,_-")
    s_ascii = strip_accents(s_norm).lower()
    m = re.search(r'(?:dia\s*chi(?:\s*thuong\s*tru)?|d\s*\/\s*c|(?:noi\s*)?thuong\s*tru)\s*[:\-]?\s*(.+)$',
                  s_ascii, flags=re.IGNORECASE)
    return m.group(1).strip(" .,_-") if m else ""

def is_household_header(s: str) -> bool:
    return bool(re.search(r'^\s*hộ\s*(?:ông|bà)?\b', s or '', flags=re.IGNORECASE))

def is_person_like(s: str) -> bool:
    return bool(
        re.search(r'(CMND|CCCD|CMT)\b', s, flags=re.IGNORECASE) or
        re.search(r'\bSinh\s+(?:năm|ngày)\b', s, flags=re.IGNORECASE)
    )

def is_plain_honorific_header(s: str) -> bool:
    s = (s or '').strip()
    if not s: return False
    if re.search(r'^(?:và|va|&)\b', s, flags=re.IGNORECASE): return False
    if not re.search(r'^(ông|bà|anh|chị|cô|chú|bác|em)\b', s, flags=re.IGNORECASE): return False
    return is_person_like(s)

# ============ Parse PERSON ============
def parse_person(text: str) -> Optional[Dict]:
    s = normalize_text(text)
    if not s or is_address_line(s) or not is_person_like(s): return None

    m_name = re.search(
        r'^(?:\s*(?:(?:và|va|&)\s*)?(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em)\s*[:\.-]?\s*)?'
        r'([A-ZÀ-ỸĐ][^,\d]+?)\s*(?=,|\bSinh\b|\bCMND\b|\bCCCD\b|\bCMT\b|$)', s, flags=re.IGNORECASE)
    name = m_name.group(1).strip() if m_name else ''
    name = HONORIFIC.sub('', name).strip(' ,.-_')

    birth_year, dob = '', ''
    m_year = re.search(r'Sinh\s*năm\s*:?[\s]*(\d{4})', s, flags=re.IGNORECASE)
    if m_year: birth_year = m_year.group(1)
    else:
        m_dob = re.search(r'Sinh\s*ngày\s*:?[\s]*' + RE_DATE, s, flags=re.IGNORECASE)
        if m_dob: dob = m_dob.group(1)

    id_number = ''
    m_id_labeled = re.search(
        r'(?:(?:CMND|CCCD|CMT)\s*(?:s[ốo]|so)?|(?:s[ốo]|so)\s*(?:CMND|CCCD|CMT))\s*[:\-\.,;]*\s*'
        r'([0-9][0-9\-\s\.]{7,24}\d)', s, flags=re.IGNORECASE)
    if m_id_labeled:
        id_number = normalize_id(m_id_labeled.group(1))
    if not id_number:
        for tok in RE_ID.findall(s):
            n = normalize_id(tok)
            if n: id_number = n; break

    issue_date = ''
    m_issue = re.search(
        r'(?:(?:cấp|cap)\s*[:,\-]?\s*ngày|ngày\s*[:,\-]?\s*(?:cấp|cap))\s*[:,\-]?\s*' + RE_DATE,
        s, flags=re.IGNORECASE
    )
    issue_date = m_issue.group(1) if m_issue else ''

    if not name or (not id_number and not birth_year and not dob): return None
    return {'full_name': name, 'birth_year': birth_year or dob, 'id_number': id_number,
            'issue_date': issue_date, 'address': ''}

# ============ Parse PARCEL ============
def _num(s: str) -> str:
    return s.replace(',', '.') if s else ''

def parse_parcel_line_v2(s: str) -> Optional[Dict]:
    t = normalize_text(s)
    if not t: return None
    m0 = re.match(r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})\b(.*)$', t)
    if not m0: return None
    ngay, rest = m0.group(1), m0.group(2).strip()

    m12 = re.match(r'[, ]*(\d+)[, ]+(\d+)(.*)$', rest)  # (số thửa) (số tờ)
    if m12:
        so_thua, so_to, rest2 = m12.group(1), m12.group(2), m12.group(3).strip()
    else:
        so_thua, so_to, rest2 = '', '', rest

    raw_ascii = strip_accents(rest2).lower()
    dt_rieng = ''
    dt_chung = ''
    m_r = re.search(r'rieng\s*([0-9]+(?:[.,][0-9]{1,3})?)', raw_ascii)
    if m_r: dt_rieng = _num(m_r.group(1))
    elif re.search(r'rieng\s*khong', raw_ascii): dt_rieng = ''

    m_c = re.search(r'chung\s*([0-9]+(?:[.,][0-9]{1,3})?)', raw_ascii)
    if m_c: dt_chung = _num(m_c.group(1))
    elif re.search(r'chung\s*khong', raw_ascii): dt_chung = ''

    m_md = re.search(r'\b(ONT|ODT|CLN|LUC|LUA|HNK|RSX|TMDV|DTS|NTS|BHK|QH|LNK|NKH|CQP|DGT|DGD|CSD|CAY)\b', rest2)
    muc_dich = m_md.group(1) if m_md else (re.search(r'\b[A-Z]{2,4}\b', rest2).group(0) if re.search(r'\b[A-Z]{2,4}\b', rest2) else '')

    m_th = re.search(r'(Lâu\s*dài|\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})', rest2, flags=re.IGNORECASE)
    thoi_han = m_th.group(1) if m_th else ''

    m_ng = re.search(r'\b([A-Z]{2,4}-[A-Z]{2,4})\b', rest2)  # DG-KTT / CNQ-CTT ...
    nguon_goc = m_ng.group(1) if m_ng else ''

    m_ph = re.search(r'\b([A-Z]{1,3}\s*\d{4,})\b', rest2)
    so_ph = m_ph.group(1).replace(' ', '') if m_ph else ''

    m_vs = re.findall(r'\b([A-Z]{1,3}\d{3,})\b', rest2)
    so_vao = m_vs[-1] if m_vs else ''

    return {
        'ngay_thang_vao_so': ngay,
        'so_thu_tu_thua_dat': so_thua,
        'so_thu_tu_to_ban_do': so_to,
        'dt_rieng_m2': dt_rieng,
        'dt_chung_m2': dt_chung,
        'muc_dich_su_dung': muc_dich,
        'thoi_han_su_dung': thoi_han,
        'nguon_goc_su_dung': nguon_goc,
        'so_phat_hanh_gcn': so_ph,
        'so_vao_so': so_vao,
        'thua_raw': t
    }

# ============ Propagate address ============
def propagate_addresses(people: List[Dict]) -> None:
    addrs = [p.get('address','').strip() for p in people]
    uniq = {a for a in addrs if a}
    if len(uniq) == 1:
        addr = next(iter(uniq))
        for p in people:
            if not p.get('address'): p['address'] = addr

# ============ Grouping (I bắt đầu bìa, II thuộc bìa) ============
def group_records_from_df(df: pd.DataFrame, col_index: int) -> list[dict]:
    """
    Bìa bắt đầu tại 'I - NGƯỜI SỬ DỤNG ĐẤT' trong cột col_index.
    'II - THỬA ĐẤT' là bảng nhiều cột ngay dưới, lấy theo tiêu đề cột.
    Trả về list hàng phẳng: thửa đất (trước) + Chủ 1/2 (sau).
    """
    # helper
    def cell(r, c):
        v = df.iat[r, c] if (0 <= r < len(df) and 0 <= c < df.shape[1]) else ''
        return normalize_text('' if pd.isna(v) else str(v))

    n_rows = len(df)
    r = 0
    rows = []

    current_people = []
    current_parcels = []

    def flush_bia():
        nonlocal rows, current_people, current_parcels
        if not current_people and not current_parcels:
            return
        propagate_addresses(current_people)
        p1 = current_people[0] if len(current_people) >= 1 else {}
        p2 = current_people[1] if len(current_people) >= 2 else {}
        for thua in (current_parcels or [dict()]):
            rows.append({
                'Ngày tháng vào sổ': thua.get('ngay_thang_vao_so',''),
                'Số thứ tự thửa đất': thua.get('so_thu_tu_thua_dat',''),
                'Số thứ tự tờ bản đồ': thua.get('so_thu_tu_to_ban_do',''),
                'Diện tích sử dụng riêng (m²)': thua.get('dt_rieng_m2',''),
                'Diện tích sử dụng chung (m²)': thua.get('dt_chung_m2',''),
                'Mục đích sử dụng': thua.get('muc_dich_su_dung',''),
                'Thời hạn sử dụng': thua.get('thoi_han_su_dung',''),
                'Nguồn gốc sử dụng': thua.get('nguon_goc_su_dung',''),
                'Số phát hành GCN': thua.get('so_phat_hanh_gcn',''),
                'Số vào sổ': thua.get('so_vao_so',''),
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

    while r < n_rows:
        s = cell(r, col_index)
        if not s:
            r += 1
            continue

        # Bắt đầu bìa mới
        if is_section_header_i(s):
            flush_bia()
            r += 1
            # thu người cho đến khi gặp II
            while r < n_rows:
                t = cell(r, col_index)
                if is_section_header_ii(t) or is_section_header_i(t):
                    break
                if is_address_line(t):
                    if current_people:
                        addr = extract_address(t)
                        if addr and not current_people[-1].get('address'):
                            current_people[-1]['address'] = addr
                            propagate_addresses(current_people)
                else:
                    p = parse_person(t)
                    if p: current_people.append(p)
                r += 1
            # nếu tiếp theo là II thì rơi vào nhánh dưới, không ++ thêm
            continue

        # Bảng thửa đất (đa cột)
        if is_section_header_ii(s):
            # dò hàng tiêu đề (ngay dưới dòng II)
            r += 1
            header_row = r
            # map tên cột -> index
            col_map = {}
            for c in range(df.shape[1]):
                h = strip_accents(cell(header_row, c)).lower()
                if 'ngay' in h and 'vao so' in h: col_map['ngay'] = c
                if 'so thu tu' in h and 'thua dat' in h: col_map['thua'] = c
                if 'so thu tu' in h and 'to ban do' in h: col_map['to'] = c
                if 'rieng' in h: col_map['rieng'] = c
                if 'chung' in h: col_map['chung'] = c
                if 'muc dich' in h: col_map['mucdich'] = c
                if 'thoi han' in h: col_map['thoihan'] = c
                if 'nguon goc' in h: col_map['nguongoc'] = c
                if ('so phat hanh' in h) or ('gcn qsdd' in h and '(9' in h): col_map['soph'] = c
                if ('so vao so' in h): col_map['sovs'] = c

            # đọc các dòng dữ liệu dưới header đến khi gặp dòng trống / dòng không có ngày
            r = header_row + 1
            while r < n_rows:
                v_ngay = cell(r, col_index if 'ngay' not in col_map else col_map['ngay'])
                if not re.match(r'\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$', v_ngay):
                    break
                def get(colkey):
                    if colkey not in col_map: return ''
                    return cell(r, col_map[colkey])
                thua = {
                    'ngay_thang_vao_so': get('ngay'),
                    'so_thu_tu_thua_dat': get('thua'),
                    'so_thu_tu_to_ban_do': get('to'),
                    'dt_rieng_m2': get('rieng').replace(',', '.'),
                    'dt_chung_m2': get('chung').replace(',', '.'),
                    'muc_dich_su_dung': get('mucdich'),
                    'thoi_han_su_dung': get('thoihan'),
                    'nguon_goc_su_dung': get('nguongoc'),
                    'so_phat_hanh_gcn': get('soph').replace(' ', ''),
                    'so_vao_so': get('sovs'),
                }
                current_parcels.append(thua)
                r += 1
            continue

        # nếu chưa vào I, bỏ qua
        r += 1

    # flush bìa cuối
    flush_bia()
    return rows

# ============ Run + GUI ============
def run_extraction(input_path: Path, output_path: Path, sheet, col_index: int):
    df = pd.read_excel(input_path, sheet_name=sheet, dtype=str)
    records = group_records_from_df(df, col_index)
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
    out_df.to_excel(output_path, index=False)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Xuất THỬA ĐẤT + Chủ (bắt đầu từ I, tiếp II)")
        self.geometry("740x360"); self.resizable(False, False)
        self.in_path_var = tk.StringVar(); self.out_path_var = tk.StringVar()
        self.col_var = tk.StringVar(value="0"); self.sheet_var = tk.StringVar(value="0")
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
        self.sheet_cb=ttk.Combobox(f3,textvariable=self.sheet_var,width=24,state="readonly",values=["0"])
        self.sheet_cb.pack(side='left',padx=8)
        ttk.Label(f3,text="Cột nguồn (index: 0=A, 1=B, ...):").pack(side='left',padx=(20,0))
        ttk.Entry(f3,textvariable=self.col_var,width=6).pack(side='left',padx=8)

        f4=ttk.Frame(self); f4.pack(fill='x', **pad)
        self.run_btn=ttk.Button(f4,text="Chạy",command=self.on_run); self.run_btn.pack(side='left')
        ttk.Button(f4,text="Thoát",command=self.destroy).pack(side='right')

        f5=ttk.Frame(self); f5.pack(fill='both',expand=True, **pad)
        ttk.Label(f5,text="Nhật ký:").pack(anchor='w'); self.log=tk.Text(f5,height=8); self.log.pack(fill='both',expand=True)

    def log_print(self,s:str):
        self.log.insert('end',s+"\n"); self.log.see('end'); self.update_idletasks()

    def choose_input(self):
        p=filedialog.askopenfilename(title="Chọn file Excel đầu vào",
            filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if not p: return
        self.in_path_var.set(p)
        try:
            xls=pd.ExcelFile(p); sheets=xls.sheet_names
            self.sheet_cb['values']=sheets; self.sheet_var.set(sheets[0] if sheets else "0")
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
        col_text=self.col_var.get().strip()
        if not in_path: messagebox.showwarning("Thiếu file đầu vào","Chọn file Excel đầu vào."); return
        if not out_path: messagebox.showwarning("Thiếu file đầu ra","Chọn nơi lưu file kết quả."); return
        try: sheet_arg=int(sheet)
        except: sheet_arg=sheet
        try: col_index=int(col_text)
        except: messagebox.showwarning("Cột nguồn không hợp lệ","Nhập số index: 0 cho A, 1 cho B, ..."); return
        try:
            self.run_btn.config(state='disabled'); self.log_print("Đang chạy trích xuất...")
            run_extraction(Path(in_path), Path(out_path), sheet_arg, col_index)
            self.log_print(f"Đã xuất: {out_path}")
            messagebox.showinfo("Hoàn tất", f"Đã xuất:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xử lý", str(e))
        finally:
            self.run_btn.config(state='normal')

if __name__ == "__main__":
    App().mainloop()
