import re
import sys
import pandas as pd
from pathlib import Path

# =========================
# Regex & Helpers
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
    - Và có CMND/CCCD trong dòng.
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
    if is_address_line(s):
        return None
    if not is_person_like(s):
        return None

    # Tên (đến dấu phẩy đầu tiên), bỏ tiền tố xưng hô
    m_name = re.search(
        r'^(?:\s*(?:(?:và|va|&)\s*)?'
        r'(?:hộ\s*(?:ông|bà)|ông|bà|anh|chị|cô|chú|bác|em)\s*[:\.-]?\s*)?'
        r'([A-ZÀ-ỸĐ][^,\d]+?)'
        r'\s*(?=,|\bSinh\b|\bCMND\b|\bCCCD\b|\bCMT\b|$)',
        s, flags=re.IGNORECASE
    )

    name = m_name.group(1).strip() if m_name else ''
    name = HONORIFIC.sub('', name).strip(' ,.-')

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
    m_issue = re.search(r'(?:cấp|cap)\s+ngày\s+' + RE_DATE, s, flags=re.IGNORECASE)
    if m_issue:
        issue_date = m_issue.group(1)

    if not name or (not id_number and not birth_year and not dob):
        return None

    return {
        'full_name': name,
        # nếu có ngày sinh cụ thể thì để ở cột Năm sinh là "dd/mm/yyyy" cho đủ thông tin
        'birth_year': birth_year or dob,
        'id_number': id_number,
        'issue_date': issue_date
    }

def record_to_wide(record_people):
    """Chuyển 1 bìa đỏ (list người) thành 1 dict dạng wide (Tên chủ i, ...)."""
    row = {}
    for idx, p in enumerate(record_people, start=1):
        row[f'Tên chủ {idx}']       = p.get('full_name', '')
        row[f'Năm sinh {idx}']      = p.get('birth_year', '')
        row[f'Số giấy tờ {idx}']    = p.get('id_number', '')
        row[f'Ngày cấp {idx}']      = p.get('issue_date', '')
    return row

def group_records(lines):
    """
    Gom dữ liệu theo 'bìa đỏ':
    - Gặp dòng bắt đầu bằng 'Hộ ...' -> mở record mới.
    - Các dòng tiếp theo (bỏ qua 'Địa chỉ ...') có CMND/CCCD -> thêm vào record hiện tại.
    - Gặp 'Hộ ...' lần nữa -> đóng record cũ, mở record mới.
    """
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

        # 4) Các dòng còn lại có CMND/CCCD -> add vào bìa hiện tại
        p = parse_person(s)
        if p:
            current_people.append(p)
            
    flush() # flush the last record
    return records


def main():
    # Cách dùng:
    # python extract_coowners_from_excel.py input.xlsx [output.xlsx] [sheet] [col_index]
    if len(sys.argv) < 2:
        print("Cách dùng:")
        print("  python extract_coowners_from_excel.py input.xlsx [output.xlsx] [sheet_name_or_index] [col_index]")
        print("Ví dụ:")
        print("  python extract_coowners_from_excel.py data.xlsx out.xlsx 0 0")
        return

    in_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else in_path.with_name(f"{in_path.stem}_coowners.xlsx")
    sheet_arg = sys.argv[3] if len(sys.argv) >= 4 else 0
    try:
        sheet_name = int(sheet_arg)
    except Exception:
        sheet_name = sheet_arg
    col_index = int(sys.argv[4]) if len(sys.argv) >= 5 else 0  # cột A = 0

    df = pd.read_excel(in_path, sheet_name=sheet_name, dtype=str)
    lines = df.iloc[:, col_index].fillna('').tolist()

    # Gom bìa đỏ -> danh sách [ [person1, person2, ...], ... ]
    records = group_records(lines)

    # Xác định số chủ tối đa để set cột đầy đủ
    max_people = max((len(r) for r in records), default=0)

    # Chuyển từng record thành wide row
    wide_rows = []
    for rec in records:
        row = record_to_wide(rec)
        wide_rows.append(row)

    out_df = pd.DataFrame(wide_rows)

    # Bổ sung cột trống đủ tới max_people để bảng đều cột
    for i in range(1, max_people + 1):
        for col in [f'Tên chủ {i}', f'Năm sinh {i}', f'Số giấy tờ {i}', f'Ngày cấp {i}']:
            if col not in out_df.columns:
                out_df[col] = ''

    # Sắp xếp cột theo thứ tự mong muốn
    ordered_cols = []
    for i in range(1, max_people + 1):
        ordered_cols += [f'Tên chủ {i}', f'Năm sinh {i}', f'Số giấy tờ {i}', f'Ngày cấp {i}']
    out_df = out_df.reindex(columns=ordered_cols)

    out_df.to_excel(out_path, index=False)
    print("Đã xuất:", out_path.resolve())

if __name__ == "__main__":
    main()
