import re
from converter import Converter

conv = Converter()

# Tách THEO KHOẢNG TRẮNG để không cắt giữa chữ & byte lạ (như "n¬i")
_WS_SPLIT = re.compile(r'(\s+)', flags=re.UNICODE)

# Dấu hiệu còn “bảng mã cũ” trong chuỗi
LEGACY_MARKERS = re.compile(
    r'[¬«¨¸µÊÆ§¯¾»¼½ÞßþØøÏîëìäåãèéæçñª]',
    flags=re.UNICODE
)

# 1) Chấm điểm “độ Việt” của một chuỗi Unicode: nhiều ký tự có dấu + từ khoá hay gặp
VN_UNI = re.compile(r'[À-ỹ]')

def score_vn(s: str) -> int:
    if not isinstance(s, str) or not s:
        return 0
    score = len(VN_UNI.findall(s))
    base = (s or '').lower()
    for kw in ("ngày", "cấp", "nơi", "công", "tỉnh", "sinh", "địa", "chỉ", "thửa", "bản", "đồ", "phú", "yên"):
        if kw in base:
            score += 3
    return score

# 2) Thử convert theo nhiều bảng mã và chọn bản điểm cao nhất
CANDIDATES = ["TCVN3","VNI_WIN","VPS_WIN","VIETWARE_X","VIETWARE_F","VISCII","VIQR"]

def convert_best(s: str) -> str:
    best = s
    best_sc = score_vn(s)
    for cs in CANDIDATES:
        try:
            t = conv.convert(s, target_charset="UNICODE", source_charset=cs)
        except Exception:
            continue
        sc = score_vn(t)
        if sc > best_sc:
            best, best_sc = t, sc
    return best

# 3) Sửa chuỗi pha trộn: theo token -> rescue toàn chuỗi
def fix_mixed_to_unicode(s: str) -> str:
    if not isinstance(s, str):
        return ""
    parts = _WS_SPLIT.split(s)  # [token, space, token, space...]
    out = []
    for tok in parts:
        if tok is None or tok == "":
            continue
        if tok.isspace():
            out.append(tok)
            continue
        # ưu tiên detect; nếu detect != UNICODE thì convert theo detect
        cs = conv.detectCharset(tok)
        if cs and cs != "UNICODE":
            try:
                tok = conv.convert(tok, target_charset="UNICODE", source_charset=cs)
            except Exception:
                pass
        elif LEGACY_MARKERS.search(tok):
            # không chắc bảng mã -> thử nhiều và chọn bản tốt nhất
            tok = convert_best(tok)
        out.append(tok)
    result = "".join(out)

    # Rescue pass: nếu vẫn còn marker, thử convert toàn chuỗi và lấy bản tốt hơn
    if LEGACY_MARKERS.search(result):
        candidate = convert_best(result)
        if score_vn(candidate) > score_vn(result):
            result = candidate
    return result

# 4) Dùng trong pipeline chuẩn hoá
def to_text(v) -> str:
    if v is None:
        return ""
    s = str(v)
    # CONVERT TRƯỚC
    s = fix_mixed_to_unicode(s)
    # rồi normalize khoảng trắng/ký tự đặc biệt
    s = s.replace('\u00A0', ' ')
    s = re.sub(r'[_]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

print(to_text("Hộ ông: Nguyễn A, sinh năm 1938, số CMND 220515369, ngày cấp 04/01/1997, n¬i cấp c«ng an tØnh Phó Yªn")) # should print empty string