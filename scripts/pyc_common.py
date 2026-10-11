# -*- coding: utf-8 -*-
"""Phần dùng chung: phân loại chỉ tiêu Hóa/Sinh, chia trang, thay nhãn trong đoạn văn."""
import re
import unicodedata

# ---------------------------------------------------------------------------
# 1. Phân loại chỉ tiêu
# ---------------------------------------------------------------------------
# 12 cột in sẵn trên phiếu Sinh học (đúng thứ tự trên phiếu gốc) + bí danh hay gặp
SINH_COLS = [
    ("TPC (30oC)",         ["tpc", "tổng vi sinh vật hiếu khí", "tong vi sinh vat hieu khi",
                             "tổng số vi sinh vật hiếu khí", "tvkhk", "apc", "vi sinh vật hiếu khí"]),
    ("Enterobacteriaceae", ["enterobacteriaceae"]),
    ("Coliforms",          ["coliforms", "coliform", "tổng coliforms", "tổng coliform"]),
    ("E.coli",             ["e.coli", "e. coli", "ecoli", "escherichia coli"]),
    ("Sta.coagulase (+)",  ["sta.coagulase (+)", "staphylococci coagulase dương tính",
                             "staphylococcus coagulase", "sta.coagulase", "tụ cầu coagulase"]),
    ("S. aureus",          ["s. aureus", "s.aureus", "staphylococcus aureus"]),
    ("Salmonella spp.",    ["salmonella", "salmonella spp", "salmonella spp."]),
    ("V.cholerae",         ["v.cholerae", "v. cholerae", "vibrio cholerae"]),
    ("V.parahaemolyticus", ["v.parahaemolyticus", "v. parahaemolyticus", "vibrio parahaemolyticus"]),
    ("L.monocytogences",   ["l.monocytogenes", "l. monocytogenes", "listeria monocytogenes",
                             "l.monocytogences", "listeria"]),
    ("C.perfringens",      ["c.perfringens", "c. perfringens", "clostridium perfringens"]),
    ("Nấm mốc/men",        ["nấm mốc/men", "nấm men, nấm mốc", "nấm men nấm mốc", "nấm mốc, nấm men",
                             "tổng nấm men nấm mốc", "nấm mốc", "nấm men", "yeast", "mold", "tymc"]),
]
SINH_NAMES = [c for c, _ in SINH_COLS]

# Chỉ tiêu vi sinh KHÔNG có cột in sẵn — vẫn thuộc phiếu Sinh học (dùng cột trống thứ 13 / thay cột)
SINH_EXTRA = ["enterococci", "enterococcus", "streptococci feacal", "liên cầu phân",
              "pseudomonas aeruginosa", "p. aeruginosa", "p.aeruginosa",
              "bacillus cereus", "b. cereus", "b.cereus", "shigella",
              "clostridia khử sulfit", "bào tử vi khuẩn kỵ khí khử sulfit",
              "coliform chịu nhiệt", "coliforms chịu nhiệt", "fecal coliform", "fecal coliforms",
              "campylobacter", "cronobacter", "vibrio spp", "vibrio", "streptococcus", "clostridium",
              "yersinia", "staphylococcus"]


def norm(s):
    s = unicodedata.normalize("NFC", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)


def _key(s):
    return norm(s).lower().rstrip(".")


# Nhận diện theo TỪ KHÓA (không cần khớp nguyên văn) — thứ tự kiểm tra quan trọng
_COL_RULES = [
    ("Sta.coagulase (+)",  [r"coagulase"]),
    ("S. aureus",          [r"aureus"]),
    ("E.coli",             [r"\be\.?\s?coli\b", r"escherichia"]),
    ("Enterobacteriaceae", [r"enterobacteriaceae"]),
    ("Coliforms",          [r"coliform"]),          # (chịu nhiệt/fecal đã loại ở dưới)
    ("Salmonella spp.",    [r"salmonella"]),
    ("V.cholerae",         [r"cholerae"]),
    ("V.parahaemolyticus", [r"parahaem"]),
    ("L.monocytogences",   [r"monocytogen", r"listeria"]),
    ("C.perfringens",      [r"perfringens"]),
    ("Nấm mốc/men",        [r"nấm\s*(men|mốc)", r"yeast", r"mou?ld", r"\btymc\b"]),
    ("TPC (30oC)",         [r"\btpc\b", r"\bapc\b", r"hiếu\s*khí", r"tvkhk"]),
]
# Không thuộc cột in sẵn nhưng chắc chắn là vi sinh -> phiếu Sinh (cột 13 / thay cột)
_MICRO_HINTS = [r"chịu\s*nhiệt", r"fecal", r"faecal", r"enterococc", r"streptococc", r"pseudomonas",
                r"bacillus", r"cereus", r"shigella", r"campylobacter", r"cronobacter", r"vibrio",
                r"clostridi", r"yersinia", r"staphylococ", r"khử\s*sulf", r"kỵ\s*khí", r"bào\s*tử",
                r"vi\s*khuẩn", r"vi\s*sinh", r"\bmpn\b", r"\bcfu\b", r"virus", r"norovirus",
                r"ký\s*sinh\s*trùng", r"\bspp\b"]


def _clean(ct):
    k = _key(ct)
    k = re.sub(r"/\s*\d*\s*(g|ml|mL)\b", " ", k)      # /25g, /g, /100ml
    k = re.sub(r"\s+", " ", k).strip()
    return k


def sinh_canonical(ct):
    """Tên cột in sẵn trên phiếu Sinh nếu chỉ tiêu thuộc 1 trong 12 cột, ngược lại None."""
    k = _clean(ct)
    for name, aliases in SINH_COLS:
        if k == _key(name) or k in [_key(a) for a in aliases]:
            return name
    if _lookup(SINH_EXTRA_MAP, ct):      # vd Listeria spp., E.coli O157 — không phải cột in sẵn
        return None
    for name, aliases in SINH_COLS:
        if k == _key(name) or k in [_key(a) for a in aliases]:
            return name
    if re.search(r"chịu\s*nhiệt|fecal|faecal", k):       # coliform chịu nhiệt != Coliforms
        return None
    for name, pats in _COL_RULES:
        if any(re.search(p, k) for p in pats):
            return name
    return None


# ---------------------------------------------------------------------------
# Danh mục tên chuẩn theo GCN 1586/TĐC-HCHQ (scripts/danh_muc_chi_tieu.json)
# ---------------------------------------------------------------------------
def _fold(s):
    """Khóa so khớp: bỏ dấu, chữ thường, bỏ ký tự đặc biệt."""
    s = unicodedata.normalize("NFD", norm(s).lower())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").replace("đ", "d")
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _load_danh_muc():
    import json
    import os
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "danh_muc_chi_tieu.json")
    hoa, sinh = {}, {}
    try:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return hoa, sinh
    for sec, tgt in (("hoa", hoa), ("sinh_them", sinh)):
        for name, aliases in d.get(sec, {}).items():
            for a in [name] + list(aliases):
                tgt.setdefault(_fold(a), name)
    return hoa, sinh


HOA_MAP, SINH_EXTRA_MAP = _load_danh_muc()
UNKNOWN = set()   # tên chưa có trong danh mục — tao_phieu.py in cảnh báo


def _lookup(mapping, t):
    for cand in (t, _clean(t)):
        name = mapping.get(_fold(cand))
        if name:
            return name
    return None


def classify(ct):
    """('S'|'H', tên chuẩn). Ép loại bằng tiền tố 'S:' / 'H:' nếu cần.
    Thứ tự: tiền tố ép -> danh mục Hóa -> 12 cột Sinh -> Sinh ngoài cột -> dấu hiệu vi sinh -> Hóa (lạ)."""
    t = norm(ct)
    m = re.match(r"^([SsHh])\s*:\s*(.+)$", t)
    if m:
        kind, name = m.group(1).upper(), norm(m.group(2))
        if kind == "H":
            return "H", _lookup(HOA_MAP, name) or name
        return "S", sinh_canonical(name) or _lookup(SINH_EXTRA_MAP, name) or name
    h = _lookup(HOA_MAP, t)
    if h:
        return "H", h
    c = sinh_canonical(t)
    if c:
        return "S", c
    s = _lookup(SINH_EXTRA_MAP, t)
    if s:
        return "S", s
    k = _clean(t)
    if any(k.startswith(e) for e in SINH_EXTRA) or any(re.search(p, k) for p in _MICRO_HINTS):
        UNKNOWN.add(("Sinh", t))
        return "S", t
    UNKNOWN.add(("Hóa", t))
    return "H", t


def split_chi_tieu(text):
    """Tách ô 'Chỉ tiêu' thành danh sách.
    Ngăn cách bằng ; hoặc xuống dòng. Dạng 'X + Y' = 2 chỉ tiêu riêng (quy ước đã chốt)."""
    out = []
    for part in re.split(r"[;\n\r]+", str(text or "")):
        part = norm(part)
        if not part:
            continue
        for p in re.split(r"\s\+\s", part):
            p = norm(p)
            if p and p not in out:
                out.append(p)
    return out


# ---------------------------------------------------------------------------
# 2. Chia trang — giữ nguyên thứ tự, chia đều, tôn trọng giới hạn chỉ tiêu/trang
# ---------------------------------------------------------------------------
def _distinct(chunk, key="chi_tieu"):
    seen = []
    for m in chunk:
        for c in m[key]:
            if c not in seen:
                seen.append(c)
    return seen


def paginate(samples, per_page, max_ct=None):
    """Chia danh sách mẫu thành các trang (giữ thứ tự).
    Ưu tiên chia đều (vd 23 mẫu/5 → 5,5,5,4,4); nếu vượt giới hạn chỉ tiêu/trang thì
    tách thêm trang (tham lam), vẫn giữ thứ tự gốc."""
    n = len(samples)
    if n == 0:
        return []
    import math
    k = math.ceil(n / per_page)
    while True:
        base, extra = divmod(n, k)
        sizes = [base + (1 if i < extra else 0) for i in range(k)]
        pages, i = [], 0
        for s in sizes:
            pages.append(samples[i:i + s])
            i += s
        if max_ct is None or all(len(_distinct(p)) <= max_ct for p in pages):
            return pages
        k += 1
        if k > n:
            break
    # fallback tham lam (1 mẫu có > max_ct chỉ tiêu sẽ báo lỗi ở bước điền)
    pages, cur = [], []
    for m in samples:
        trial = cur + [m]
        if cur and (len(trial) > per_page or len(_distinct(trial)) > max_ct):
            pages.append(cur)
            cur = [m]
        else:
            cur = trial
    if cur:
        pages.append(cur)
    return pages


# ---------------------------------------------------------------------------
# 3. Điền giá trị vào sau nhãn trong 1 đoạn văn — không phụ thuộc cách Word cắt "run"
# ---------------------------------------------------------------------------
PLACEHOLDER = r"[ \t\.…_ ]*"


def fill_after_label(paragraph, label, value, placeholder=PLACEHOLDER):
    """Tìm `label` (regex) trong đoạn văn, xóa chuỗi chấm/… ngay sau nó rồi chèn `value`.
    Giữ nguyên định dạng từng run. Trả về True nếu tìm thấy nhãn."""
    runs = paragraph.runs
    full = "".join(r.text for r in runs)
    m = re.search(label, full)
    if not m:
        return False
    ph = re.compile(placeholder).match(full, m.end())
    ps, pe = m.end(), (ph.end() if ph else m.end())
    # Giữ lại 1 khoảng trắng nếu nhãn kết thúc bằng ":" và value không rỗng
    insert = (" " if value and not full[m.end() - 1:m.end()].isspace() else "") + str(value)
    if value and pe < len(full) and not full[pe].isspace():
        insert += " "
    pos = 0
    inserted = False
    for r in runs:
        t = r.text
        a, b = pos, pos + len(t)
        pos = b
        if b <= ps and not (b == ps and not inserted and pe == ps):
            continue
        if a >= pe and inserted:
            break
        lo, hi = max(ps, a) - a, min(pe, b) - a
        if lo < 0:
            lo = 0
        if hi < lo:
            hi = lo
        if not inserted:
            r.text = t[:lo] + insert + t[hi:]
            inserted = True
        else:
            r.text = t[:lo] + t[hi:]
    return inserted
