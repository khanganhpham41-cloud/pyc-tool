# -*- coding: utf-8 -*-
"""Điền 1 trang Phiếu YCPT Hóa học (02.11-ST/6) / Sinh học (02.11-ST/5) và ghép nhiều trang.

Khác bản cũ trong Project:
- Dùng thẳng file .docx gốc thật trong thư mục mau_goc/ (bỏ cơ chế base64).
- Điền nhãn không phụ thuộc cách Word cắt run; nhãn nào không tìm thấy -> cảnh báo, không im lặng.
- Có điền Điện thoại / Địa chỉ đúng dòng in sẵn.
- Sinh học: hỗ trợ chỉ tiêu ngoài 12 cột (cột trống thứ 13, rồi mới thay cột không dùng),
  đổi đơn vị sang CFU/100mL cho mẫu nước.
"""
import copy
import os
import re

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from pyc_common import fill_after_label, SINH_NAMES, sinh_canonical, classify

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TPL_HOA = os.path.join(ROOT, "mau_goc", "PYC_PT_HOA_HOC_2025_11_17.docx")
TPL_SINH = os.path.join(ROOT, "mau_goc", "PYC_PT_SINH_HOC_2025_11_17.docx")

CHECK = "✓"
MAX_HOA_CT = 13

ABBREV = {
    "Nhóm Dư lượng thuốc bảo vệ thực vật Carbamate": "BVTV Carbamate",
    "Dư lượng thuốc bảo vệ thực vật nhóm Clo hữu cơ": "BVTV Clo hữu cơ",
    "Dư lượng thuốc bảo vệ thực vật nhóm Lân hữu cơ": "BVTV Lân hữu cơ",
    "Dư lượng thuốc bảo vệ thực vật nhóm Cúc tổng hợp": "BVTV Cúc tổng hợp",
}


# ---------------------------------------------------------------------------
# Tiện ích XML
# ---------------------------------------------------------------------------
def set_cell_text(cell, text, center=False, size_half_pt=None):
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
        run = p.runs[0]
    else:
        run = p.add_run(text)
    if size_half_pt:
        run.font.size = Pt(size_half_pt / 2)
    if center:
        p.alignment = 1
    return run


def set_row_height(row, twips, rule="atLeast"):
    trPr = row._tr.get_or_add_trPr()
    h = trPr.find(qn("w:trHeight"))
    if h is None:
        h = trPr.makeelement(qn("w:trHeight"), {})
        trPr.append(h)
    h.set(qn("w:val"), str(twips))
    h.set(qn("w:hRule"), rule)


def widen_cols(table, deltas):
    """deltas: {grid_index: +twips}; tổng phần cộng được trừ vào cột Mô tả mẫu (index 2)."""
    grid = table._tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    take = sum(deltas.values())
    new_w = {i: int(cols[i].get(qn("w:w"))) + d for i, d in deltas.items()}
    new_w[2] = int(cols[2].get(qn("w:w"))) - take
    for i, w in new_w.items():
        cols[i].set(qn("w:w"), str(w))
    for row in table.rows:
        tcs = row._tr.findall(qn("w:tc"))
        gi = 0
        for tc in tcs:
            tcPr = tc.find(qn("w:tcPr"))
            span = 1
            if tcPr is not None and tcPr.find(qn("w:gridSpan")) is not None:
                span = int(tcPr.find(qn("w:gridSpan")).get(qn("w:val")))
            if span == 1 and gi in new_w and tcPr is not None:
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is not None:
                    tcW.set(qn("w:w"), str(new_w[gi]))
            gi += span


def merge_group(t, row_idx, start_col, span):
    row = t.rows[row_idx]
    if span == 1:
        return row.cells[start_col]
    for c in range(start_col + 1, start_col + span):
        for r in row.cells[c].paragraphs[0].runs:
            r.text = ""
    merged = row.cells[start_col]
    for c in range(start_col + 1, start_col + span):
        merged = merged.merge(row.cells[c])
    for p in merged.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    return merged


def remove_rows(rows):
    for row in rows:
        row._tr.getparent().remove(row._tr)


def body_paragraphs_before_table(doc):
    """Các đoạn văn phần đầu phiếu (trước bảng mẫu) + đoạn sau bảng tới THỎA THUẬN."""
    out = []
    for p in doc.paragraphs:
        if "THỎA THUẬN VỀ CUNG CẤP DỊCH VỤ" in p.text:
            break
        out.append(p)
    return out


# ---------------------------------------------------------------------------
# Điền phần đầu phiếu (chung 2 loại)
# ---------------------------------------------------------------------------
HEADER_LABELS = [
    # key            regex nhãn                            bắt buộc phải tìm thấy?
    ("so_phieu",     None,                                  True),   # Số HH / Số VS (gán riêng)
    ("khach_hang",   r"Khách hàng\s*:?",                    True),
    ("dia_chi",      r"Địa chỉ\s*:?",                       False),
    ("dien_thoai",   r"Điện thoại\s*:?",                    False),
    ("email",        r"E-?mail\s*:?",                       False),
    ("so_luong_mau", r"Số lượng mẫu\s*:?",                  True),
    ("tai_lieu_kem", r"Tài liệu kèm theo(?:\s*\(nếu có\))?\s*:?", False),
    ("yeu_cau_khac", r"Các yêu cầu khác\s*:?", False),
]


def fill_header(doc, so_label, info, warnings):
    paras = body_paragraphs_before_table(doc)
    done = set()
    for key, rx, must in HEADER_LABELS:
        if key == "so_phieu":
            rx = so_label + r"\s*:?"
        val = info.get(key, "")
        if val in (None,):
            val = ""
        found = False
        for p in paras:
            if re.search(rx, p.text):
                if val != "":
                    fill_after_label(p, rx, val)
                found = True
                break
        if not found and (must or val):
            warnings.append(f"Không tìm thấy dòng '{rx}' trên phiếu gốc để điền '{val}'")
        done.add(key)
    # Trang x/y — sửa trong đúng run chứa "Trang" để giữ cỡ chữ gốc
    page_txt = f"Trang {info['trang_so']}/{info['tong_trang']}"
    rx = re.compile(r"Trang\s*[…\.]+\s*/\s*[…\.]+")
    ok = False
    for p in doc.paragraphs:
        for r in p.runs:
            if rx.search(r.text):
                r.text = rx.sub(page_txt, r.text, count=1)
                ok = True
                break
        if not ok:
            full = "".join(r.text for r in p.runs)
            m = rx.search(full)
            if m:
                # nhãn bị cắt nhiều run: gom phần "Trang…" vào run đầu tiên chứa chữ "Trang"
                pos, first = 0, None
                for r in p.runs:
                    a, b = pos, pos + len(r.text)
                    pos = b
                    if b <= m.start() or a >= m.end():
                        continue
                    lo, hi = max(m.start(), a) - a, min(m.end(), b) - a
                    if first is None:
                        r.text = r.text[:lo] + page_txt + r.text[hi:]
                        first = r
                    else:
                        r.text = r.text[:lo] + r.text[hi:]
                ok = True
        if ok:
            break
    if not ok:
        warnings.append("Không tìm thấy dòng 'Trang …./….' để đánh số trang")


# ---------------------------------------------------------------------------
# HÓA HỌC
# ---------------------------------------------------------------------------
SAFE_ROW_H = 2100
BASE_SIZE = 18          # 9pt — khớp cột Khối lượng mẫu gốc
MIN_SIZE = 10
TWIPS_PER_CHAR = 214


def _distribute(names, n_slots=13):
    """Mỗi chỉ tiêu 1 cột; tên nào không vừa chiều cao hàng ở cỡ 9pt thì gộp thêm cột
    (lấy từ số cột dư) thay vì thu nhỏ chữ. Hết cột dư mới thu nhỏ chữ."""
    import math
    n = len(names)
    sizes = [1] * n
    surplus = n_slots - n
    for i in sorted(range(n), key=lambda i: -len(names[i])):
        if surplus <= 0:
            break
        need = math.ceil(len(names[i]) * TWIPS_PER_CHAR / SAFE_ROW_H)
        take = min(max(0, need - 1), surplus)
        sizes[i] += take
        surplus -= take
    return sizes


HEADER_SIZE = 16        # 8pt — cỡ chữ chung cho mọi tên chỉ tiêu Hóa
HEADER_MARGIN = 20      # lề trái/phải ô tên chỉ tiêu (twips) để tên dài xuống được 2 dòng


def _set_header_uniform(cell, text):
    tc = cell._tc
    for p in list(tc.findall(qn("w:p")))[1:]:
        tc.remove(p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    ppr = p._p.get_or_add_pPr()
    old = ppr.find(qn("w:rPr"))
    if old is not None:
        ppr.remove(old)
    p.alignment = 1
    run = p.add_run(text)
    f = run.font
    f.name = "Times New Roman"
    f.size = Pt(HEADER_SIZE / 2)
    f.bold = False
    f.italic = False
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Times New Roman")
    tcPr = tc.get_or_add_tcPr()
    mar = tcPr.find(qn("w:tcMar"))
    if mar is None:
        mar = tcPr.makeelement(qn("w:tcMar"), {})
        # tcMar phải đứng trước textDirection/vAlign theo thứ tự schema
        anchor = next((e for e in (tcPr.find(qn(t)) for t in
                       ("w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")) if e is not None), None)
        if anchor is not None:
            anchor.addprevious(mar)
        else:
            tcPr.append(mar)
    for side in ("w:left", "w:right", "w:start", "w:end"):
        e = mar.find(qn(side))
        if e is None:
            e = mar.makeelement(qn(side), {})
            mar.append(e)
        e.set(qn("w:w"), str(HEADER_MARGIN))
        e.set(qn("w:type"), "dxa")


def build_hoa_page(out_path, info, mau_list, warnings):
    cols = []
    for m in mau_list:
        for c in m["chi_tieu"]:
            c = ABBREV.get(c, c)
            if c not in cols:
                cols.append(c)
    if len(cols) > MAX_HOA_CT:
        raise ValueError(f"Trang có {len(cols)} chỉ tiêu Hóa (> 13)")
    if len(mau_list) > 7:
        raise ValueError("Vượt 7 mẫu/trang Hóa học")

    doc = Document(TPL_HOA)
    t = doc.tables[0]
    widen_cols(t, {0: 160, 3: 220})  # cột KL đủ chỗ cho 4 chữ số (vd 1000)
    fill_header(doc, r"Số HH", info, warnings)

    # Định dạng thống nhất: 13 cột chỉ tiêu rộng bằng nhau (không gộp cột), cùng font/cỡ chữ,
    # không đậm/nghiêng (ô mẫu gốc mỗi ô một kiểu chữ khác nhau). Tên dài tự xuống dòng trong ô.
    ntp = set()
    raw = info.get("ntp") or ""          # "H2O2; Formaldehyde" — chỉ tiêu gửi nhà thầu phụ
    for x in (raw if isinstance(raw, list) else re.split(r"[;,]", str(raw))):
        if str(x).strip():
            ntp.add(ABBREV.get(classify(x)[1], classify(x)[1]))
    set_row_height(t.rows[1], SAFE_ROW_H)
    starts = [(4 + j, 1) for j in range(len(cols))]
    for j in range(MAX_HOA_CT):
        cell = t.rows[1].cells[4 + j]
        name = cols[j] if j < len(cols) else ""
        if name in ntp:
            name += " (NTP)"
        _set_header_uniform(cell, name)
    rows = list(t.rows[4:11])
    for i, (row, m) in enumerate(zip(rows, mau_list)):
        set_cell_text(row.cells[0], str(m["stt"]), center=True)
        set_cell_text(row.cells[1], m.get("ma", ""), center=True)
        set_cell_text(row.cells[2], m["mota"])
        set_cell_text(row.cells[3], str(m.get("kl", "")), center=True)
        mine = [ABBREV.get(c, c) for c in m["chi_tieu"]]
        for (start, sz), name in zip(starts, cols):
            cell = merge_group(t, 4 + i, start, sz)
            if name in mine:
                set_cell_text(cell, CHECK, center=True)
    # Hàng "Phương pháp thử" / "Yêu cầu công bố" (2 hàng cuối bảng) gộp cột theo đúng nhóm chỉ tiêu
    for ri in (len(t.rows) - 2, len(t.rows) - 1):
        for (start, sz) in starts:
            if sz > 1:
                merge_group(t, ri, start, sz)
    remove_rows(rows[len(mau_list):])
    doc.save(out_path)


# ---------------------------------------------------------------------------
# SINH HỌC
# ---------------------------------------------------------------------------
SINH_FIRST_COL = 4          # cell index của cột chỉ tiêu đầu tiên
SINH_NAME_ROW = 1           # hàng tên chỉ tiêu: mỗi ô có 2 đoạn (tên | đơn vị)
SINH_SAMPLE_ROWS = (5, 11)  # 6 dòng mẫu cố định
SINH_REF_COL = 9            # ô mẫu "S. aureus | cfu/g" để sao định dạng cho cột mới
WATER_UNIT = "CFU/100mL"


def _set_para_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def _set_sinh_header(cell, ref_cell, name, unit):
    tc = cell._tc
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)
    for p in ref_cell._tc.findall(qn("w:p")):
        tc.append(copy.deepcopy(p))
    paras = cell.paragraphs
    _set_para_text(paras[0], name)
    if len(paras) > 1:
        _set_para_text(paras[1], unit)


def _force_break_before_thoathuan(doc):
    """Phiếu Sinh gốc không có ngắt trang trước THỎA THUẬN (dựa vào dòng trống đẩy xuống).
    Khi xóa bớt dòng mẫu, phần thỏa thuận có thể bị kéo lên trang 1 -> xóa dòng trống đệm
    và đặt 'ngắt trang trước' cho đoạn THỎA THUẬN."""
    body = doc.element.body
    ch = list(body)
    idx = _thoathuan_idx(body)
    j = idx - 1
    while j >= 0 and ch[j].tag == qn("w:p"):
        el = ch[j]
        has_text = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
        has_obj = el.find(".//" + qn("w:drawing")) is not None or el.find(".//" + qn("w:pict")) is not None \
            or el.find(".//" + qn("w:sectPr")) is not None or el.find(".//" + qn("w:br")) is not None
        if has_text or has_obj:
            break
        body.remove(el)
        j -= 1
    para = None
    for p in doc.paragraphs:
        if "THỎA THUẬN VỀ CUNG CẤP DỊCH VỤ" in p.text:
            para = p
            break
    para.paragraph_format.page_break_before = True


def build_sinh_page(out_path, info, mau_list, warnings, nen_nuoc=False):
    if len(mau_list) > 6:
        raise ValueError("Vượt 6 mẫu/trang Sinh học")
    doc = Document(TPL_SINH)
    t = doc.tables[0]
    name_row = t.rows[SINH_NAME_ROW]
    widen_cols(t, {0: 180, 3: 160})
    fill_header(doc, r"Số VS", info, warnings)

    labels = list(SINH_NAMES) + [""]          # 13 vị trí: 12 in sẵn + 1 trống
    needed = []
    for m in mau_list:
        for c in m["chi_tieu"]:
            canon = sinh_canonical(c) or c
            if canon not in needed:
                needed.append(canon)
    extras = [c for c in needed if c not in labels]
    unused = [i for i, n in enumerate(labels[:12]) if n not in needed]
    prefer = ["C.perfringens", "V.cholerae", "Enterobacteriaceae", "Sta.coagulase (+)",
              "V.parahaemolyticus", "L.monocytogences", "S. aureus", "Nấm mốc/men"]
    unused.sort(key=lambda i: prefer.index(labels[i]) if labels[i] in prefer else 99)
    free_slots = [12] + unused
    if len(extras) > len(free_slots):
        raise ValueError(f"Quá nhiều chỉ tiêu Sinh ngoài danh mục trên 1 trang: {extras}")
    ref = name_row.cells[SINH_FIRST_COL + SINH_REF_COL]
    for c, slot in zip(extras, free_slots):
        old = labels[slot]
        labels[slot] = c
        unit = WATER_UNIT if nen_nuoc else "cfu/g"
        _set_sinh_header(name_row.cells[SINH_FIRST_COL + slot], ref, c, unit)
        if slot == 12:
            warnings.append(f"Phiếu Sinh: '{c}' điền vào cột trống thứ 13 (đơn vị {unit} — kiểm lại)")
        else:
            warnings.append(f"Phiếu Sinh: cột '{old}' (không mẫu nào kiểm) đổi nhãn thành '{c}' (đơn vị {unit} — kiểm lại)")

    if nen_nuoc:
        for slot, lab in enumerate(labels):
            if not lab or lab not in needed:
                continue
            paras = name_row.cells[SINH_FIRST_COL + slot].paragraphs
            if len(paras) > 1:
                u = paras[1].text.strip().lower()
                if u == "cfu/g":
                    _set_para_text(paras[1], WATER_UNIT)
                elif u and u != WATER_UNIT.lower():
                    warnings.append(f"Mẫu nước: cột '{lab}' đơn vị in sẵn '{paras[1].text.strip()}' — chưa tự đổi, ghi tay nếu cần")

    a, b = SINH_SAMPLE_ROWS
    rows = list(t.rows[a:b])
    for row, m in zip(rows, mau_list):
        set_cell_text(row.cells[0], str(m["stt"]), center=True)
        set_cell_text(row.cells[1], m.get("ma", ""), center=True)
        set_cell_text(row.cells[2], m["mota"])
        set_cell_text(row.cells[3], str(m.get("kl", "")), center=True)
        mine = [sinh_canonical(c) or c for c in m["chi_tieu"]]
        for slot, lab in enumerate(labels):
            if lab and lab in mine:
                set_cell_text(row.cells[SINH_FIRST_COL + slot], CHECK, center=True)
    remove_rows(rows[len(mau_list):])
    _force_break_before_thoathuan(doc)
    doc.save(out_path)


# ---------------------------------------------------------------------------
# GHÉP TRANG
# ---------------------------------------------------------------------------
def _thoathuan_idx(body):
    for i, el in enumerate(body):
        if el.tag == qn("w:p"):
            txt = "".join(x.text or "" for x in el.iter(qn("w:t")))
            if "THỎA THUẬN VỀ CUNG CẤP DỊCH VỤ" in txt:
                # lùi về đoạn ngắt trang ngay trước (nếu có) để giữ nó cho phần thỏa thuận
                return i
    raise ValueError("Không tìm thấy THỎA THUẬN trong phiếu")


def _renumber(elements, ctr):
    bk_map = {}
    for el in elements:
        for node in el.iter():
            if not isinstance(node.tag, str):
                continue
            tag = node.tag
            if tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                old = node.get(qn("w:id"))
                if old not in bk_map:
                    ctr["bk"] += 1
                    bk_map[old] = str(ctr["bk"])
                node.set(qn("w:id"), bk_map[old])
                if tag == qn("w:bookmarkStart") and node.get(qn("w:name")):
                    node.set(qn("w:name"), f"{node.get(qn('w:name'))}_{ctr['bk']}")
            elif tag == qn("wp:docPr"):
                ctr["dp"] += 1
                node.set("id", str(ctr["dp"]))
                node.set("name", f"Shape {ctr['dp']}")
            elif tag.startswith("{urn:schemas-microsoft-com:vml}") and node.get("id"):
                ctr["vml"] += 1
                node.set("id", f"_x0000_s{2000 + ctr['vml']}")
            if node.get("{urn:schemas-microsoft-com:office:office}spid"):
                ctr["vml"] += 1
                node.set("{urn:schemas-microsoft-com:office:office}spid", f"_x0000_s{2000 + ctr['vml']}")


def _ends_with_page_break(elements):
    for e in reversed(elements):
        if e.tag != qn("w:p"):
            return False
        for br in e.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
        txt = "".join(x.text or "" for x in e.iter(qn("w:t"))).strip()
        if txt:
            return False
    return False


def _page_break_para():
    from docx.oxml import parse_xml
    return parse_xml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                     '<w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>'
                     '<w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:br w:type="page"/></w:r></w:p>')


def compose_pages(page_files, out_path):
    base = Document(page_files[0])
    body = base.element.body
    parts = []
    tail = None
    for f in page_files:
        d = Document(f) if f != page_files[0] else base
        b = d.element.body
        idx = _thoathuan_idx(b)
        children = list(b)
        parts.append([copy.deepcopy(e) for e in children[:idx]])
        if f == page_files[-1]:
            tail = [copy.deepcopy(e) for e in children[idx:]]
    ctr = {"bk": 0, "dp": 1000, "vml": 0}
    for p in parts:
        _renumber(p, ctr)
    _renumber(tail, ctr)
    for e in list(body):
        body.remove(e)
    for k, p in enumerate(parts):
        if k > 0 and not _ends_with_page_break(parts[k - 1]):
            body.append(_page_break_para())
        for e in p:
            body.append(e)
    for e in tail:
        body.append(e)
    base.save(out_path)
