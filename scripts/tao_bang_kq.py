# -*- coding: utf-8 -*-
"""Bảng Excel tổng hợp kết quả cho 1 lô mẫu (chuẩn SKILL.md mục 7, Vinh chốt 14/10/2026).

Bật bằng khóa "bang_kq" trong thong_tin của file JSON hàng chờ:
  "bang_kq": {"tieu_de": "KẾT QUẢ KIM LOẠI NẶNG", "ten_file": "KQ_KLN_NAFI4_141026",
              "chi_tieu": ["Hg","Cd","Pb"],          # tùy chọn — mặc định lấy theo thứ tự xuất hiện
              "loai": "H",                          # H = Phòng KN Hóa học (mặc định) / S = Vi sinh
              "don_vi": {"thuc_pham": "mg/kg", "nuoc": "mg/L"}}   # tùy chọn
  "chi_bang_kq": true   -> chỉ tạo Excel, không tạo phiếu Word
Mỗi mẫu có thể có "so_phieu" riêng (lô nhiều phiếu); thiếu thì lấy so_hh / so_vs.
Mẫu có "so_bb" (số biên bản / số nhận diện) -> cột riêng; không có thì để trống.
"""
import datetime
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from pyc_common import split_chi_tieu

FONT = "Times New Roman"
YEL = PatternFill("solid", fgColor="FFF2CC")
GRY = PatternFill("solid", fgColor="D9D9D9")
HDR = PatternFill("solid", fgColor="BDD7EE")
_th = Side(style="thin", color="808080")
BOX = Border(left=_th, right=_th, top=_th, bottom=_th)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEF = Alignment(horizontal="left", vertical="center", wrap_text=True)


def F(size=11, **k):
    return Font(name=FONT, size=size, **k)


def _range(xs):
    xs = [x for x in xs if x]
    uniq = list(dict.fromkeys(xs))
    if not uniq:
        return ""
    return uniq[0] if len(uniq) == 1 else f"{uniq[0]} → {uniq[-1]} ({len(uniq)} phiếu)"


def build(info, mau_raw, root, out_name=None):
    """info: thong_tin (dict); mau_raw: list mẫu như trong JSON. Trả về đường dẫn file .xlsx."""
    cfg = info.get("bang_kq") or {}
    if not isinstance(cfg, dict):
        cfg = {}
    loai = (cfg.get("loai") or "H").upper()
    so_mac_dinh = info.get("so_vs") if loai == "S" else info.get("so_hh")
    dv = {"thuc_pham": "mg/kg", "nuoc": "mg/L", **(cfg.get("don_vi") or {})}

    rows, ct_all = [], []
    for i, m in enumerate(mau_raw, 1):
        cts = split_chi_tieu(m.get("chi_tieu", ""))
        for c in cts:
            if c not in ct_all:
                ct_all.append(c)
        nuoc = str(m.get("nen", "")).strip().lower() == "nước"
        stt = str(m.get("stt") or i)
        rows.append({"stt": int(stt) if stt.isdigit() else stt, "so": m.get("so_phieu") or so_mac_dinh or "",
                     "ma": m.get("ma_vs" if loai == "S" else "ma_hh") or m.get("ma", ""),
                     "mota": m.get("mota", ""), "bb": m.get("so_bb", ""), "kl": m.get("kl", ""),
                     "nen": "Nước" if nuoc else "Thực phẩm",
                     "dv": m.get("don_vi") or (dv["nuoc"] if nuoc else dv["thuc_pham"]), "ct": cts})
    cols_ct = cfg.get("chi_tieu") or ct_all

    wb = Workbook()
    ws = wb.active
    ws.title = "KQ"
    fixed = ["STT", "Số phiếu", "Mã mẫu", "Tên mẫu", "Số BB / Số nhận diện", "KL (g)", "Nền", "Đơn vị"]
    n = len(fixed) + len(cols_ct)
    half = n // 2

    def merge(r, c1, c2, text, font=None, al=LEF):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        x = ws.cell(r, c1, text)
        x.font = font or F()
        x.alignment = al

    merge(1, 1, n, cfg.get("tieu_de") or "BẢNG KẾT QUẢ PHÂN TÍCH", F(14, bold=True), CEN)
    dt_mail = " – ".join(x for x in (info.get("dien_thoai", ""), info.get("email", "")) if x)
    head = [("Khách hàng:", info.get("khach_hang", "")), ("Địa chỉ:", info.get("dia_chi", "")),
            ("Điện thoại / Email:", dt_mail), ("Tài liệu kèm:", info.get("tai_lieu_kem", "")),
            ("Yêu cầu:", info.get("yeu_cau_khac", "")),
            ("Số phiếu YCPT:", f"{_range([r['so'] for r in rows])} – {len(rows)} mẫu")]
    r = 2
    for k, v in head:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.cell(r, 1, k).font = F(bold=True)
        merge(r, 4, n, v)
        r += 1
    H = r + 1
    for j, c in enumerate(fixed + cols_ct, 1):
        x = ws.cell(H, j, c)
        x.font, x.fill, x.alignment, x.border = F(bold=True), HDR, CEN, BOX
    r = H + 1
    for d in rows:
        kl = d["kl"]
        try:
            kl = int(float(kl))
        except (TypeError, ValueError):
            pass
        for j, v in enumerate([d["stt"], d["so"], d["ma"], d["mota"], d["bb"], kl, d["nen"], d["dv"]], 1):
            x = ws.cell(r, j, v)
            x.font, x.border = F(), BOX
            x.alignment = LEF if j in (4, 5) else CEN
        for k, c in enumerate(cols_ct):
            x = ws.cell(r, len(fixed) + 1 + k)
            x.border, x.alignment = BOX, CEN
            if c in d["ct"]:
                x.fill, x.font = YEL, F(color="0000FF")
            else:
                x.value, x.fill, x.font = "—", GRY, F(color="808080")
        ws.row_dimensions[r].height = 20
        r += 1
    last = r - 1

    # Cuối trang kiểu 02.11-ST — chỉ 2 bên ký
    r = last + 2
    f12 = F(12)
    merge(r, 1, half, "Ngày nhận mẫu: ………/………./20….", f12)
    merge(r, half + 2, n, "Ngày phân tích: …….... / ……..... / 20……", f12)
    merge(r + 1, 1, half, "Ngày hẹn trả kết quả: …….. / ……../ 20….", f12)
    merge(r + 1, half + 2, n, "Ngày có kết quả: …….... / ……..... / 20……", f12)
    merge(r + 3, 1, half, "Tổ nhận mẫu", f12, CEN)
    merge(r + 3, half + 2, n, "Phòng kiểm nghiệm " + ("Vi sinh" if loai == "S" else "Hóa học"), f12, CEN)
    end = r + 9

    for j, w in enumerate([5, 11, 9, 18, 20, 7, 10, 8] + [9] * len(cols_ct), 1):
        ws.column_dimensions[L(j)].width = w
    ws.freeze_panes = ws.cell(H + 1, 5)
    ws.auto_filter.ref = f"A{H}:{L(n)}{last}"
    ps = ws.page_setup
    ps.orientation, ps.paperSize, ps.fitToWidth, ps.fitToHeight = "landscape", 9, 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_title_rows = f"{H}:{H}"
    ws.print_area = f"A1:{L(n)}{end}"
    ws.oddFooter.center.text = "Trang &P/&N"
    ws.oddFooter.center.font = FONT

    now = datetime.datetime.now()
    sub = os.path.join(root, "VI_SINH" if loai == "S" else "HOA_HOC", now.strftime("%Y"), now.strftime("%Y-%m"))
    os.makedirs(sub, exist_ok=True)
    name = out_name or cfg.get("ten_file") or ("KQ_" + (info.get("ten_file") or "LO") + now.strftime("_%d%m%y"))
    out = os.path.join(sub, name if name.lower().endswith(".xlsx") else name + ".xlsx")
    wb.save(out)   # cùng tên -> ghi đè (bản đúng thay bản sai)
    return out, len(rows), len(cols_ct)
