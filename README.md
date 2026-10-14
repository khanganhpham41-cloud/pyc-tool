# pyc-tool
Cong cu tao Phieu yeu cau phan tich Hoa/Sinh TTV6 (02.11-ST/5, 02.11-ST/6). Xem HUONG_DAN.txt.

Khong luu du lieu khach hang / so theo doi trong kho nay.

## Bảng Excel kết quả (từ 14/10/2026)
Thêm vào `thong_tin` của JSON: `"bang_kq": {"tieu_de": "...", "ten_file": "KQ_...", "chi_tieu": [...]}` → tạo
`HOA_HOC\<năm>\<năm-tháng>\KQ_....xlsx` (cấu trúc chuẩn trong `scripts/tao_bang_kq.py`).
`"chi_bang_kq": true` → chỉ tạo Excel. Mẫu có `"so_phieu"` khác nhau → mỗi số 1 phiếu Word.
