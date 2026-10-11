# -*- coding: utf-8 -*-
"""Chạy ngầm: thấy file .json mới trong HANG_CHO (laptop + Google Drive) -> tự tạo phiếu, tự làm sạch bằng Word,
hiện thông báo góc màn hình, ghi nhật ký NHAT_KY_TU_DONG.txt. Không cần bấm TAO_PHIEU.bat.

Cài 1 lần: CAI_TU_DONG.bat (tự chạy mỗi khi mở máy).   Tắt: TAT_TU_DONG.bat
Chạy thử 1 vòng: python tu_dong.py --mot-lan
"""
import datetime
import os
import socket
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import tao_phieu  # noqa: E402  (chỉ dùng hàm tìm hàng chờ)

LOG = os.path.join(ROOT, "NHAT_KY_TU_DONG.txt")
CHU_KY = 15          # giây giữa 2 lần kiểm tra
CHO_ON_DINH = 5      # file phải đứng yên 5 giây (Drive đồng bộ xong) mới xử lý
WIN = sys.platform == "win32"
NO_WINDOW = 0x08000000 if WIN else 0


def ghi_log(text):
    try:
        if os.path.exists(LOG) and os.path.getsize(LOG) > 500_000:   # giữ nhật ký gọn
            with open(LOG, encoding="utf-8", errors="ignore") as f:
                tail = f.read()[-200_000:]
            with open(LOG, "w", encoding="utf-8") as f:
                f.write(tail)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(text.rstrip() + "\n")
    except Exception:
        pass


def thong_bao(tieu_de, noi_dung):
    """Thông báo góc phải màn hình (Windows 10/11). Lỗi thì bỏ qua."""
    if not WIN:
        print(f"[THÔNG BÁO] {tieu_de}: {noi_dung}")
        return
    esc = lambda s: s.replace("'", "''")  # noqa: E731
    ps = (
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
        "$t=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent("
        "[Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
        "$x=$t.GetElementsByTagName('text');"
        f"$x.Item(0).AppendChild($t.CreateTextNode('{esc(tieu_de)}')) > $null;"
        f"$x.Item(1).AppendChild($t.CreateTextNode('{esc(noi_dung)}')) > $null;"
        "$app='{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe';"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show("
        "[Windows.UI.Notifications.ToastNotification]::new($t))"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                       creationflags=NO_WINDOW, timeout=30, capture_output=True)
    except Exception:
        pass


def file_cho():
    """Các file .json đã ổn định trong mọi hàng chờ."""
    out, now = [], time.time()
    for q in tao_phieu.queue_dirs():
        try:
            names = os.listdir(q)
        except OSError:
            continue
        for f in names:
            p = os.path.join(q, f)
            if f.lower().endswith(".json") and not f.startswith(("~", ".")) and os.path.isfile(p):
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                if now - st.st_mtime >= CHO_ON_DINH and st.st_size > 0:
                    out.append((p, st.st_mtime))
    return out


def chay_mot_lan(da_loi):
    """Tạo phiếu cho mọi file mới. da_loi: {(đường dẫn, mtime)} file lỗi đã báo — không báo lại tới khi file được sửa."""
    moi = [x for x in file_cho() if x not in da_loi]
    if not moi:
        return
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    py = sys.executable
    if WIN and py.lower().endswith("pythonw.exe"):
        py = py[:-len("pythonw.exe")] + "python.exe"
    try:
        r = subprocess.run([py, os.path.join(HERE, "tao_phieu.py")], cwd=ROOT, env=env,
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           creationflags=NO_WINDOW, timeout=300)
        out = r.stdout + r.stderr
    except subprocess.TimeoutExpired:
        out = "*** Tạo phiếu quá 5 phút — đã dừng. Chạy tay TAO_PHIEU.bat để xem lỗi."
    ds = os.path.join(HERE, "_danh_sach_file_moi.txt")
    tao = []
    if os.path.exists(ds):
        with open(ds, encoding="utf-8") as f:
            tao = [ln.split("|")[0] for ln in f if ln.strip()]
    if tao and WIN:
        try:
            w = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                                "-File", os.path.join(HERE, "lam_sach_bang_word.ps1"), ds],
                               cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace",
                               creationflags=NO_WINDOW, timeout=180)
            out += "\n" + w.stdout + w.stderr
        except subprocess.TimeoutExpired:
            # Word treo (thường do hộp thoại kích hoạt/khôi phục file). KHÔNG tắt Word bằng lệnh —
            # có thể làm mất văn bản Vinh đang soạn. Chỉ báo để mở tay kiểm tra.
            out += ("\n!!! Word không phản hồi sau 3 phút (có thể đang hiện hộp thoại) — phiếu ĐÃ tạo nhưng chưa "
                    "được Word làm sạch: mở từng file bằng Word, lưu lại trước khi gửi. Nếu có hộp thoại Word đang chờ, bấm đóng nó.")
    ghi_log(f"\n######## {datetime.datetime.now():%d/%m/%Y %H:%M:%S} ########\n{out}")

    # File còn nằm lại trong hàng chờ = lỗi dữ liệu -> nhớ để không lặp lại mỗi 15 giây
    con_lai = {p for p, _ in file_cho()}
    for x in moi:
        if x[0] in con_lai:
            da_loi.add(x)
    canh_bao = [ln.strip() for ln in out.splitlines() if ln.strip().startswith(("!", "***"))]
    if tao:
        ten = ", ".join(os.path.basename(p) for p in tao)
        thong_bao(f"Đã tạo {len(tao)} phiếu YCPT" + (f" — {len(canh_bao)} cảnh báo" if canh_bao else ""),
                  ten + (" | Xem NHAT_KY_TU_DONG.txt" if canh_bao else ""))
    elif any(x in da_loi for x in moi):
        thong_bao("Chưa tạo được phiếu YCPT", "Thiếu/sai dữ liệu — xem NHAT_KY_TU_DONG.txt")


def main():
    if "--mot-lan" in sys.argv:
        chay_mot_lan(set())
        return
    # Chỉ 1 bản chạy ngầm cùng lúc
    lock = socket.socket()
    try:
        lock.bind(("127.0.0.1", 47613))
    except OSError:
        return
    ghi_log(f"\n=== Bắt đầu chạy ngầm {datetime.datetime.now():%d/%m/%Y %H:%M} — theo dõi: "
            f"{'; '.join(tao_phieu.queue_dirs())}")
    da_loi = set()
    while True:
        try:
            chay_mot_lan(da_loi)
        except Exception as ex:
            ghi_log(f"Lỗi chạy ngầm: {ex}")
        time.sleep(CHU_KY)


if __name__ == "__main__":
    main()
