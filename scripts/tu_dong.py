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
KQ_DIR = os.path.join(ROOT, "KET_QUA")
CHU_KY = 15          # giây giữa 2 lần kiểm tra
CHO_ON_DINH = 5      # file phải đứng yên 5 giây (Drive đồng bộ xong) mới xử lý
WIN = sys.platform == "win32"
NO_WINDOW = 0x08000000 if WIN else 0
LOCK = None
# File công cụ được phép cập nhật qua CAP_NHAT (tên phẳng: thư mục con nối bằng "__", vd scripts__tao_phieu.py)
CHO_PHEP_CAP_NHAT = ("scripts/", "mau_goc/")
DUOI_CAP_NHAT = (".py", ".json", ".ps1", ".docx", ".txt", ".bat")


def ghi_log(text):
    try:
        if os.path.exists(LOG) and os.path.getsize(LOG) > 150_000:   # giữ nhật ký gọn
            with open(LOG, encoding="utf-8", errors="ignore") as f:
                tail = f.read()[-50_000:]
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


def thu_muc_cap_nhat():
    ds = [os.path.join(ROOT, "CAP_NHAT")]
    d = tao_phieu.drive_dir()
    if d:
        ds.append(os.path.join(d, "CAP_NHAT"))
    return [x for x in ds if os.path.isdir(x)]


def ap_dung_cap_nhat():
    """Chép file công cụ mới từ CAP_NHAT (laptop hoặc Drive) đè lên bản đang dùng, ghi md5 vào KET_QUA\\CAP_NHAT.txt.
    Trả về True nếu chính tu_dong.py vừa được cập nhật (cần khởi động lại)."""
    import hashlib
    import shutil
    now, xong, tu_khoi_dong = time.time(), [], False
    for d in thu_muc_cap_nhat():
        for f in sorted(os.listdir(d)):
            src = os.path.join(d, f)
            if not os.path.isfile(src) or f.startswith(("~", ".")) or now - os.path.getmtime(src) < CHO_ON_DINH:
                continue
            rel = f.replace("__", "/")
            ok = (".." not in rel and rel.lower().endswith(DUOI_CAP_NHAT)
                  and (rel.startswith(CHO_PHEP_CAP_NHAT) or "/" not in rel))
            if not ok:
                xong.append(f"!  Bỏ qua {f} (tên không hợp lệ — dạng đúng: scripts__ten.py)")
                continue
            dst = os.path.join(ROOT, *rel.split("/"))
            try:
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copyfile(src, dst)
                with open(dst, "rb") as fh:
                    md5 = hashlib.md5(fh.read()).hexdigest()
                os.remove(src)
                xong.append(f"{rel}  md5={md5}")
                if rel == "scripts/tu_dong.py":
                    tu_khoi_dong = True
            except Exception as ex:
                xong.append(f"!  Lỗi cập nhật {rel}: {ex} (file đang mở?) — sẽ thử lại")
    if xong:
        txt = "\n".join(xong)
        ghi_log(f"\n=== Cập nhật công cụ {datetime.datetime.now():%d/%m/%Y %H:%M:%S}\n{txt}")
        try:
            os.makedirs(KQ_DIR, exist_ok=True)
            with open(os.path.join(KQ_DIR, "CAP_NHAT.txt"), "w", encoding="utf-8") as fh:
                fh.write(f"{datetime.datetime.now():%d/%m/%Y %H:%M:%S}\n{txt}\n")
        except Exception:
            pass
        thong_bao("Đã cập nhật công cụ phiếu YCPT", f"{len(xong)} file")
    return tu_khoi_dong


def khoi_dong_lai():
    """tu_dong.py vừa được thay: nhả khóa, mở bản mới chạy ngầm, thoát bản cũ (không cần bấm TAT/CAI)."""
    global LOCK
    if LOCK:
        LOCK.close()
        LOCK = None
    py = sys.executable
    flags = (0x00000008 | NO_WINDOW) if WIN else 0     # DETACHED_PROCESS
    subprocess.Popen([py, os.path.abspath(__file__)], cwd=ROOT, creationflags=flags, close_fds=True)
    os._exit(0)


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


def ghi_ket_qua(ten_json, text):
    """KET_QUA\\<tên json>.txt — kết quả GỌN của đúng 1 lô (Claude chỉ đọc file nhỏ này, không đọc cả nhật ký)."""
    try:
        os.makedirs(KQ_DIR, exist_ok=True)
        han = time.time() - 30 * 86400            # dọn file kết quả cũ hơn 30 ngày
        for f in os.listdir(KQ_DIR):
            p = os.path.join(KQ_DIR, f)
            if os.path.isfile(p) and os.path.getmtime(p) < han:
                os.remove(p)
        stem = os.path.splitext(os.path.basename(ten_json))[0]
        with open(os.path.join(KQ_DIR, stem + ".txt"), "w", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%d/%m/%Y %H:%M:%S}\n{text.strip()}\n")
    except Exception:
        pass


def _python():
    py = sys.executable
    if WIN and py.lower().endswith("pythonw.exe"):
        py = py[:-len("pythonw.exe")] + "python.exe"
    return py


def chay_mot_lan(da_loi):
    """Tạo phiếu cho mọi file mới — từng file một, mỗi file 1 kết quả riêng.
    da_loi: {(đường dẫn, mtime)} file lỗi đã báo — không báo lại tới khi file được sửa."""
    moi = [x for x in file_cho() if x not in da_loi]
    if not moi:
        return
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    ds = os.path.join(HERE, "_danh_sach_file_moi.txt")
    out, tao, tung_lo = "", [], []
    for p, _ in moi:
        try:
            r = subprocess.run([_python(), os.path.join(HERE, "tao_phieu.py"), p], cwd=ROOT, env=env,
                               capture_output=True, text=True, encoding="utf-8", errors="replace",
                               creationflags=NO_WINDOW, timeout=300)
            o = r.stdout + r.stderr
        except subprocess.TimeoutExpired:
            o = "*** Tạo phiếu quá 5 phút — đã dừng. Chạy tay TAO_PHIEU.bat để xem lỗi."
        moi_tao = []
        if os.path.exists(ds):
            with open(ds, encoding="utf-8") as f:
                moi_tao = [ln.strip() for ln in f if ln.strip()]
        tao += moi_tao
        tung_lo.append((p, o))
        out += o + "\n"
    with open(ds, "w", encoding="utf-8") as f:     # Word làm sạch tất cả phiếu vừa tạo trong 1 lần
        f.write("\n".join(tao))
    tao = [ln.split("|")[0] for ln in tao]
    word_out = ""
    if tao and WIN:
        try:
            w = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                                "-File", os.path.join(HERE, "lam_sach_bang_word.ps1"), ds],
                               cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace",
                               creationflags=NO_WINDOW, timeout=180)
            word_out = w.stdout + w.stderr
        except subprocess.TimeoutExpired:
            # Word treo (thường do hộp thoại kích hoạt/khôi phục file). KHÔNG tắt Word bằng lệnh —
            # có thể làm mất văn bản Vinh đang soạn. Chỉ báo để mở tay kiểm tra.
            word_out = ("!!! Word không phản hồi sau 3 phút (có thể đang hiện hộp thoại) — phiếu ĐÃ tạo nhưng chưa "
                        "được Word làm sạch: mở từng file bằng Word, lưu lại trước khi gửi. Nếu có hộp thoại Word đang chờ, bấm đóng nó.")
        out += "\n" + word_out
    ghi_log(f"\n######## {datetime.datetime.now():%d/%m/%Y %H:%M:%S} ########\n{out}")
    word_canh_bao = "\n".join(ln for ln in word_out.splitlines() if ln.strip().startswith(("!", "***")))
    for p, o in tung_lo:
        ghi_ket_qua(p, tao_phieu.loc_gon(o) + ("\n" + word_canh_bao if word_canh_bao else ""))

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
    global LOCK
    if "--mot-lan" in sys.argv:
        ap_dung_cap_nhat()
        chay_mot_lan(set())
        return
    # Chỉ 1 bản chạy ngầm cùng lúc (bản vừa khởi động lại có thể phải chờ bản cũ nhả khóa)
    LOCK = socket.socket()
    for _ in range(10):
        try:
            LOCK.bind(("127.0.0.1", 47613))
            break
        except OSError:
            time.sleep(1)
    else:
        return
    ghi_log(f"\n=== Bắt đầu chạy ngầm {datetime.datetime.now():%d/%m/%Y %H:%M} — theo dõi: "
            f"{'; '.join(tao_phieu.queue_dirs())}")
    da_loi = set()
    while True:
        try:
            if ap_dung_cap_nhat():
                khoi_dong_lai()
            chay_mot_lan(da_loi)
        except Exception as ex:
            ghi_log(f"Lỗi chạy ngầm: {ex}")
        time.sleep(CHU_KY)


if __name__ == "__main__":
    main()
