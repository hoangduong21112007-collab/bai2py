import re
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

THU_MUC = Path(__file__).resolve().parent
FILE = THU_MUC / "Thuc_hanh_Python_Ban_hang.xlsx"
FILE_LOG = THU_MUC / "log_xu_ly.txt"
FILE_BAO_CAO = THU_MUC / "BaoCaoBanHang.xlsx"
FILE_DON_MOI = THU_MUC / "DonHang_Moi.xlsx"
PHUONG_THUC = {"Tiền mặt", "Chuyển khoản", "Thẻ"}

class DonHangKhongHopLeError(Exception):
    pass

def kiem_tra_don_hang(so_luong, giam_gia):
    if so_luong <= 0:
        raise DonHangKhongHopLeError(f"Số lượng phải > 0 (đang là {so_luong:g})")
    if not 0 <= giam_gia <= 1:
        raise DonHangKhongHopLeError(f"Giảm giá phải trong [0, 1] (đang là {giam_gia:g})")

def ghi_log(noi_dung):
    with open(FILE_LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {noi_dung}\n")

def tien(df, *cot):
    return df.assign(**{c: df[c].map("{:,.0f}".format) for c in cot})

def doc_excel(duong_dan, ten_sheet):
    try:
        return pd.read_excel(duong_dan, sheet_name=ten_sheet)
    except FileNotFoundError:
        loi = f"Không tìm thấy file {duong_dan}"
    except ValueError:
        loi = f"Sheet '{ten_sheet}' không tồn tại"
    except Exception as e:
        loi = f"Lỗi khi đọc file: {e}"
    print("[LỖI]", loi)
    ghi_log(loi)
    return None


def kiem_tra_dong(r, ds_sp, ds_ch, trung):
    loi = []
    if trung:
        loi.append("Mã đơn hàng trùng")
    try:
        ngay = pd.to_datetime(r["NgayBan"], format="%d/%m/%Y")
    except (ValueError, TypeError):
        ngay = pd.NaT
        loi.append(f"Ngày bán không hợp lệ ({r['NgayBan']})")
    sl = pd.to_numeric(r["SoLuong"], errors="coerce")
    giam = pd.to_numeric(r["TyLeGiamGia"], errors="coerce")
    if pd.isna(sl):
        loi.append(f"Số lượng trống/không phải số ({r['SoLuong']})")
    if pd.isna(giam):
        loi.append(f"Giảm giá trống/không phải số ({r['TyLeGiamGia']})")
    if pd.notna(sl) and pd.notna(giam):
        try:
            kiem_tra_don_hang(sl, giam)
        except DonHangKhongHopLeError as e:
            loi.append(str(e))
    if pd.isna(r["MaSP"]):
        loi.append("Mã sản phẩm trống")
    elif r["MaSP"] not in ds_sp:
        loi.append(f"Mã sản phẩm {r['MaSP']} không có trong danh mục")
    if pd.isna(r["MaCH"]) or r["MaCH"] not in ds_ch:
        loi.append(f"Mã cửa hàng {r['MaCH']} không có trong danh mục")
    if pd.isna(r["KhachHang"]):
        loi.append("Khách hàng trống")
    if pd.isna(r["SoDienThoai"]):
        loi.append("Số điện thoại trống")
    elif not re.fullmatch(r"0\d{9}", str(r["SoDienThoai"])):
        loi.append(f"Số điện thoại không hợp lệ ({r['SoDienThoai']})")
    if r["PhuongThucTT"] not in PHUONG_THUC:
        loi.append(f"Phương thức thanh toán không hợp lệ ({r['PhuongThucTT']})")
    return loi, ngay, sl, giam

def xu_ly(dh, sp, ch):
    ds_sp, ds_ch = set(sp["MaSP"]), set(ch["MaCH"])
    trung = dh["MaDH"].duplicated()
    ly_do, ngay, sl, giam = [], [], [], []
    for i, r in dh.iterrows():
        a, b, c, d = kiem_tra_dong(r, ds_sp, ds_ch, trung[i])
        ly_do.append("; ".join(a)); ngay.append(b); sl.append(c); giam.append(d)
    loi_mask = pd.Series([bool(x) for x in ly_do], index=dh.index)

    loi = dh[loi_mask].copy()
    loi["LyDoLoi"] = [x for x in ly_do if x]

    hop_le = dh[~loi_mask].copy()
    hop_le["NgayBan"] = [x for x, m in zip(ngay, loi_mask) if not m]
    hop_le["SoLuong"] = [int(x) for x, m in zip(sl, loi_mask) if not m]
    hop_le["TyLeGiamGia"] = [x for x, m in zip(giam, loi_mask) if not m]
    hop_le = hop_le.merge(sp[["MaSP", "TenSP", "NhomHang", "DonGia"]], on="MaSP").merge(ch, on="MaCH")
    hop_le["ThanhTien"] = hop_le["SoLuong"] * hop_le["DonGia"] * (1 - hop_le["TyLeGiamGia"])
    hop_le["Thang"] = hop_le["NgayBan"].dt.strftime("%Y-%m")
    return {"dh": dh, "sp": sp.copy(), "ch": ch, "hop_le": hop_le.reset_index(drop=True),
            "loi": loi.reset_index(drop=True), "don_moi": []}


def nap_du_lieu():
    try:
        dh, sp, ch = (doc_excel(FILE, s) for s in ("DonHang", "SanPham", "CuaHang"))
        if dh is None or sp is None or ch is None:
            raise ValueError("Không đọc đủ 3 sheet")
        du_lieu = xu_ly(dh, sp, ch)
    except Exception as e:
        print("[LỖI] Không đọc/kiểm tra được dữ liệu:", e)
        ghi_log(f"Đọc/kiểm tra thất bại: {e}")
        return None
    else:
        print(f"Đã đọc {len(dh)} đơn: {len(du_lieu['hop_le'])} hợp lệ, {len(du_lieu['loi'])} lỗi")
        return du_lieu
    finally:
        ghi_log("Kết thúc bước đọc và kiểm tra")


def kham_pha(d):
    for ten in ("dh", "sp", "ch"):
        df = d[ten]
        print(f"\n--- {ten} ---\n{df.head().to_string(index=False)}")
        print("Kích thước:", df.shape)
        print("Tên cột:", list(df.columns))
        print(df.dtypes.to_string())
    dh = d["dh"]
    thang = pd.to_datetime(dh["NgayBan"], format="%d/%m/%Y", errors="coerce").dt.strftime("%Y-%m")
    print("\nGiao dịch theo cửa hàng:\n", dh["MaCH"].value_counts().to_string())
    print("\nTheo phương thức thanh toán:\n", dh["PhuongThucTT"].value_counts().to_string())
    print("\nTheo tháng:\n", thang.fillna("Không rõ").value_counts().sort_index().to_string())
    ghep = dh.merge(d["sp"], on="MaSP", how="left").merge(d["ch"], on="MaCH", how="left")
    print("\nDữ liệu sau khi ghép:\n", ghep[["MaDH", "TenSP", "NhomHang", "DonGia",
                                            "TenCuaHang", "TinhThanh"]].head().to_string(index=False))


def xem_loi(d):
    dh, loi = d["dh"], d["loi"]
    print("Các dòng trùng mã:\n", dh[dh["MaDH"].duplicated(keep=False)].to_string(index=False))
    print("\nCác dòng lỗi và lý do:\n", loi[["MaDH", "LyDoLoi"]].to_string(index=False))
    print(f"\nHợp lệ: {len(d['hop_le'])} | Lỗi: {len(loi)} | Gốc: {len(dh)} (không bị sửa)")


def doanh_thu(hl, cot):
    return (hl.groupby(cot).agg(SoDon=("MaDH", "count"), DoanhThu=("ThanhTien", "sum"))
            .reset_index().sort_values("DoanhThu", ascending=False))


def top10(hl):
    t = hl.sort_values("ThanhTien", ascending=False).head(10)
    return t[["MaDH", "NgayBan", "TenCuaHang", "TenSP", "SoLuong", "ThanhTien"]]


def xem_doanh_thu(d):
    hl = d["hop_le"]
    print(f"Tổng doanh thu: {hl['ThanhTien'].sum():,.0f} VNĐ")
    for nhan, cot in (("cửa hàng", "TenCuaHang"), ("tỉnh/thành", "TinhThanh"),
                      ("nhóm hàng", "NhomHang"), ("tháng", "Thang")):
        print(f"\nDoanh thu theo {nhan}:\n", tien(doanh_thu(hl, cot), "DoanhThu").to_string(index=False))
    print("\nTop 10 đơn hàng:\n", tien(top10(hl), "ThanhTien").to_string(index=False))
    sp = doanh_thu(hl, "TenSP").iloc[0]
    print(f"\nSản phẩm doanh thu cao nhất: {sp['TenSP']} ({sp['DoanhThu']:,.0f} VNĐ)")

def ve_bieu_do(d):
    hl = d["hop_le"]
    ch, nhom = doanh_thu(hl, "TenCuaHang"), doanh_thu(hl, "NhomHang")
    plt.figure(figsize=(9, 5))
    plt.bar(ch["TenCuaHang"], ch["DoanhThu"] / 1e6)
    plt.title("Doanh thu theo cửa hàng (triệu VNĐ)")
    plt.xticks(rotation=15)
    plt.tight_layout()
    plt.savefig(THU_MUC / "bieu_do_cot.png")
    plt.figure(figsize=(7, 7))
    plt.pie(nhom["DoanhThu"], labels=nhom["NhomHang"], autopct="%1.1f%%")
    plt.title("Tỷ trọng doanh thu theo nhóm hàng")
    plt.savefig(THU_MUC / "bieu_do_tron.png")
    print("Đã lưu bieu_do_cot.png và bieu_do_tron.png")
    plt.show()

def xuat_bao_cao(d):
    hl = d["hop_le"]
    try:
        with pd.ExcelWriter(FILE_BAO_CAO) as w:
            hl.drop(columns="Thang").assign(NgayBan=hl["NgayBan"].dt.strftime("%d/%m/%Y")) \
                .to_excel(w, sheet_name="DonHang_HopLe", index=False)
            d["loi"].to_excel(w, sheet_name="DonHang_Loi", index=False)
            doanh_thu(hl, "TenCuaHang").to_excel(w, sheet_name="DoanhThu_CuaHang", index=False)
            doanh_thu(hl, "NhomHang").to_excel(w, sheet_name="DoanhThu_NhomHang", index=False)
            top10(hl).assign(NgayBan=lambda t: t["NgayBan"].dt.strftime("%d/%m/%Y")) \
                .to_excel(w, sheet_name="Top10_DonHang", index=False)
    except PermissionError:
        print("[LỖI] Hãy đóng file BaoCaoBanHang.xlsx trong Excel rồi thử lại")
    except Exception as e:
        print("[LỖI] Xuất báo cáo thất bại:", e)
        ghi_log(f"Xuất báo cáo thất bại: {e}")
    else:
        print("Đã xuất", FILE_BAO_CAO.name)
    finally:
        ghi_log("Kết thúc bước xuất báo cáo")

def nhap_san_pham(sp):
    while True:
        ma = input("Mã sản phẩm (Enter = hủy): ").strip().upper()
        if not ma:
            return None
        try:
            dong = sp[sp["MaSP"] == ma]
            if dong.empty:
                raise KeyError(ma)
            return dong.iloc[0]
        except KeyError:
            print("Mã không tồn tại")

def nhap_so_luong(ton):
    while True:
        s = input("Số lượng (Enter = hủy): ").strip()
        if not s:
            return None
        try:
            sl = int(s)
            kiem_tra_don_hang(sl, 0)
            if sl > ton:
                raise OverflowError
            return sl
        except ValueError:
            print("Số lượng phải là số nguyên")
        except DonHangKhongHopLeError as e:
            print(e)
        except OverflowError:
            print(f"Vượt tồn kho (chỉ còn {ton})")

def them_don_hang(d):
    sp_chon = nhap_san_pham(d["sp"])
    if sp_chon is None:
        return
    ton = int(sp_chon["TonKho"])
    print(f"{sp_chon['TenSP']} - {sp_chon['DonGia']:,.0f} VNĐ - tồn kho {ton}")
    sl = nhap_so_luong(ton)
    if sl is None:
        return
    don = {"MaDH": f"DH{len(d['dh']) + len(d['don_moi']) + 1:03d}",
           "NgayBan": datetime.now().strftime("%d/%m/%Y"),
           "MaCH": input("Mã cửa hàng: ").strip().upper(), "MaSP": sp_chon["MaSP"],
           "KhachHang": input("Khách hàng: ").strip() or None,
           "SoDienThoai": input("Số điện thoại: ").strip() or None, "SoLuong": sl,
           "TyLeGiamGia": input("Giảm giá (0-1, Enter = 0): ").strip() or 0,
           "PhuongThucTT": input("Phương thức (Tiền mặt/Chuyển khoản/Thẻ): ").strip()}
    moi = xu_ly(pd.DataFrame([don]), d["sp"], d["ch"])
    if len(moi["loi"]):
        print("Đơn không hợp lệ:", moi["loi"]["LyDoLoi"].iloc[0])
        return
    d["hop_le"] = pd.concat([d["hop_le"], moi["hop_le"]], ignore_index=True)
    d["sp"].loc[d["sp"]["MaSP"] == don["MaSP"], "TonKho"] -= sl
    d["don_moi"].append(don)
    pd.DataFrame(d["don_moi"]).to_excel(FILE_DON_MOI, index=False)
    print(f"Đã thêm đơn {don['MaDH']}: {moi['hop_le']['ThanhTien'].iloc[0]:,.0f} VNĐ")

def tim_don_hang(d):
    tu = input("Nhập mã đơn / khách hàng / sản phẩm / cửa hàng: ").strip()
    hl = d["hop_le"]
    chuoi = hl["MaDH"] + " " + hl["KhachHang"] + " " + hl["TenSP"] + " " + hl["TenCuaHang"]
    kq = hl[chuoi.str.contains(tu, case=False, regex=False)]
    if kq.empty:
        print("Không tìm thấy")
        return
    cot = ["MaDH", "NgayBan", "TenCuaHang", "TenSP", "KhachHang", "SoLuong", "ThanhTien"]
    print(tien(kq[cot], "ThanhTien").to_string(index=False))

def kiem_tra_ton_kho(d):
    sp = d["sp"][["MaSP", "TenSP", "TonKho"]].sort_values("TonKho")
    sp["CanhBao"] = sp["TonKho"].map(lambda t: "SẮP HẾT" if t <= 20 else "")
    print(sp.to_string(index=False))

def doc_tat_ca_file(d):
    cac_bang = []
    for f in sorted(THU_MUC.glob("*.xlsx")):
        if f.name in (FILE_BAO_CAO.name, FILE_DON_MOI.name) or f.name.startswith("~$"):
            continue
        try:
            cac_bang.append(pd.read_excel(f, sheet_name="DonHang"))
            print("OK  ", f.name)
        except Exception as e:
            print("LỖI ", f.name, "(đã ghi vào log_xu_ly.txt)")
            ghi_log(f"File {f.name} bị bỏ qua: {type(e).__name__}: {e}")
    if cac_bang:
        gop = pd.concat(cac_bang, ignore_index=True)
        print(f"Đã ghép {len(cac_bang)} file, {len(gop)} dòng")
        return xu_ly(gop, d["sp"], d["ch"])
    return d

MENU = [("Khám phá dữ liệu", kham_pha), ("Xem dòng lỗi và lý do", xem_loi),
        ("Xem doanh thu", xem_doanh_thu), ("Vẽ biểu đồ", ve_bieu_do),
        ("Tìm đơn hàng", tim_don_hang), ("Thêm đơn hàng", them_don_hang),
        ("Kiểm tra tồn kho", kiem_tra_ton_kho), ("Xuất báo cáo", xuat_bao_cao),
        ("Đọc và ghép tất cả file .xlsx", None)]

def main():
    d = nap_du_lieu()
    if d is None:
        return
    while True:
        print("\n===== QUẢN LÝ BÁN HÀNG =====")
        for i, (ten, _) in enumerate(MENU, 1):
            print(f"{i}. {ten}")
        print("0. Thoát")
        c = input("Chọn: ").strip()
        if c == "0":
            break
        if not (c.isdigit() and 1 <= int(c) <= len(MENU)):
            print("Lựa chọn không hợp lệ")
            continue
        try:
            ham = MENU[int(c) - 1][1]
            if ham is None:
                d = doc_tat_ca_file(d)
            else:
                ham(d)
        except Exception as e:
            print(f"[LỖI] {type(e).__name__}: {e}")
            ghi_log(f"Lỗi menu: {type(e).__name__}: {e}")

if __name__ == "__main__":
    main()