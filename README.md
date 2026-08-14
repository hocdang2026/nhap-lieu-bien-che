# HỆ THỐNG NHẬP LIỆU BIÊN CHẾ 2026

Ứng dụng web Python/Flask dùng cho 80 đơn vị trong sheet **CBCC** của file Excel mẫu.

## Chức năng
- 1 tài khoản riêng cho mỗi đơn vị.
- Tài khoản chỉ được nhập dữ liệu của chính đơn vị đó.
- Bắt buộc đổi mật khẩu ở lần đăng nhập đầu.
- Khóa tạm 15 phút sau 5 lần nhập sai mật khẩu.
- Admin xem trạng thái đã nhập/chưa nhập, khóa tài khoản, đặt lại mật khẩu.
- Ghi nhật ký cập nhật số liệu.
- Sau mỗi lần lưu, hệ thống xây dựng lại `data/current_output.xlsx`.
- Admin tải Excel tổng hợp từ `/admin/export.xlsx`.

## Chạy trên Windows
1. Cài Python 3.11 hoặc 3.12.
2. Mở CMD trong thư mục này.
3. Chạy:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

4. Mở trình duyệt: `http://127.0.0.1:5000`

## Admin ban đầu
- Tên đăng nhập: `admin`
- Mật khẩu tạm: `Admin@2026!`
- Hệ thống sẽ yêu cầu đổi ngay ở lần đăng nhập đầu.

## Tài khoản đơn vị
Lần chạy đầu tiên, ứng dụng tạo 80 tài khoản và ghi mật khẩu tạm vào:

`data/initial_credentials.csv`

Hãy lưu file này ở nơi an toàn và **không gửi toàn bộ cho các đơn vị**. Chỉ gửi đúng tài khoản/mật khẩu tương ứng từng đơn vị.

## Đưa lên Internet không cần tên miền
Có thể đưa thư mục này lên Render. Sau khi deploy, Render cấp một URL dạng:

`https://nhap-lieu-bien-che-2026.onrender.com`

Tất cả đơn vị dùng chung URL đăng nhập, nhưng tài khoản riêng sẽ giới hạn dữ liệu theo đơn vị.

### Quan trọng về dữ liệu
Nếu dùng hosting không có ổ đĩa lưu bền vững, SQLite có thể mất sau khi redeploy/reset. Với hệ thống dùng chính thức, nên dùng persistent disk hoặc PostgreSQL. Bản hiện tại ưu tiên dễ chạy và thử nghiệm trước.

## Cấu trúc dữ liệu Excel
Với mỗi đơn vị:
- C: Biên chế tạm giao
- D: Biên chế có mặt
- E: Biên chế giao
- F: `=E-C`
- G: `=E-D`
- H: Ghi chú

Các công thức tổng hợp sẵn có trong file mẫu được giữ lại.

## Cách nhanh nhất trên Windows
Nhấp đúp `BAT_DAU_CHAY_WINDOWS.bat`. Lần đầu máy cần Internet để tải các thư viện Python.

## File phát tài khoản
`TAI_KHOAN_BAN_DAU.csv` chứa 80 tài khoản và mật khẩu tạm. Sau khi triển khai và phát tài khoản xong, hãy cất file này ở nơi an toàn, không đặt trong thư mục public và không đưa lên repository công khai.

## Chức năng mới: Admin tự chọn dữ liệu cần nhập

Trong trang **Admin**, khu vực **CẤU HÌNH DỮ LIỆU CẦN NHẬP** cho phép quản trị viên:

- Tick **Cho đơn vị nhập** để bật/tắt từng trường trên phiếu nhập liệu.
- Tick **Bắt buộc** để không cho đơn vị lưu khi trường đó còn trống.
- Các trường hiện hỗ trợ tương ứng file tổng CBCC:
  - Cột C: Biên chế tạm giao năm 2026
  - Cột D: Biên chế có mặt tính đến ngày 15/6/2026
  - Cột E: Biên chế giao năm 2026
  - Cột H: Ghi chú
- Cột F và G vẫn là cột tự tính, không cho đơn vị nhập trực tiếp.
- Khi Admin thay đổi cấu hình, phiếu nhập của tất cả đơn vị tự thay đổi theo ngay mà không cần tạo lại tài khoản.
- Nếu một trường bị tắt sau khi đã có dữ liệu, dữ liệu cũ được giữ lại, không bị xóa.
