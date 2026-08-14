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

## V3 - Tải và tự nhận dạng file Excel mới

Admin có thêm nút **TẢI / NHẬN DẠNG FILE EXCEL**.

Quy trình:
1. Tải file `.xlsx`, `.xlsm` hoặc `.xls` lên.
2. Hệ thống tự dò sheet, dòng tiêu đề và cột có khả năng là tên đơn vị.
3. Admin chọn đúng sheet và có thể sửa lại dòng tiêu đề/cột tên đơn vị.
4. Admin tick các cột muốn đơn vị nhập, đặt tên hiển thị, kiểu Số hoặc Chữ/Ghi chú và đánh dấu Bắt buộc nếu cần.
5. Bấm **ÁP DỤNG**. Chỉ lúc này hệ thống mới đổi mẫu đang dùng.
6. Các đơn vị trùng tên giữ nguyên tài khoản/mật khẩu; đơn vị mới được tạo tài khoản mới.
7. Dữ liệu có sẵn trong file ở các cột được chọn sẽ được đọc vào form. Công thức Excel không được coi là dữ liệu nhập thủ công.
8. Khi đơn vị nhập, dữ liệu được ghi lại đúng dòng/cột của mẫu mới và Admin có thể tải file tổng hợp.

### Định dạng hỗ trợ
- `.xlsx`: hỗ trợ chính, giữ cấu trúc/định dạng tốt nhất.
- `.xlsm`: hỗ trợ và giữ VBA khi đọc/ghi bằng openpyxl ở mức workbook hỗ trợ.
- `.xls`: hệ thống chuyển dữ liệu sang `.xlsx`; dữ liệu ô được giữ nhưng định dạng/công thức phức tạp của định dạng cũ có thể bị mất.

### Lưu ý khi đổi mẫu
Áp dụng một mẫu Excel mới sẽ đặt lại trạng thái Đã nhập/Chưa nhập của đợt hiện tại. Nên tải file tổng hợp cũ để lưu trữ trước khi thay mẫu.

# V5 - Cập nhật phiên bản không làm mất dữ liệu

Bản V5 tách **mã nguồn** và **dữ liệu vận hành** thành hai phần độc lập.

## Những dữ liệu phải được giữ nguyên qua các lần cập nhật
- Database `app.db`: tài khoản, mật khẩu đã băm, trạng thái khóa/mở, cấu hình, dữ liệu đã nhập, lịch sử chỉnh sửa.
- File Excel mẫu đang áp dụng: `active_template.xlsx` hoặc `active_template.xlsm`.
- File Excel tổng hợp hiện tại.
- Các file hỗ trợ trong thư mục dữ liệu.

## Quy tắc quan trọng
Ứng dụng chỉ chép file mẫu/tài khoản seed vào `DATA_DIR` **khi file đó chưa tồn tại**. Khi cập nhật code mới, ứng dụng không ghi đè database hoặc file Excel đã có.

## Chạy trên Render với dữ liệu bền vững
Render Free dùng filesystem tạm thời nên không phù hợp để giữ SQLite/file upload lâu dài. Nếu dùng kiến trúc SQLite + Excel của phần mềm này, hãy dùng Web Service trả phí và gắn Persistent Disk.

Khuyến nghị:
- Disk mount path: `/var/data`
- Environment variable: `DATA_DIR=/var/data/nhaplieu`
- Environment variable: `PERSISTENT_STORAGE=1`
- Giữ `SECRET_KEY` cố định trong Render Environment.

Chỉ dữ liệu nằm bên dưới mount path của Persistent Disk mới được giữ qua deploy/restart.

## Sao lưu / Khôi phục trong Admin
Admin có khu vực **Dữ liệu & sao lưu khi cập nhật**:
- **TẢI BẢN SAO LƯU**: xuất một file ZIP chứa database + file Excel đang dùng.
- **KHÔI PHỤC BẢN SAO LƯU**: phục hồi toàn bộ trạng thái hệ thống từ file ZIP.

Nên tải một bản sao lưu trước khi:
1. Đổi cấu trúc Excel lớn.
2. Nâng cấp phiên bản phần mềm.
3. Thay đổi hosting hoặc Render service.
4. Thực hiện thao tác dữ liệu hàng loạt.

## Cập nhật V5/V6/V7 về sau
Chỉ thay code (`app.py`, `templates`, `static`, `requirements.txt`...) trên GitHub. Không đưa `app.db`, `active_template.xlsx`, `uploads/`, `backups/` từ máy cá nhân lên ghi đè Persistent Disk.
