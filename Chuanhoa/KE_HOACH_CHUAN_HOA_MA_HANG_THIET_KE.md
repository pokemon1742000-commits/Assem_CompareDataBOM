# Kế Hoạch Phần Mềm Chuẩn Hóa Mã Hàng Thiết Kế

## Mục tiêu

Xây dựng phần mềm Electron version `2.0.0` để chuẩn hóa mã hàng thiết kế bằng cách so sánh file **Dữ Liệu Thiết Kế/BOM** với danh sách **Mã Đã Đặt Hàng**. Phần mềm tập trung vào thao tác Excel, tự động phân loại kết quả và hỗ trợ người dùng xác nhận các trường hợp mã gần giống hoặc sai đơn vị tính.

## Phạm vi hiện tại

- Ứng dụng desktop Electron chạy bằng `npm start` hoặc bản build Windows.
- Nhập dữ liệu từ `.xlsx`, `.xls`, `.xlsm`, `.csv`.
- Hỗ trợ load nhiều file và chọn nhiều sheet.
- Lưu danh sách file đã load gần nhất để tự khôi phục khi mở lại app.
- Hiển thị dữ liệu theo các tab: **Bảng thống kê**, **Dữ Liệu Mã**, **Dữ Liệu Thiết Kế**, **So Sánh**, **Mã Mới**.
- Xuất Excel theo mẫu **Đề Nghị Mua Hàng**, gồm cả mã giống nhau và mã mới.
- Kiểm tra update trên bản exe đã build.

## Nguồn dữ liệu

### 1. Mã Đã Đặt Hàng

Cột bắt buộc là cột mã hàng. Các tiêu đề được nhận diện:

- `Mã hàng tồn kho`
- `Mã hàng đặt hàng`
- `Mã đặt hàng`
- `Part No`, `Part Number`, `Item Code`, `Code`

Cột bổ sung:

- Tên hàng: `Tên hàng`, `Tên mặt hàng`, `Item Name`, `Name`
- Đơn vị tính: `Đơn vị tính`, `DVT`, `Unit`, `Units`, `UOM`, `Unit of Measure`

Kết quả hiển thị trong tab **Dữ Liệu Mã**:

- `STT`
- `Mã bản vẽ`
- `Tên hàng`
- `Đơn vị tính`

### 2. Dữ Liệu Thiết Kế

Cột bắt buộc:

- Mã bản vẽ hoặc `Model`
- Số lượng/máy

Cột bổ sung:

- Tên mặt hàng
- Nhà sản xuất
- Đơn vị tính

Các tiêu đề được nhận diện trong 20 dòng đầu của file. Nếu không tìm được cột mã bản vẽ hoặc số lượng/máy, app báo lỗi và không load file đó.

Kết quả hiển thị trong tab **Dữ Liệu Thiết Kế**:

- `STT`
- `Tên mặt hàng`
- `Mã bản vẽ`
- `Nhà sản xuất`
- `Số lượng/máy`
- `Đơn vị tính`

## Luồng xử lý

1. Người dùng load một hoặc nhiều file **Mã Đã Đặt Hàng**.
2. Người dùng load một hoặc nhiều file **Dữ Liệu Thiết Kế**.
3. Nếu file có nhiều sheet, người dùng chọn sheet cần load.
4. App lưu đường dẫn file và sheet đã chọn để khôi phục ở lần mở sau.
5. App tự động chạy so sánh khi có đủ hai nguồn dữ liệu.
6. Dữ liệu cùng mã được gộp theo mã chuẩn hóa.
7. Mã được chuẩn hóa bằng cách viết hoa, bỏ khoảng trắng và bỏ đuôi file thiết kế phổ biến.
8. Kết quả được phân loại vào **Giống nhau**, **Cần xác nhận** hoặc **Mã Mới**.
9. Người dùng có thể chỉnh ngưỡng so khớp và app sẽ tính lại kết quả.
10. Người dùng có thể clear riêng dữ liệu mã, clear từng file thiết kế hoặc clear toàn bộ dữ liệu.

## Quy tắc so sánh

- **Trùng tuyệt đối**: mã thiết kế và mã đã đặt hàng có cùng khóa chuẩn hóa.
- **Trùng mã nhưng khác đơn vị tính**: đưa vào bảng **Cần xác nhận** với lý do sai đơn vị tính.
- **Gần giống**: dùng độ giống chuỗi giữa mã thiết kế và mã đã đặt hàng, có tính thêm đơn vị tính nếu cả hai bên đều có đơn vị.
- **Tự ghép**: ngưỡng mặc định `95%`.
- **Cần xác nhận**: ngưỡng mặc định `70%`.
- Mỗi mã cần xác nhận có tối đa 3 đề xuất.
- Mã không có đề xuất đạt ngưỡng được đưa vào **Mã Mới**.
- Với **Mã Mới**, app cố gắng gợi ý mã đã đặt hàng dựa trên độ giống giữa mã thiết kế và tên hàng đã đặt hàng.

## Kết quả trên giao diện

### Bảng thống kê

Hiển thị:

- Số mã đã đặt hàng
- Số dòng dữ liệu thiết kế
- Số mã giống nhau
- Số mã mới
- Số mã khác/cần xử lý
- Số mã cần xác nhận

### Dữ Liệu Mã

Hiển thị danh sách mã đã đặt hàng. Bảng có phân trang khi dữ liệu lớn để tránh quá tải giao diện.

### Dữ Liệu Thiết Kế

Hiển thị dữ liệu thiết kế/BOM đã load và cộng dồn từ nhiều file.

### So Sánh

Gồm hai phần:

- **Bảng Xác Nhận**: mã gần giống hoặc mã trùng nhưng sai đơn vị tính.
- **Bảng Giống Nhau**: mã đã khớp hoặc đã được người dùng xác nhận.

Khi người dùng bấm `Chọn`, kết quả được chuyển xuống bảng **Giống nhau**. Nếu mã hoặc đơn vị tính được sửa, ô sai sẽ bị gạch ngang và ghi chú giá trị đúng.

Khi người dùng bấm `Bỏ qua`, đề xuất bị loại và app chạy lại so sánh.

### Mã Mới

Hiển thị các mã thiết kế chưa tồn tại trong danh sách mã đã đặt hàng, kèm thông tin file/sheet/dòng nguồn trong cột ghi chú.

## Kết quả xuất Excel

File xuất là `DeNghiMuaHang_YYYYMMDD_HHMM.xlsx`, gồm sheet `De Nghi Mua Hang`.

Phần tiêu đề dùng mẫu đề nghị mua hàng MEIKO:

- Logo/tên **MEIKO AUTOMATION**
- `Mã tài liệu`
- Địa chỉ, số điện thoại
- Tiêu đề **ĐỀ NGHỊ MUA HÀNG**
- `Mã dự án`
- `Tên dự án`
- `Số lượng máy`
- Khu vực ký tiếp nhận, duyệt, kiểm tra, lập

Thông tin `Mã dự án`, `Tên dự án`, `Số lượng máy`, `Mã tài liệu`, `Bộ phận` được lấy từ vùng tiêu đề file thiết kế nếu tìm thấy nhãn tương ứng.

Bảng chi tiết xuất các bản vẽ thiết kế từ:

- Bảng **Giống nhau**
- Tab **Mã Mới**

Cột chính:

- `No`
- `Name`
- `Model/ Mã bản vẽ`
- `Maker`
- `Spec`: `Marking part`, `Material`, `Surface`
- `Unit`
- `Q'ty/Machine`
- `Number Machine`
- `Grand Total`
- `Explain`

Nếu còn dòng trong bảng **Cần xác nhận**, app hiện popup: `Vẫn còn mã chưa xác nhận bạn có muốn tiếp tục không?`

- Người dùng chọn **Tiếp tục** thì vẫn xuất file.
- Người dùng chọn **Hủy bỏ** thì hủy xuất file.

## Tiêu chí hoàn thành

- Giao diện chính dùng đúng tên **Chuẩn Hóa Mã Hàng Thiết Kế**.
- Nút nhập dữ liệu chính là **Mã Đã Đặt Hàng** và **Dữ Liệu Thiết Kế**.
- Không còn mô tả luồng so sánh kho/BOM cũ trong README và kế hoạch.
- Tài liệu mô tả đúng các tab, bảng, ngưỡng so khớp, clear dữ liệu, khôi phục file gần đây và xuất Excel đề nghị mua hàng hiện tại.
- Version package là `2.0.0`.
- Build Windows tạo artifact trong `release/`.
