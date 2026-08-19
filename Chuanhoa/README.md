# Chuẩn Hóa Mã Hàng Thiết Kế

Ứng dụng Electron version `2.0.0` dùng để chuẩn hóa mã hàng thiết kế từ file Excel. App so sánh **Dữ Liệu Thiết Kế/BOM** với danh sách **Mã Đã Đặt Hàng**, tự phân loại mã giống nhau, mã cần xác nhận và mã mới.

## Chức năng chính

- Load nhiều file Excel/CSV cho **Mã Đã Đặt Hàng** và **Dữ Liệu Thiết Kế**.
- Chọn sheet cần load khi file có nhiều sheet.
- Tự khôi phục các file đã load ở lần mở app trước nếu file vẫn còn tồn tại.
- Tự động so sánh khi đã có đủ dữ liệu mã đã đặt hàng và dữ liệu thiết kế.
- Dashboard thống kê tổng số mã đã đặt hàng, dòng BOM, mã giống nhau, mã mới và mã cần xác nhận.
- Tìm kiếm nhanh trên dữ liệu đang xem.
- Điều chỉnh ngưỡng so khớp:
  - **Tự ghép**: mặc định `95%`.
  - **Cần xác nhận**: mặc định `70%`.
- Xác nhận thủ công mã gần giống bằng nút `Chọn` hoặc loại đề xuất bằng nút `Bỏ qua`.
- Clear riêng dữ liệu mã đã đặt hàng, clear từng file thiết kế, hoặc clear toàn bộ dữ liệu đã load.
- Chọn theme giao diện.
- Kiểm tra update trên bản đã build exe.
- Xuất file Excel **Đề Nghị Mua Hàng** theo mẫu MEIKO, gồm cả mã giống nhau và mã mới.

## Nguồn dữ liệu

### Mã Đã Đặt Hàng

File Excel cần có cột mã hàng. App tự nhận diện các tên cột phổ biến:

- `Mã hàng tồn kho`
- `Mã hàng đặt hàng`
- `Mã đặt hàng`
- `Part No`, `Part Number`, `Item Code`, `Code`

Các cột bổ sung được nhận diện nếu có:

- `Tên hàng`, `Tên mặt hàng`, `Item Name`, `Name`
- `Đơn vị tính`, `DVT`, `Unit`, `UOM`, `Unit of Measure`

Dữ liệu sau khi load hiển thị ở tab **Dữ Liệu Mã** với các cột:

| STT | Mã bản vẽ | Tên hàng | Đơn vị tính |
| --- | --- | --- | --- |
| 1 | PM113077-A | BRACKET EMG | PCS |

### Dữ Liệu Thiết Kế

App tự nhận diện các cột chính trong 20 dòng đầu:

- Tên mặt hàng: `Tên mặt hàng`, `Name`, `Item Name`, `Part Name`
- Mã bản vẽ: `Mã bản vẽ`, `Model`
- Nhà sản xuất: `Nhà sản xuất`, `Maker`, `Manufacturer`
- Số lượng/máy: `Số lượng`, `Q'ty/Machine`, `Qty/Machine`, `Quantity per Machine`
- Đơn vị tính: `Đơn vị tính`, `DVT`, `Unit`, `UOM`

Dữ liệu thiết kế được cộng dồn khi load thêm nhiều file. Tab **Dữ Liệu Thiết Kế** hiển thị:

| STT | Tên mặt hàng | Mã bản vẽ | Nhà sản xuất | Số lượng/máy | Đơn vị tính |
| --- | --- | --- | --- | --- | --- |

## Luồng so sánh

1. Load **Mã Đã Đặt Hàng**.
2. Load một hoặc nhiều file **Dữ Liệu Thiết Kế**.
3. App chuẩn hóa mã bằng cách viết hoa, bỏ khoảng trắng và bỏ đuôi file thiết kế phổ biến như `.PDF`, `.DWG`, `.DXF`, `.STEP`, `.XLSX`.
4. Nếu mã thiết kế trùng tuyệt đối với mã đã đặt hàng và đơn vị tính khớp, dòng được đưa vào bảng **Giống nhau** trong tab **So Sánh**.
5. Nếu mã trùng nhưng đơn vị tính khác, dòng được đưa vào bảng **Cần xác nhận**.
6. Nếu mã gần giống theo ngưỡng xác nhận, app đưa tối đa 3 đề xuất vào bảng **Cần xác nhận**.
7. Nếu không tìm được mã phù hợp, dòng được đưa vào tab **Mã Mới**. App có thể gợi ý mã đã đặt hàng dựa trên độ giống tên hàng.
8. Khi người dùng bấm `Chọn`, mã được chuyển vào bảng **Giống nhau**; mã thiết kế hoặc đơn vị tính sai sẽ bị gạch ngang và ghi chú giá trị đúng.
9. Khi người dùng bấm `Bỏ qua`, đề xuất hiện tại bị loại và app tính lại kết quả.

## Các tab trong ứng dụng

- **Bảng thống kê**: tổng quan số lượng dữ liệu và kết quả so sánh.
- **Dữ Liệu Mã**: danh sách mã đã đặt hàng, có phân trang khi dữ liệu lớn.
- **Dữ Liệu Thiết Kế**: danh sách BOM/thiết kế đã load.
- **So Sánh**: gồm bảng **Cần xác nhận** ở trên và bảng **Giống nhau** ở dưới.
- **Mã Mới**: các mã thiết kế chưa có trong danh sách mã đã đặt hàng.

## Xuất Excel

Nút **Xuất Excel** tạo file `DeNghiMuaHang_YYYYMMDD_HHMM.xlsx` với sheet `De Nghi Mua Hang`.

File xuất dùng mẫu tiêu đề đề nghị mua hàng MEIKO và lấy thông tin còn thiếu từ vùng tiêu đề của file thiết kế nếu có, ví dụ `Mã dự án`, `Tên dự án`, `Số lượng máy`, `Mã tài liệu`, `Bộ phận`.

Bảng chi tiết chỉ xuất các bản vẽ thiết kế đã phân loại:

- Mã trong bảng **Giống nhau**.
- Mã trong tab **Mã Mới**.

Nếu vẫn còn mã trong bảng **Cần xác nhận**, app sẽ hiện popup: `Vẫn còn mã chưa xác nhận bạn có muốn tiếp tục không?`

- Chọn **Tiếp tục**: vẫn xuất file Excel.
- Chọn **Hủy bỏ**: hủy xuất file.

## Update

- Nút **Update** kiểm tra bản phát hành mới qua `electron-updater`; chức năng này chỉ chạy trên bản đã build exe.
- Nút **Thông tin** mở trang GitHub phát hành.

## Chạy phát triển

```bash
npm install
npm start
```

## Build Windows

```bash
npm run build:win
```

File build được tạo trong thư mục `release/`, gồm bản cài đặt NSIS và bản portable x64.
