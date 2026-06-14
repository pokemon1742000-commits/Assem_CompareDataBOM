# Kế Hoạch Phần Mềm Chuẩn Hóa Mã Hàng Thiết Kế

## Mục tiêu

Xây dựng phần mềm version `1.0.0` để chuẩn hóa mã hàng thiết kế dựa trên danh sách mã đã đặt hàng.

## Nguồn dữ liệu

1. **Mã Đã Đặt Hàng**
   - Cột bắt buộc: `Mã hàng tồn kho`.
   - Cột bổ sung: `Tên hàng`, `Đơn vị tính`.
   - Hiển thị trong tab **Dữ Liệu Mã**.

2. **Dữ Liệu Thiết Kế**
   - Cột bắt buộc: mã bản vẽ.
   - Cột bổ sung: tên mặt hàng, nhà sản xuất, số lượng, đơn vị tính.

## Luồng xử lý

1. Load file **Mã Đã Đặt Hàng**.
2. Load file **Dữ Liệu Thiết Kế**.
3. So sánh từng mã thiết kế với danh sách mã đã đặt hàng.
4. Mã trùng tuyệt đối chuyển vào bảng **Giống nhau**.
5. Mã gần giống theo fuzzy matching mã + đơn vị tính chuyển vào bảng **Xác nhận**.
6. Người dùng bấm `Chọn` để nhận mã đề xuất hoặc `Bỏ qua` để chuyển mã thiết kế sang **Mã Mới**.
7. Mã không tìm thấy chuyển vào sheet **Mã Mới**.

## Kết quả xuất Excel

- **So Sanh**
  - Bảng xác nhận phía trên.
  - Bảng giống nhau phía dưới.
  - Mã thiết kế được xác nhận sẽ gạch ngang và ghi chú mã đúng.

- **Ma Moi**
  - Chỉ gồm mã bản vẽ từ file thiết kế chưa có trong danh sách mã đã đặt hàng.

## Tiêu chí hoàn thành

- Không còn luồng dữ liệu kho cũ trên giao diện.
- Nút chính là **Mã Đã Đặt Hàng**.
- Tiêu đề phần mềm là **Chuẩn Hóa Mã Hàng Thiết Kế**.
- Version package là `1.0.0`.
- Build Windows tạo được artifact trong `release/`.
