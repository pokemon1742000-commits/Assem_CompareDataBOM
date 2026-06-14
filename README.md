# Inventory Compare App

Phần mềm Electron dùng để load **Dữ Liệu Kho** và **Dữ Liệu Thiết Kế** từ Excel, tự động so sánh mã bản vẽ và số lượng/máy, sau đó xuất Excel theo bảng đang xem.

## Tính năng chính

- Chọn một hoặc nhiều file **Dữ Liệu Kho** trong cùng một lần load.
- Chọn một hoặc nhiều file **Dữ Liệu Thiết Kế** trong cùng một lần load.
- Nếu file Excel có nhiều sheet, app cho chọn một hoặc nhiều sheet cần load trước khi so sánh.
- Có thể load thêm nhiều lần; dữ liệu được cộng dồn.
- Tách dữ liệu kho từ cột A theo dấu `,`.
- Tự chuẩn hóa các dòng kho bị đảo vị trí `nhà sản xuất, số lượng` thành `số lượng, nhà sản xuất`.
- Đọc file thiết kế theo các cột: tên mặt hàng, mã bản vẽ, nhà sản xuất, số lượng/máy.
- Cộng dồn số lượng khi một mã bản vẽ xuất hiện nhiều dòng.
- Tự động so sánh sau khi load dữ liệu.
- Fuzzy matching để đề xuất mã bản vẽ gần giống.
- Gộp phần xác nhận mã vào đầu bảng `So Sánh`.
- Xuất riêng Excel cho bảng `So Sánh` hoặc `Thiếu Thừa`.
- Lưu đường dẫn file đã load gần nhất vào file temp để tự khôi phục khi mở lại app.
- Có nút `Update` để kiểm tra và tải phiên bản mới từ GitHub Releases.

## Cài đặt để phát triển

Yêu cầu:

- Node.js
- npm

Cài dependencies:

```bash
npm install
```

Chạy app ở môi trường dev:

```bash
npm start
```

## Build file EXE

Build bản Windows:

```bash
npm run build:win
```

File sau khi build nằm trong thư mục `release`:

- `Inventory-Compare-Setup-3.0.0-x64.exe`: bản cài đặt.
- `Inventory-Compare-Portable-3.0.0-x64.exe`: bản portable.
- `win-unpacked/Inventory Compare.exe`: bản unpacked để chạy thử nhanh.

## Hướng dẫn sử dụng

### 1. Load Dữ Liệu Kho

1. Bấm `Dữ Liệu Kho`.
2. Chọn một hoặc nhiều file Excel kho.
3. Nếu file có nhiều sheet, chọn một hoặc nhiều sheet cần dùng trong hộp `Chọn sheet để load`.
4. Sau khi load file đầu tiên, nút đổi thành `Thêm Dữ Liệu Kho`.
5. Nếu chọn nhiều file, app sẽ đọc tất cả sheet đã chọn và cộng dồn vào bảng kho.

Format dữ liệu kho:

```text
AUTS260311, S260311-FR7-001, 2, PMA, NK-PMA-260520, 20/05/2026, MKAC-PMA-260415-2
```

App lấy các trường:

- Tên mã dự án
- Mã bản vẽ
- Số lượng/máy
- Nhà sản xuất
- Ngày nhập kho
- N/A

Nếu dữ liệu bị đảo như:

```text
MEC2510010, 2102003-CV2-21, FBT, 1, NK-FBT-260512-25, 14/05/2026, MKAC-FBT-260421
```

App sẽ tự hiểu `Số lượng/máy = 1` và `Nhà sản xuất = FBT`.

### 2. Load Dữ Liệu Thiết Kế

1. Bấm `Dữ Liệu Thiết Kế`.
2. Chọn một hoặc nhiều file Excel BOM/thiết kế.
3. Nếu file có nhiều sheet, chọn một hoặc nhiều sheet cần dùng trong hộp `Chọn sheet để load`.
4. Dữ liệu được cộng dồn vào bảng thiết kế.

Các cột được đọc:

- Tên mặt hàng
- Model / Mã bản vẽ
- Nhà sản xuất
- Số lượng/máy

### 3. Tự động so sánh

App tự động so sánh sau khi load dữ liệu.

- Nếu chỉ có dữ liệu kho, tất cả mã nằm trong bảng `Thiếu Thừa` với trạng thái `Thừa`.
- Nếu chỉ có dữ liệu thiết kế, tất cả mã nằm trong bảng `Thiếu Thừa` với trạng thái `Thiếu`.
- Nếu có cả hai nguồn, app so sánh mã bản vẽ và số lượng/máy.
- Mã đủ hàng nằm trong bảng `So Sánh`.
- Mã thiếu, thừa hoặc lệch số lượng nằm trong bảng `Thiếu Thừa`.
- Mã gần giống nhưng chưa chắc chắn nằm ở đầu bảng `So Sánh` để xác nhận.

Ngưỡng mặc định:

- Tự ghép: 95%
- Cần xác nhận: 70%

### 4. Xác nhận mã gần giống

Trong bảng `So Sánh`:

- Cột `Mã BOM đề xuất` có dropdown để chọn mã phù hợp.
- Bấm `Chọn` để ghép mã kho với mã BOM đang chọn.
- Bấm `Bỏ qua` để bỏ qua các đề xuất.
- Sau khi chọn hoặc bỏ qua, app tính lại kết quả nhưng không tự chuyển tab.

### 5. Xuất Excel

Nút `Xuất Excel` hoạt động theo bảng đang xem:

- Đang xem `So Sánh`: xuất file Excel bảng so sánh.
- Đang xem `Thiếu Thừa`: xuất file Excel bảng thiếu thừa.

### 6. Clear dữ liệu

Bấm `Clear dữ liệu` để xóa:

- Dữ liệu kho đã load
- Dữ liệu thiết kế đã load
- Kết quả so sánh
- Bảng thiếu/thừa
- Các lựa chọn xác nhận
- Đường dẫn file đã lưu trong temp

## Update phiên bản mới

Nút `Update` kiểm tra bản mới trên GitHub Releases của repo:

```text
https://github.com/pokemon1742000-commits/Assem_CompareDataBOM
```

Để app tự update hoạt động đúng:

1. Build bản mới với version cao hơn trong `package.json`.
2. Publish file build lên GitHub Releases.
3. Release cần có file `.exe`, `.blockmap` và `latest.yml` do `electron-builder` tạo trong thư mục `release`.
4. Người dùng bấm `Update`; nếu có bản mới, app sẽ tải về và hỏi khởi động lại để cài.

Lưu ý: nút `Update` chỉ hoạt động trên bản đã build `.exe`, không hoạt động khi chạy bằng `npm start`.

## Trial và License

- Khi chưa nhập license, phần mềm hiển thị `Trial` và cho dùng thử 7 ngày.
- Bộ đếm thời gian trial nằm dưới nút `Update` và `Thông tin`.
- Khi hết hạn trial, phần mềm sẽ khóa thao tác và chỉ cho nhập license hoặc tắt phần mềm.
- Khi nhập đúng license, phần mềm chuyển sang bản vĩnh viễn và hiển thị version hiện tại, ví dụ `v3.0.0`.
- License hợp lệ có dạng `XXX-XXX-XXX-XXXX`.

## Cấu trúc file chính

- `main.js`: Electron main process, đọc/xuất Excel, lưu temp, auto-update.
- `preload.js`: cầu nối API an toàn giữa renderer và Electron.
- `renderer.js`: xử lý giao diện, parse dữ liệu, so sánh, fuzzy matching.
- `index.html`: cấu trúc giao diện.
- `style.css`: giao diện và theme.
- `KE_HOACH_PHAN_MEM_SO_SANH_KHO_BOM.md`: kế hoạch và ghi chú triển khai.
