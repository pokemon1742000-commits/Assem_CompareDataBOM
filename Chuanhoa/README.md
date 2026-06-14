# Chuẩn Hóa Mã Hàng Thiết Kế

Phần mềm Electron version `1.0.0` dùng để so sánh mã bản vẽ trong file **Dữ Liệu Thiết Kế** với danh sách **Mã Đã Đặt Hàng**.

## Chức năng chính

- Load file **Mã Đã Đặt Hàng** có các cột `Mã hàng tồn kho`, `Tên hàng`, `Đơn vị tính`.
- Hiển thị dữ liệu đã đặt hàng trong sheet/tab **Dữ Liệu Mã** gồm `STT`, `Mã bản vẽ`, `Tên hàng`, `Đơn vị tính`.
- Load file **Dữ Liệu Thiết Kế** và tự nhận diện cột mã bản vẽ, tên mặt hàng, số lượng, nhà sản xuất, đơn vị tính.
- So sánh mã thiết kế với mã đã đặt hàng.
- Mã giống nhau được đưa vào bảng **Giống nhau** trong sheet **So Sánh**.
- Mã chưa có trong danh sách đặt hàng được đưa vào sheet **Mã Mới**.
- Mã gần giống được đưa vào bảng xác nhận ở đầu sheet **So Sánh** bằng fuzzy matching theo mã và đơn vị tính.
- Khi bấm `Chọn`, mã được đưa xuống bảng giống nhau, mã thiết kế bị gạch ngang và có ghi chú mã đúng.
- Khi bấm `Bỏ qua`, mã thiết kế được chuyển sang sheet **Mã Mới**.

## Định dạng file Mã Đã Đặt Hàng

File Excel cần có hàng tiêu đề chứa:

| Mã hàng tồn kho | Tên hàng | Đơn vị tính |
| --- | --- | --- |
| PM113077-A | BRACKET EMG | PCS |

Sau khi load, phần mềm hiển thị dữ liệu này tại tab **Dữ Liệu Mã**.

## Xuất Excel

- Đang xem **So Sánh**: xuất file có sheet `So Sanh`, gồm bảng xác nhận phía trên và bảng giống nhau phía dưới.
- Đang xem **Mã Mới**: xuất file có sheet `Ma Moi`, chỉ hiển thị mã bản vẽ từ file thiết kế.

## Chạy phát triển

```bash
npm install
npm start
```

## Build

```bash
npm run build:win
```

File build được tạo trong thư mục `release/`.
