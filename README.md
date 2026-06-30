# Web kiểm soát vật tư đặt hàng

Ứng dụng nội bộ để quản lý vật tư theo từng dự án.

## Chức năng chính

- Tạo danh mục dự án.
- Vào từng dự án để upload file Excel đặt hàng.
- File Excel được lưu vào thư mục riêng của dự án: `data/uploads/<project_id>/`.
- Tự động đọc Excel và hiển thị vật tư trong bảng của đúng dự án.
- Cập nhật tình trạng từng vật tư ngay trên web.
- Lọc theo tình trạng hoặc tìm kiếm theo tên vật tư, mã vật tư, PO, nhà cung cấp.
- Xuất dữ liệu theo dõi ra CSV để tổng hợp báo cáo.
- Chạy trong mạng nội bộ để mọi người cùng truy cập.

## Cấu trúc dữ liệu Excel khuyến nghị

File Excel nên có các cột sau. Tên cột có thể viết khác một chút, hệ thống sẽ tự dò theo tiếng Việt hoặc tiếng Anh:

| Mã vật tư | Tên vật tư | Quy cách | ĐVT | Số lượng | Nhà cung cấp | PO | Ngày yêu cầu | Ngày cần | Ghi chú |
|---|---|---|---|---:|---|---|---|---|---|

Các tên cột hệ thống có thể nhận diện:

- Mã vật tư: `Mã vật tư`, `Mã hàng`, `Code`, `Item Code`, `Material Code`
- Tên vật tư: `Tên vật tư`, `Tên hàng`, `Nội dung`, `Description`, `Material`
- Quy cách: `Quy cách`, `Thông số`, `Spec`, `Specification`, `Model`, `Vật liệu`
- Đơn vị: `Đơn vị`, `ĐVT`, `Unit`, `UOM`
- Số lượng: `Số lượng`, `SL`, `Qty`, `Quantity`, `Khối lượng`
- Nhà cung cấp: `Nhà cung cấp`, `Supplier`, `Vendor`
- PO: `PO`, `Số PO`, `Purchase Order`, `Đơn hàng`
- Ngày cần: `Ngày cần`, `Delivery Date`, `ETA`, `Deadline`
- Ghi chú: `Ghi chú`, `Note`, `Remark`

## Cài đặt trên máy chủ nội bộ hoặc máy văn phòng

### 1. Cài Python

Cài Python 3.10 trở lên.

### 2. Giải nén thư mục

Giải nén file `material_control_web.zip` vào máy dùng làm máy chủ nội bộ.

### 3. Cài thư viện

Mở Command Prompt/PowerShell tại thư mục này và chạy:

```bash
pip install -r requirements.txt
```

### 4. Chạy web

```bash
python app.py
```

Sau khi chạy, mở trên chính máy đó:

```text
http://127.0.0.1:5000
```

Để máy khác trong cùng mạng LAN truy cập, lấy IP của máy chủ, ví dụ `192.168.1.20`, rồi mở:

```text
http://192.168.1.20:5000
```

Lưu ý: cần mở firewall cho cổng `5000` nếu máy khác chưa truy cập được.

## Cách sử dụng

1. Vào trang chủ.
2. Tạo dự án mới.
3. Bấm vào dự án.
4. Upload file Excel đặt hàng trong đúng trang của dự án đó.
5. Kiểm tra bảng vật tư đã tự hiển thị.
6. Người phụ trách cập nhật tình trạng hàng hóa tại cột cuối.
7. Khi cần tổng hợp, bấm `Xuất CSV`.

## Tình trạng hàng hóa mặc định

- Chưa đặt hàng
- Đã gửi yêu cầu báo giá
- Đã đặt hàng
- Đang sản xuất/chuẩn bị
- Đang vận chuyển
- Đã về kho
- Đã bàn giao công trình
- Thiếu / Chờ bổ sung
- Hủy / Không mua

## Ghi chú triển khai thực tế

Bản này phù hợp để chạy nội bộ trong công ty hoặc mạng LAN công trường. Nếu muốn đưa lên internet công khai, nên bổ sung đăng nhập, phân quyền người dùng, sao lưu tự động và HTTPS.

---

## Bản deploy online có link truy cập

Bản này đã bổ sung cấu hình để đưa lên hosting/cloud và tạo link truy cập trên trình duyệt:

- `Procfile`: chạy app bằng Gunicorn.
- `wsgi.py`: entrypoint production.
- `Dockerfile`: chạy bằng Docker/VPS.
- `render.yaml`: cấu hình nhanh cho Render.
- `DEPLOY_ONLINE.md`: hướng dẫn đưa lên online để có link.

Xem chi tiết trong file `DEPLOY_ONLINE.md`.


## Cập nhật mới
- Nút xóa dự án chỉ đặt trong trang chi tiết dự án.
- Có thể upload file Excel hoặc PDF. Excel được import vật tư; PDF được lưu vào Google Drive làm hồ sơ gốc, chưa tự bóc tách vật tư.

## Cập nhật đọc Excel nhiều form

Bản này tự dò dòng tiêu đề trong 50 dòng đầu và tự nhận các cột gần đúng như: Mã vật tư, Tên vật tư/Tên hàng/Description, Quy cách/Spec/Model, ĐVT/Unit/UOM, SL/Qty/Quantity, NCC/Supplier/Vendor, PO, Ghi chú/Note/Remark. Không có bước chọn lại cột thủ công. Nếu không nhận ra cột tên vật tư/mô tả thì file sẽ không import để tránh sai dữ liệu.
