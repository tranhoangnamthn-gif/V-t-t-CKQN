# Web kiểm soát vật tư đặt hàng

Ứng dụng quản lý vật tư theo dự án:

- Tạo dự án trên web.
- Tự tạo thư mục dự án trong Google Drive sau khi kết nối OAuth.
- Upload Excel/PDF vào từng dự án.
- Excel tự import vật tư vào Supabase.
- PDF lưu hồ sơ gốc trong Google Drive.
- Cập nhật tình trạng Chưa đặt hàng / Đã đặt hàng.
- Xóa dự án trong app sẽ đưa thư mục Drive vào thùng rác.
- Có nút Đồng bộ Drive để app tự xóa dự án khi thư mục Drive đã bị xóa thủ công.

## Deploy

```bash
pip install -r requirements.txt
gunicorn wsgi:app
```

Xem thêm `README_OAUTH_DRIVE.md` để cấu hình Google Drive OAuth.
