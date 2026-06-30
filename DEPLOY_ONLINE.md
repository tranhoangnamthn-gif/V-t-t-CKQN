# Đưa web quản lý vật tư lên online để có link truy cập

Mục tiêu: sau khi deploy xong, bạn sẽ có link dạng:

```text
https://material-control-web.onrender.com
```

Hoặc có thể gắn tên miền riêng dạng:

```text
https://vattu.tencongty.com
```

## Cách khuyên dùng: Render + GitHub

Bản này đã được cấu hình sẵn để deploy Flask bằng Gunicorn, có `Procfile`, `render.yaml`, `Dockerfile` và biến `MATERIAL_DATA_DIR` để lưu dữ liệu.

### Bước 1. Tạo GitHub repository

1. Đăng nhập GitHub.
2. Tạo repository mới, ví dụ: `material-control-web`.
3. Upload toàn bộ thư mục này lên repository đó.

### Bước 2. Deploy trên Render

1. Đăng nhập Render.
2. Chọn **New +** → **Blueprint** nếu dùng file `render.yaml`, hoặc **Web Service** nếu tạo thủ công.
3. Kết nối GitHub repository `material-control-web`.
4. Nếu tạo thủ công, nhập:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn wsgi:app --bind 0.0.0.0:$PORT
```

5. Thêm Environment Variables:

```text
MATERIAL_APP_SECRET = một_chuỗi_bí_mật_bất_kỳ
MATERIAL_DATA_DIR = /var/data
PYTHON_VERSION = 3.12.4
```

6. Thêm Persistent Disk:

```text
Mount Path: /var/data
Size: 1 GB hoặc lớn hơn
```

Nếu không thêm Persistent Disk, dữ liệu SQLite và file Excel upload có thể mất khi server redeploy/restart.

### Bước 3. Lấy link web

Sau khi deploy thành công, Render sẽ cấp link dạng:

```text
https://ten-service.onrender.com
```

Gửi link này cho mọi người để truy cập bằng Chrome/Edge trên máy tính hoặc điện thoại.

## Chạy local để kiểm tra trước khi deploy

```bash
pip install -r requirements.txt
python app.py
```

Mở:

```text
http://127.0.0.1:5000
```

## Chạy bằng Docker

```bash
docker build -t material-control-web .
docker run -p 5000:5000 -v material_data:/data material-control-web
```

Mở:

```text
http://127.0.0.1:5000
```

## Ghi chú bảo mật

Bản hiện tại ưu tiên đơn giản, mọi người có link đều có thể vào xem/sửa. Nếu đưa lên internet công khai, nên bổ sung đăng nhập và phân quyền trước khi dùng chính thức.
