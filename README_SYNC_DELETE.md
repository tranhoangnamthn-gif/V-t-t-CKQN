# Bản cập nhật xóa / đồng bộ Google Drive

Bản này bổ sung:

- Nút **Xóa dự án**: xóa dữ liệu dự án trong Supabase và đưa thư mục Google Drive vào thùng rác.
- Nút **Đồng bộ Drive**: nếu bạn xóa thư mục dự án trực tiếp trên Google Drive, bấm nút này để app xóa dự án tương ứng khỏi Supabase.
- Khi tạo dự án, tên thư mục Drive sẽ là `Mã dự án - Tên dự án` nếu có mã dự án.
- Chặn tạo trùng mã dự án để tránh tạo nhiều thư mục giống nhau.

Cập nhật bằng cách upload đè toàn bộ file/thư mục trong gói này lên GitHub, sau đó Render sẽ redeploy.


## Cập nhật mới
- Nút xóa dự án chỉ đặt trong trang chi tiết dự án.
- Có thể upload file Excel hoặc PDF. Excel được import vật tư; PDF được lưu vào Google Drive làm hồ sơ gốc, chưa tự bóc tách vật tư.
