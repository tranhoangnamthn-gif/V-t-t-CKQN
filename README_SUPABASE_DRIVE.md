# Web vật tư CKQN - bản Supabase + Google Drive

Bản này không dùng SQLite trên Render nữa.

- Supabase lưu dự án, vật tư, trạng thái Chưa đặt hàng / Đã đặt hàng.
- Google Drive tự tạo thư mục dự án và lưu file Excel gốc.
- Render/Vercel redeploy không làm mất dữ liệu.

## Biến môi trường cần nhập trên Render

SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_KEY=sb_publishable_xxxx
GOOGLE_DRIVE_ROOT_FOLDER_ID=ID_THU_MUC_VAT_TU_CKQN
GOOGLE_SERVICE_ACCOUNT_JSON={...toàn bộ nội dung file json...}

## Supabase

Vào Supabase > SQL Editor, chạy file `SUPABASE_SCHEMA.sql`.

## Google Drive

Tạo thư mục `VẬT TƯ CKQN`, share quyền Editor cho email service account.


## Cập nhật mới
- Nút xóa dự án chỉ đặt trong trang chi tiết dự án.
- Có thể upload file Excel hoặc PDF. Excel được import vật tư; PDF được lưu vào Google Drive làm hồ sơ gốc, chưa tự bóc tách vật tư.
