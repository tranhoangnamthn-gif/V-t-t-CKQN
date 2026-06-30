# Deploy online trên Render

## 1. Chạy SQL trong Supabase

Mở `SUPABASE_SCHEMA.sql`, copy toàn bộ, dán vào Supabase > SQL Editor > Run.

## 2. Upload code lên GitHub

Upload toàn bộ nội dung thư mục này lên repository GitHub.

## 3. Cấu hình Environment Variables trên Render

```text
SUPABASE_URL=https://xxxxx.supabase.co
SUPABASE_KEY=sb_publishable_xxxxx
GOOGLE_DRIVE_ROOT_FOLDER_ID=ID_THU_MUC_GOC_TREN_DRIVE
GOOGLE_OAUTH_CLIENT_ID=xxxxx.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=xxxxx
GOOGLE_OAUTH_REDIRECT_URI=https://vat-tu-ckqn.onrender.com/google/callback
MATERIAL_APP_SECRET=chuoi_bat_ky_kho_doan
```

## 4. Deploy

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
gunicorn wsgi:app
```

## 5. Kết nối Google Drive

Sau khi web Live, mở web và bấm **Kết nối Google Drive**.
