# Cấu hình Google Drive OAuth

Bản này **không dùng Service Account** nữa. App sẽ kết nối Google Drive bằng OAuth, nên file upload sẽ tính vào dung lượng Google Drive của tài khoản đã cấp quyền.

## Biến môi trường cần có trên Render

```text
SUPABASE_URL=https://xxxxx.supabase.co
SUPABASE_KEY=sb_publishable_xxxxx
GOOGLE_DRIVE_ROOT_FOLDER_ID=ID_THU_MUC_VAT_TU_CKQN
GOOGLE_OAUTH_CLIENT_ID=xxxxx.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=xxxxx
GOOGLE_OAUTH_REDIRECT_URI=https://vat-tu-ckqn.onrender.com/google/callback
MATERIAL_APP_SECRET=chuoi_bat_ky_kho_doan
```

Không cần dùng biến `GOOGLE_SERVICE_ACCOUNT_JSON` nữa.

## Sau khi deploy

1. Mở trang web.
2. Bấm **Kết nối Google Drive**.
3. Đăng nhập đúng tài khoản Google có thư mục `Vật tư CKQN`.
4. Cấp quyền cho app.
5. Sau khi quay lại web, tạo dự án/upload file như bình thường.

## Lưu ý

- Tài khoản Google cấp quyền phải có quyền Editor với thư mục gốc `Vật tư CKQN`.
- Nếu app Google Cloud đang ở chế độ Testing, cần thêm email của bạn vào **Test users**.
- Không upload Client Secret hoặc token lên GitHub.
