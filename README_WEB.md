# Chạy bản web online

Project hiện có thêm bản web Flask, giữ lõi xử lý PDF sang Excel trong `app.py`.

## Chạy thử trên máy

```bat
python -m pip install -r requirements.txt
python web_app.py
```

Mở trình duyệt:

```text
http://localhost:8000
```

## Upload lên VPS Linux

Ví dụ thư mục deploy:

```bash
cd /var/www/pdf_hoc_ba_to_excel_tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Chạy thử:

```bash
gunicorn -w 2 -b 0.0.0.0:8000 web_app:web
```

Sau đó trỏ Nginx/Apache reverse proxy về `127.0.0.1:8000`.

## Biến môi trường

```bash
export HOCBA_SECRET_KEY="doi-chuoi-bi-mat-nay"
export HOCBA_WEB_STORAGE="/var/www/pdf_hoc_ba_storage"
export HOCBA_MAX_UPLOAD_MB="200"
export HOCBA_JOB_TTL_SECONDS="86400"
```

## Luồng sử dụng

1. Mở web.
2. Upload nhiều file PDF học bạ.
3. Chọn file Excel mẫu hoặc dùng `templates/hoc_ba_mau (2).xlsx`.
4. Kiểm tra danh sách, bỏ chọn file không cần chuyển.
5. Bấm `Chuyển và tải ZIP`.

Kết quả là file ZIP chứa các file `.xlsx` đã chuyển và file `LOI_*.txt` nếu file nào bị lỗi.
