# Tool chuyển PDF học bạ THCS sang Excel mẫu

**Tác giả:** Võ Thành Khải - 0913046881 / 0913784333

## Chức năng
- Chọn thư mục input chứa PDF học bạ
- Chọn thư mục output để lưu Excel kết quả
- Chọn file Excel mẫu `.xlsx` (mặc định: `templates/hoc_ba_mau (2).xlsx`)
- Grid danh sách file PDF, có tick chọn từng file
- Check chọn tất cả, bỏ chọn tất cả
- Xóa list file đã chọn
- Chuyển file đã chọn
- Tên file kết quả: `<Lớp>_<Tên học sinh>.xlsx`

## Chạy bằng Python
Cài Python 3.10 trở lên, sau đó mở CMD tại thư mục này:

```bat
run_app.bat
```

Hoặc chạy tay:

```bat
python -m pip install -r requirements.txt
python app.py
```

## Biên dịch thành EXE
Mở CMD tại thư mục này và chạy:

```bat
build_exe.bat
```

File exe sau khi build nằm tại:

```text
dist\HocBaPdfToExcel.exe
```

## Cấu trúc thư mục
```text
pdf_hoc_ba_to_excel_tool/
├─ app.py
├─ requirements.txt
├─ run_app.bat
├─ build_exe.bat
├─ templates/
│  └─ hoc_ba_mau (2).xlsx
├─ input_sample/
│  └─ so_hoc_ba_1a1.01.quan_hong_thuy_an.2520695407.pdf
└─ output/
```

## Lưu ý
- PDF cần là PDF có text, không phải ảnh scan hoàn toàn. Nếu là ảnh scan, cần bổ sung OCR.
- File Excel mẫu hiện lấy theo mẫu THCS `hoc_ba_mau (2).xlsx`. Nếu mẫu thay đổi vị trí dòng/cột quá nhiều, cần chỉnh mapping trong `app.py`.
- Các thư viện dùng: PyMuPDF để đọc PDF, openpyxl để ghi Excel, Tkinter làm giao diện.
"# pdf_hoc_ba_to_excel_tool_C23" 
