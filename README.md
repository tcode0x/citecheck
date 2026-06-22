# citecheck — Trình kiểm tra trích dẫn cho bài báo khoa học

> **Repo:** https://github.com/tcode0x/citecheck

Kiểm tra trích dẫn của bài báo từ **`.docx`**, **`.pdf`**, **`.txt`/`.md`**, hoặc **`.tex` + `.bib`**.
Tự phát hiện: tài liệu **trùng lặp** (cùng DOI/tiêu đề nhưng khác số), trích dẫn **không có mục**,
mục **không được trích**, **marker gãy** (`[?]`, `??`, `\ref{}`), **số bị thiếu** trong dải đánh số,
và **sai khớp tên tác giả** ("Author et al. [n]" không khớp danh mục). Có thể **đối chiếu online**
qua Crossref + OpenAlex + arXiv để xác minh tiêu đề/năm/DOI. Xuất báo cáo `.json`, `.md`, `.csv`
để rà tay.

---

## 1. Yêu cầu

- Python **3.8+**
- `git` (nếu clone)
- Kết nối mạng (chỉ khi dùng `--api`)

Kiểm tra Python:
```bash
python --version        # hoặc: python3 --version
```

---

## 2. Cài đặt

### Cách A — Clone từ git

```bash
git clone https://github.com/tcode0x/citecheck.git
cd citecheck

# Tạo môi trường ảo (khuyến nghị)
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# Cài tất cả phụ thuộc (gồm cả PDF)
pip install -r requirements.txt
```

### Cách B — Chưa có repo: tạo repo của riêng bạn từ các file này

```bash
mkdir citecheck && cd citecheck
# chép citecheck.py và requirements.txt vào thư mục này, rồi:
git init
git add citecheck.py requirements.txt README.md .gitignore
git commit -m "init citecheck"
# (tùy chọn) đẩy lên GitHub:
git branch -M main
git remote add origin https://github.com/tcode0x/citecheck.git
git push -u origin main

python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

### Cách C — Không dùng git, cài nhanh phụ thuộc

```bash
# Bắt buộc:
pip install requests python-docx
# Để đọc PDF (chọn 1 trong 2; pymupdf khuyến nghị):
pip install pymupdf
# hoặc:  pip install pdfplumber
```

> Lưu ý PDF: tool tự ưu tiên **PyMuPDF (`fitz`)**, nếu không có sẽ thử **pdfplumber**.
> Nếu cả hai đều thiếu, dùng `--input *.pdf` sẽ báo lỗi nhắc cài.

---

## 3. Cách dùng

### 3.1. Kiểm tra file Word (.docx)
```bash
python citecheck.py --input paper.docx --out report
```

### 3.2. Kiểm tra file PDF  ← (loại PDF)
```bash
python citecheck.py --input paper.pdf --out report
```

### 3.3. Kiểm tra LaTeX (.tex + .bib)
```bash
python citecheck.py --tex main.tex intro.tex --bib refs.bib --out report
```

### 3.4. Nhiều file một lúc / file .txt
```bash
python citecheck.py --input body.txt references.txt --out report
```

### 3.5. Bật đối chiếu online (Crossref + OpenAlex)
```bash
python citecheck.py --input paper.pdf --api --mailto ban@email.com --out report
```

### 3.6. Thêm arXiv và giới hạn số mục tra (chạy thử nhanh)
```bash
python citecheck.py --input paper.docx --api \
    --providers crossref openalex arxiv \
    --mailto ban@email.com --limit 15 --out report
```

---

## 4. Các tham số

| Cờ | Ý nghĩa |
|---|---|
| `--input, -i` | File `.docx` / `.pdf` / `.txt` / `.md` (nhận nhiều file) |
| `--tex` + `--bib` | Chế độ LaTeX (thay cho `--input`) |
| `--api` | Bật đối chiếu online |
| `--providers` | Nguồn tra: `crossref openalex arxiv` (mặc định: `crossref openalex`) |
| `--mailto` | Email cho "polite pool" của Crossref/OpenAlex (khuyến nghị) |
| `--limit N` | Chỉ tra online N mục đầu (chạy thử) |
| `--no-cache` | Không dùng cache `.citecache.json` |
| `--dup-threshold` | Ngưỡng coi 2 tiêu đề là trùng (mặc định `0.85`) |
| `--out TÊN` | Tên file xuất (không đuôi). Sinh `TÊN.json/.md/.csv` |

---

## 5. Kết quả xuất ra

- **`report.md`** — báo cáo người đọc: tóm tắt + danh sách trùng lặp, undefined, unused,
  marker gãy, sai tên tác giả, và **bảng đầy đủ từng tham chiếu** (tiêu đề parse, năm, DOI,
  trạng thái API, tiêu đề tra được, link) để **rà tay**.
- **`report.csv`** — mở Excel/Sheets, mỗi dòng một tham chiếu kèm kết quả API.
- **`report.json`** — dữ liệu máy đọc, tích hợp pipeline khác.

Trạng thái đối chiếu API mỗi mục: `OK`, `TITLE_MISMATCH`, `YEAR_MISMATCH`,
`DOI_MISMATCH`, `LOW_CONFIDENCE`, `NOT_FOUND`.

---

## 6. Ví dụ thực tế

Chạy trên một bài DR grading (PDF, 42 tài liệu), không cần `--api`:

```
=== TÓM TẮT ===
Mục tham khảo            : 42
Nhóm trùng lặp           : 5      ← [11]=[26]=[35], [22]=[29], [23]=[38], [24]=[34], [25]=[42]
Trích dẫn không có mục    : 0
Mục không được trích      : 4
Marker gãy               : 2      ← [?] và ??
Sai khớp tên tác giả      : 2      ← ravi[36], hossein[31]
Mục thiếu trường          : 34
```

Xem `example_report.md` để thấy định dạng báo cáo đầy đủ.

---

## 7. Xử lý sự cố

- **`Cần cài python-docx`** → `pip install python-docx`
- **`Để đọc PDF cần pip install pymupdf`** → cài `pymupdf` (hoặc `pdfplumber`)
- **`Cần cài requests`** → `pip install requests`
- **API bị 429 / chậm** → tool tự retry + đợi `Retry-After`; thêm `--mailto` để vào polite pool;
  cache giúp lần chạy sau nhanh hơn (xóa `.citecache.json` để tra lại từ đầu).
- **Danh mục không được nhận** → tool dò các tiêu đề "References / Bibliography /
  Tài liệu tham khảo / Trích dẫn". Nếu bài dùng tiêu đề khác, hãy đảm bảo mỗi mục bắt đầu bằng
  `[n]` hoặc `n.`.
- **PDF hai cột / scan** → văn bản trích từ PDF có thể lộn xộn; ưu tiên dùng bản `.docx`
  hoặc bản `.tex+.bib` nếu có, kết quả parse sẽ chính xác hơn.

---

## 8. Quy trình rà tay đề xuất

1. Mở `report.md`, xử lý trước **mục trùng lặp** và **marker gãy** (lỗi nặng, dễ desk-reject).
2. Soát **sai khớp tên tác giả** — mở ngữ cảnh để xác nhận đúng/sai.
3. Bật `--api`, lọc các dòng `TITLE_MISMATCH / YEAR_MISMATCH / NOT_FOUND` trong `report.csv`,
   bấm link để kiểm tra từng cái.
4. Sửa `.bib`/danh mục, chạy lại tới khi sạch cảnh báo.

---

## 9. Tác giả & giấy phép

- **Tác giả:** Hieu Nd
- **Repo:** https://github.com/tcode0x/citecheck
- Đóng góp / báo lỗi: mở issue hoặc pull request trên repo.
