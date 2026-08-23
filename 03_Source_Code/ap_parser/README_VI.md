# AP Payment List Parser (bản đang dùng: `ap_payment_template_parser_fixed_v4.py`)

Đọc toàn bộ file AP payment list trong `input_ap/`, gộp lại thành **một** file
Excel duy nhất (`AP_PAYMENT_TEMPLATE_OUTPUT.xlsx`, sheet `PAYMENT_LIST`) để
upload vào dashboard qua nút **"Import AP Summary"**.

## Cách chạy (đơn giản nhất)

1. Copy toàn bộ file AP của batch vào folder `input_ap/`.
2. Double-click `Run_AP_Parser.cmd`.
3. Lấy file `AP_PAYMENT_TEMPLATE_OUTPUT.xlsx` (nằm ngay trong folder này, không phải trong `output/`).

**Nếu chạy xong không thấy gì / báo lỗi:** nguyên nhân phổ biến nhất là file
`AP_PAYMENT_TEMPLATE_OUTPUT.xlsx` đang **mở sẵn trong Excel** — Windows sẽ
khóa file đó lại và Python không ghi đè được. Đóng file trong Excel rồi chạy
lại `Run_AP_Parser.cmd`. Bản CMD hiện tại sẽ in rõ lỗi này ra màn hình và chờ
bạn nhấn phím trước khi đóng cửa sổ, thay vì tắt ngay không kịp đọc.

## Chạy trực tiếp bằng command line (nếu cần)

```powershell
py -3 ap_payment_template_parser_fixed_v4.py --input input_ap --output AP_PAYMENT_TEMPLATE_OUTPUT.xlsx
```

2 tham số bắt buộc: `--input` (folder chứa file AP) và `--output` (đường dẫn
file Excel kết quả) — không cần `--config`, `--batch-date`, hay
`--bank-history`.

### `--period YYYY-MM` (tùy chọn) — cho file dạng "sổ dồn nhiều năm"

```powershell
py -3 ap_payment_template_parser_fixed_v4.py --input input_ap --output AP_PAYMENT_TEMPLATE_OUTPUT.xlsx --period 2026-08
```

Một số file AP (VD: `VNGSing-CITI Payment list...`, `Greennode SING - DB...`)
không phải là 1 file = 1 batch, mà là **sổ dồn chạy từ nhiều năm trước** (từng
thấy trải từ 2024 đến hiện tại). Nếu không lọc, parser sẽ đưa cả hàng ngàn
dòng đã cũ/có thể đã trả vào chung với batch hiện tại.

`--period` chỉ tác động đến **file nào có dữ liệu trải hơn 6 tháng** (tự động
phát hiện) — file batch bình thường (như ZPI, KMZ...) dù có vài invoice pending
trải 2-3 tháng vẫn được giữ nguyên, không bị lọc, dù có truyền `--period` hay
không. Double-click `Run_AP_Parser.cmd` sẽ hỏi giá trị này (Enter để bỏ qua,
không lọc gì).

## Output columns

```text
Due Date(Payment date)
Supplier Number
Invoice Number
Due Date(Value date)
Bene name
Amount
Currency
E-form No.
Bank charge
Detail
Account number   <- tài khoản ngân hàng của VENDOR, không phải của công ty
Swiftcode
Entity
Source File       <- tên file AP gốc, để truy ngược khi cần
```

Entity được đoán tự động từ tên file (`guess_entity()`, khớp đúng logic với
`guessEntity()` phía dashboard) — không dựa vào nội dung sheet. Đã hỗ trợ thêm
entity **GREENNODE SG** (file tên chứa "greennode").

⚠️ **GREENNODE SG là entity mới, chưa có ORG_ID và default Payment Bank/Method
trên dashboard** (map hiện tại chỉ có 7 entity gốc). Payment List sẽ vẫn hiện
đúng dữ liệu GREENNODE SG, nhưng ERP Export và Cash Position cho entity này sẽ
chưa đúng cho tới khi có ORG_ID thật + rule Payment Bank mặc định (GREENNODE SG
không có tài khoản Citi, chỉ có Deutsche Bank/HSBC).

## Sheet nào được đọc trong 1 file

- Chỉ đọc sheet nào có đủ header nhận diện được (Supplier Number/Amount/
  Currency/E-form No./Bene name — kể cả tên cột khác chuẩn như "Eform No."
  không gạch nối, "Supplier name"/"TÊN ĐỐI TÁC" thay cho "Bene name").
- **Bỏ qua mọi sheet có "pending" trong tên** (VD "Pending payment", "Pending
  Payment") — đây là khoản chưa duyệt/chưa đến hạn, không tính vào Payment
  List hiện tại (khớp đúng rule cũ). Sheet loại này từng thấy tới hơn 1 triệu
  dòng "ảo" (định dạng kéo dài, không phải data thật) nên bỏ qua cũng giúp
  chạy nhanh hơn nhiều.
- Dòng cần có Invoice Number **hoặc** E-form No. (không bắt buộc phải có cả
  hai) — một số file chỉ có E-form No./Transaction No. làm reference, không
  có cột Invoice Number riêng.

## Các file khác trong folder này

Các file của tool cũ (`ap_payment_list_parser.py`, `AP_Parser_Config.xlsx`,
`AP_Import_Output_Template.xlsx`, `AP_IMPORT_SCHEMA_V2.json`,
`Build_Standalone_EXE.cmd`, `BANK_Reference_History_Template.xlsx`,
`QA_REPORT.txt`, cùng các snapshot `_archive/`) **đã bị xóa khỏi repo**
trong đợt dọn dẹp — schema và luồng của chúng khác hẳn và không còn được
dùng. Cần xem lại thì lấy từ git history (commit trước đợt cleanup).
Script đang dùng thật là `ap_payment_template_parser_fixed_v4.py`, chạy qua
`Run_AP_Parser.cmd`.
