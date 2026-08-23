# Bank Statement Transaction Parser v1

Công cụ này đọc các file sổ phụ gốc trong một folder và tạo output **transaction-level** để đối soát payment đã bị trừ tiền hay chưa trong Treasury Payment Control Tower.

## Output chính

Sau khi chạy, folder `output` có:

- `bank_transaction_detail.xlsx`: file dùng để import vào dashboard.
- `bank_transaction_detail.csv`: cùng dữ liệu ở dạng CSV.
- `summary_rebuilt.xlsx`: dữ liệu số dư được giữ lại nếu input có `summary.xlsx` hoặc balance summary.

Trong `bank_transaction_detail.xlsx`:

- `TRANSACTION_DETAIL`: mỗi dòng là một giao dịch ngân hàng.
- `QUALITY_CHECK`: số dòng, tổng debit/credit, dòng cần review.
- `PARSER_LOG`: kết quả đọc từng file.
- `BALANCE_SUMMARY`: opening/closing balance.
- `BALANCE_RECON`: kiểm tra Opening + Credit - Debit = Closing.
- `UNMAPPED_ROWS`: dòng hoặc trang PDF parser chưa hiểu được.
- `DUPLICATES`: giao dịch trùng đã loại khỏi output chính.
- `CONFIG_USED`: cấu hình đã sử dụng.

## Cách chạy trên Windows

### Lần đầu

1. Cài Python 3.11 hoặc mới hơn.
2. Double-click `Install_Requirements.bat`.

### Mỗi ngày

1. Đặt các file sổ phụ gốc vào một folder. Có thể có subfolder theo bank/date.
2. Kéo folder đó thả lên `Run_Parse_Bank_Statements.bat`.
3. Mở `output\bank_transaction_detail.xlsx`.
4. Kiểm tra `QUALITY_CHECK`, `PARSER_LOG` và các dòng `RECORD_STATUS = REVIEW`.
5. Đóng dashboard, sau đó kéo `bank_transaction_detail.xlsx` vào `Run_4_Import_Bank_Statement.bat` của Treasury Control Tower v6.

Có thể chạy bằng command line:

```bat
python bank_statement_transaction_parser.py ^
  --input "C:\Bank Statements\2026-07-30" ^
  --output-dir "C:\Bank Statements\Parsed" ^
  --config "parser_config.json" ^
  --verbose
```

## Quy tắc dữ liệu

- Một dòng output phải là một transaction, không aggregate theo ngày.
- Debit luôn được chuẩn hóa thành số âm.
- Credit luôn được chuẩn hóa thành số dương.
- Account number được giữ dạng text và được sửa leading zero theo `account_master.csv` khi có mapping.
- Raw reference và description được giữ nguyên.
- `EFORM_CANDIDATE` được trích từ customer reference, bank reference, payment details và description.
- Các dòng opening balance, closing balance, subtotal và total không được đưa vào `TRANSACTION_DETAIL`.

## File `summary.xlsx` hiện tại

Nếu input chỉ có `summary.xlsx`, parser vẫn tạo output và giữ 59 balance rows, nhưng `TRANSACTION_DETAIL` sẽ rỗng. Đây là đúng logic vì balance summary không chứa từng giao dịch.

Để có trạng thái `DEBIT_CONFIRMED`, input cần là file bank export hoặc statement có các dòng giao dịch thực tế.

## Tùy chỉnh bank format

### Thêm tên cột mới

Mở `parser_config.json` và thêm alias vào trường tương ứng. Ví dụ:

```json
{
  "header_aliases": {
    "CUSTOMER_REFERENCE": [
      "Customer Reference",
      "Your Ref",
      "Client Payment ID"
    ]
  }
}
```

### Thêm account mới

Thêm một dòng vào `account_master.csv`:

```text
BANK_GROUP,RAW_ACCOUNT_ALIASES,NORMALIZED_ACCOUNT,CURRENCY,ENTITY,...
CITI,123456789|0123456789,0123456789,USD,ENTITY NAME,...
```

### Ngày bị đọc ngược tháng/ngày

Trong `parser_config.json`, đặt bank tương ứng thành `false` nếu source dùng định dạng tháng/ngày:

```json
"date_dayfirst_by_bank": {
  "CITI": false
}
```

## PDF

Parser dùng `pdfplumber` để đọc PDF text/table. Nếu PDF là bản scan hoặc bảng quá đặc thù, sheet `PARSER_LOG` sẽ có trạng thái `REVIEW` và `UNMAPPED_ROWS` giữ nội dung trang. Khi đó cần một raw PDF mẫu để bổ sung bank-specific adapter. Công cụ không dùng OCR mặc định để tránh đọc sai số tiền và account.

## Kiểm soát trước khi import dashboard

Chỉ import khi:

- `TRANSACTION_DETAIL` có dữ liệu.
- Debit amount là số âm.
- `ACCOUNT_NUMBER`, `ACCOUNT_CURRENCY`, `ENTRY_DATE`, `TRANSACTION_AMOUNT` có đủ.
- Dòng `RECORD_STATUS = REVIEW` đã được xử lý hoặc chấp nhận.
- `PARSER_LOG` không có file quan trọng bị `NO_DATA` hoặc `ERROR`.

## Hạn chế đã biết

Code là parser framework dùng header alias và metadata detection. Do hiện tại mới có `summary.xlsx`, chưa có file sổ phụ gốc của từng bank, phần Excel/CSV đã được kiểm thử bằng sample; PDF được hỗ trợ theo table text nhưng có thể cần adapter riêng cho Deutsche Bank hoặc PDF scan.
