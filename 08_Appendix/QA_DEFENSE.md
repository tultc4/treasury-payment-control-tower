# Q&A Defense — khi bị challenge về độ chính xác

*Tài liệu ôn tập cho người trình bày. Câu tiếng Anh in nghiêng dùng được nguyên văn trong Q&A.*

## Khung trả lời 30 giây (thuộc lòng)

> Em không yêu cầu mọi người tin tool. Em thiết kế để tool **tự chứng minh** qua 3 tầng:
> 1. **Logic tất định** — không có AI lúc chạy; cùng input luôn ra cùng output, mọi rule đọc được.
> 2. **Control đối soát tự động ở mỗi bước** — mọi con số đều có phương trình cân đối kèm theo.
> 3. **Con người quyết định mọi hành động** — tool không bao giờ tự chuyển tiền.
>
> *"I don't ask anyone to trust the tool — it proves itself: deterministic logic, a reconciliation control at every step, and a human decision before any money moves."*

## Trả lời từng challenge

### 1. "Lỡ parser đọc THIẾU dòng từ file AP thì sao?"
Mở **Admin → Parse completeness control** (hoặc sheet PARSE_CONTROL trong file output) ngay trên máy:
- Phương trình bắt buộc từng file: **Số dòng quét = Giữ + Loại theo từng lý do**.
- Ví dụ thật: ZPI `83 = 11 giữ + 68 ngoài ngày chạy lệnh + 1 thiếu số tiền + 3 thiếu reference`. File VNGSING 8.367 dòng cũng được giải trình đủ từng loại.
- File không cân → đóng dấu **CHECK** + cảnh báo đỏ ngay trang chủ.
- **Mời giám khảo tự cộng lại** — khoảnh khắc thuyết phục nhất.

*"Every candidate row must be accounted for: kept + excluded-by-reason = scanned. Nothing can vanish silently."*

### 2. "Lỡ parse SAI số tiền?"
- Tổng tiền giữ lại **theo từng currency, từng file** nằm ngay trong control sheet → đối chiếu với dòng total của file AP gốc mất 30 giây.
- Kể chuyện ngược: chính tool đã **phát hiện** lỗi hệ thống trong file nguồn (cột amount sai) mà quy trình kiểm tay trước đây sống chung với nó suốt.

### 3. "Code do AI viết thì tin sao được?"
- *"There is no AI at runtime — that's why this is a Non-Agent entry."* Toàn bộ là rule tất định, đọc được, kiểm thử được.
- 22 kịch bản test có bằng chứng (06_Testing/TEST_CASES.md); **ghi trung thực 3 case NOT TESTED** thay vì đánh PASS hết.
- AI chỉ là công cụ viết code — giống như Excel không tự đảm bảo công thức đúng; người xây + bộ test + control đảm bảo.

### 4. "Sai thì ai chịu trách nhiệm?"
- Tool không thay quyết định: export phải người duyệt · override phải nhập **lý do** · bulk update có xác nhận + undo.
- Tất cả nằm trên **audit log**: ai, lúc nào, sửa từ gì thành gì.
- *"The tool doesn't replace accountability — it gives accountability evidence."*

### 5. "Đã chạy thật chưa hay chỉ demo?"
- Batch thật hàng tuần: **46 lệnh · 7 entity · 6 currency · 3 bank**.
- Bắt được thiếu hụt **~5 triệu THB trước khi chạy lệnh** — thủ công thì phát hiện sau khi đã submit.

### 6. "Sao dám nói bank reconciliation đúng?"
- So khớp **4 trường độc lập** trên từng lệnh: amount, beneficiary, account, SWIFT — lệch trường nào nêu đích danh trường đó kèm 2 giá trị (bank X vs ours Y).
- Xác nhận debit **chỉ bằng giao dịch chi tiết trên sổ phụ** — biến động số dư không bao giờ được chấp nhận làm bằng chứng.
- Reference trùng lịch sử bank → chặn cứng, muốn qua phải override có lý do.

## Tuyệt đối KHÔNG nói
- ❌ "Chính xác 100%" / "AI đảm bảo không sai"
- ❌ Giấu các case NOT TESTED — ngược lại **chủ động khai** ("duplicate statement em chưa test, đã ghi trong Known Limitations") là điểm cộng với auditor
- ❌ Đổ cho AI khi bị hỏi khó — luôn kéo về: logic tất định + control + con người quyết định

## Câu chốt khi bị dồn

> Quy trình cũ bằng tay không có bất kỳ control nào trong số này — **sai là sai im lặng**. Tool này có thể chưa hoàn hảo, nhưng mỗi lỗi của nó đều **hiện hình và có dấu vết**. Đó chính là định nghĩa của kiểm soát.
>
> *"The manual process had none of these controls — its errors were silent. This tool may not be perfect, but every error it makes is visible and traceable. That is what control means."*
