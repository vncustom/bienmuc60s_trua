# App biên mục 60s trưa

Ứng dụng tạo file Map tổng cho bản tin **60s trưa** từ thư mục input gồm playlist Excel và các file kịch bản RTF.

App hiện là file độc lập: `app_bien_muc_60s_trua.py`.

## Cài đặt

```powershell
pip install -r requirements.txt
```

## Chạy app

```powershell
python app_bien_muc_60s_trua.py
```

Trên giao diện:

- Chọn thư mục input.
- Thư mục input cần chứa file playlist `BT60STRUA_*.xlsx`, file `ENDING.rtf`, và các file kịch bản `.rtf`.
- Chọn thư mục output, hoặc để trống để app tự tạo thư mục `output` trong input.
- Nhập `Mã bản tin $a090`. Ô này chỉ gợi ý mờ `K303324`, không tự dùng làm giá trị.
- Kiểm tra hoặc sửa `Người biên mục $a911`. Mặc định là `Nguyễn Thị Quyên`.
- Bấm `Bắt đầu biên mục`.

## Input

Playlist Excel:

- Tên file bắt đầu bằng `BT60STRUA_` và có đuôi `.xlsx`.
- App dùng sheet đầu tiên.
- Cột A: tên file/tên mục trong playlist.
- Cột C: ID video phải bắt đầu bằng số. Nếu có chữ cái đứng sau số, ví dụ `260611056a`, app vẫn bắt dòng.
- Cột D: trạng thái.
- Cột F: thời lượng.

File RTF:

- Tên file RTF của từng tin nên khớp với cột A trong playlist để app tìm đúng kịch bản.
- File `ENDING.rtf` được dùng để tạo cột `$a500`.
- Cột `$a500` lấy toàn bộ các dòng không rỗng trong `ENDING.rtf`, giữ nguyên thứ tự xuất hiện.

## Output

App tạo file:

```text
Map_BT60STRUA_{YYYY}_ Thang{MMDD}.xlsx
```

Ví dụ:

```text
Map_BT60STRUA_2026_ Thang0615.xlsx
```

Các cột trong file Map:

- `$a090`: mã bản tin người dùng nhập.
- `$a500`: toàn bộ nội dung không rỗng từ `ENDING.rtf`.
- `$a505`: danh sách tin theo thứ tự playlist.
- `$a911`: người biên mục, chỉ ghi ở dòng dữ liệu đầu tiên.

## Log

App ghi log vào `app_bien_muc_60s_trua.log` trong cùng thư mục chương trình. Khi mở app, log của các ngày cũ được xóa để file log không phình to theo thời gian.

## Quy tắc chính

Quy tắc bắt dòng tin:

- Cột A được trim và không phân biệt chữ hoa/thường.
- Cột A phải bắt đầu bằng một trong các prefix: `60s`, `gat60s `, `60st`, hoặc `live -`.
- Không bắt nếu cột A bắt đầu bằng `60s W `.
- Bỏ qua nếu cột A chứa `coming up`, `nhung nguoi thuc hien`, hoặc ` end`.
- Cột C phải bắt đầu bằng số. Nếu cột C trống hoặc bắt đầu bằng chữ, ví dụ `qc123`, dòng đó không được bắt. Nếu chữ cái đứng sau số, ví dụ `260611056a`, dòng đó vẫn được bắt.

Tiêu đề tin được tách từ phần đầu file RTF theo thứ tự xuất hiện:

- Lấy dòng đầu tiên IN HOA toàn bộ, dài hơn 16 ký tự và thuộc phần đầu kịch bản.
- Không yêu cầu dòng tiêu đề phải BOLD hoặc có màu GREEN.
- Không lấy các dòng cue bắt đầu bằng `C1`, `C2`, `C3`, `C4` làm tiêu đề.
- Không lấy các dòng cue hình như `CẬN GIỮA`, `TOÀN GIỮA`, `CẬN PHẢI` làm tiêu đề.
- Nếu dòng ngay phía dưới cũng IN HOA toàn bộ và dài hơn 16 ký tự thì nối thêm vào tiêu đề.
- Tiêu đề chỉ gồm 1 dòng hoặc nối tối đa 2 dòng, không nối nhiều hơn.
