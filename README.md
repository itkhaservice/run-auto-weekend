# Auto Delete Paid Fees (Weekly)

Tool tự động chạy hàng tuần để xóa dữ liệu báo phí "Đã thanh toán" trên hệ thống quản lý vận hành.
Sử dụng **GitHub Actions** để chạy hoàn toàn miễn phí và tự động.

## 🕒 Lịch chạy
*   **Thời gian:** 09:00 sáng Chủ Nhật (Giờ Việt Nam) hàng tuần.
*   **Cơ chế:** GitHub Actions tự động kích hoạt workflow.

## 📂 Cấu trúc Repository
*   `cap_nhat_xoa_da_thanh_toan.py`: Script chính thực hiện logic xóa dữ liệu.
*   `data.xlsx`: File cấu hình danh sách dự án cần xử lý.
*   `.github/workflows/run.yml`: File cấu hình lịch trình cho GitHub.
*   `requirements.txt`: Danh sách thư viện Python cần thiết.

## 📊 Dữ liệu JSON cho Website khác
Sau mỗi lần chạy, tool sẽ tự động cập nhật file `report.json`. Bạn có thể dùng website khác để fetch dữ liệu từ link sau:
*   **JSON Raw Link:** `https://raw.githubusercontent.com/itkhaservice/run-auto-weekend/main/report.json`

## 🚀 Cách chạy thủ công (Kiểm tra)
1.  Vào tab **Actions** trên GitHub Repository này.
2.  Chọn workflow **"Auto Delete Paid Fees Weekly"** ở cột bên trái.
3.  Bấm nút **Run workflow** (màu xanh lá) ở bên phải.
4.  Chờ khoảng 1-2 phút và xem kết quả trong log.

## 🛠 Cập nhật dữ liệu
Khi cần thay đổi danh sách dự án hoặc thông tin đăng nhập:
1.  Sửa file `data.xlsx` trên máy tính.
2.  Commit và Push file `data.xlsx` mới lên GitHub.
3.  Tool sẽ tự động dùng dữ liệu mới trong lần chạy tiếp theo.
