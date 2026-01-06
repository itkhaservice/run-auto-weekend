from playwright.sync_api import Page
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime
import logging

# --- VÍ DỤ CẤU HÌNH VÀ HÀM HỖ TRỢ LÙI THÁNG ---
# Giả sử BASE_DIR đã được định nghĩa
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
def get_previous_month(month_str):
    """Chuyển đổi chuỗi MM/YYYY thành đối tượng datetime và lùi lại 1 tháng."""
    try:
        # Giả định tháng hiện tại là 02/2025
        date_obj = datetime.strptime(f"01/{month_str}", '%d/%m/%Y')
        new_month = date_obj.month - 1
        new_year = date_obj.year
        if new_month == 0:
            new_month = 12
            new_year -= 1
        return f"{new_month:02d}/{new_year}"
    except ValueError:
        return None
# --- HÀM CHÍNH TỰ ĐỘNG HÓA ---
def test_xoa_du_lieu_bao_phi_da_thanh_toan(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")

    if not os.path.exists(excel_path):
        logging.error(f"Không tìm thấy file Excel tại đường dẫn: {excel_path}")
        return

    # Load dữ liệu
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()
    wb = load_workbook(excel_path)
    ws = wb["BaoCao1"]

    # 🌟 LẤY THÁNG HIỆN TẠI ĐỂ BẮT ĐẦU VÒNG LẶP
    # Yêu cầu: Xóa từ tháng thứ 3 về trước so với hiện tại
    # Ví dụ: Hiện tại 01/2026 -> Start = 10/2025
    now = pd.Timestamp.now()
    start_date = now - pd.DateOffset(months=3)
    start_month_str = start_date.strftime("%m/%Y")
    logging.error(f"Tháng bắt đầu vòng lặp: {start_month_str}")

    # 1. ĐĂNG NHẬP
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

    for idx, project_val in enumerate(project_list, start=2):
        print(f"\n[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        # 2. CHỌN DỰ ÁN
        try:
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_val))
            page.locator("#combo-box-demo-option-0").click()
        except Exception:
            logging.error(f"[{idx}] - Lỗi khi chọn dự án {project_val}. Bỏ qua.")
            continue

        # 3. CHUYỂN ĐẾN TRANG BÁO PHÍ VÀ LẤY THÁNG CŨ NHẤT
        page.locator("//a[@href='/fee-reports']").click()
        page.wait_for_load_state("networkidle")

        # Click để chuyển sang trang cuối (tháng cũ nhất)
        page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
        page.wait_for_timeout(1000)

        try:
            # Lấy tháng cũ nhất từ cột Tháng của hàng đầu tiên (Giả sử td[5])
            thangcunhat_locator = page.locator(
                'xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
            thangcunhat = thangcunhat_locator.text_content().strip()
            logging.error(f"[{idx}] - Tháng cũ nhất được tìm thấy: {thangcunhat}")
        except Exception:
            thangcunhat = "01/2000"  # Giá trị mặc định an toàn
            logging.error(f"[{idx}] - Lỗi khi tìm tháng cũ nhất. Đặt mặc định: {thangcunhat}")

        # Quay lại trang đầu
        page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[2]/button").click()
        page.wait_for_timeout(1000)

        # Click để mở rộng danh sách hiển thị
        page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[2]/button").click()
        page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[8]").click()
        page.wait_for_timeout(2000)

        current_month_str = start_month_str  # BẮT ĐẦU TỪ THÁNG HIỆN TẠI

        # 4. VÒNG LẶP XÓA NGƯỢC THÁNG
        while True:
            # 🌟 ĐIỀU KIỆN DỪNG VÒNG LẶP (Kiểm tra xem đã lùi quá tháng cũ nhất chưa)
            try:
                date_current = datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y')
                date_oldest = datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y')

                # Dừng nếu tháng hiện tại nhỏ hơn tháng cũ nhất
                if date_current < date_oldest:
                    logging.error(f"[{idx}] - Đã lùi quá tháng cũ nhất ({thangcunhat}). THOÁT VÒNG LẶP.")
                    # Click để thu nhỏ danh sách hiển thị
                    page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[2]/button").click()
                    page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[1]").click()
                    page.wait_for_timeout(2000)
                    break
            except ValueError:
                logging.error(f"[{idx}] - Lỗi định dạng tháng trong quá trình so sánh. THOÁT VÒNG LẶP.")
                break

            print(f"[{idx}] Đang xử lý tháng: {current_month_str}")
            logging.error(f"[{idx}] - Đang xử lý tháng: {current_month_str}")

            try:
                # LOCATORs CHUNG
                filter_button = page.locator(
                    "xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                checkbox_all_locator = page.locator(
                    "xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input")
                delete_button_locator = page.locator(
                    'xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/div[2]/div/div[2]/button')

                # 4.1. MỞ FILTER và ÁP DỤNG LỌC
                filter_button.click()
                page.wait_for_timeout(500)

                # ÁP DỤNG THÁNG MỚI VÀ TRẠNG THÁI 'Đã thanh toán'
                page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                page.locator("xpath=//*[@data-value='1']").click()
                page.locator("xpath=//*[@placeholder='MM/YYYY']").fill(current_month_str)
                page.keyboard.press("Escape")

                page.wait_for_timeout(3000)  # Đợi dữ liệu load sau khi filter

                # 4.2. KIỂM TRA DỮ LIỆU VÀ XÓA

                if checkbox_all_locator.is_visible():
                    logging.error(
                        f"[{idx}] - TÌM THẤY dữ liệu Đã Thanh Toán cho tháng {current_month_str}. Bắt đầu xóa.")

                    # A. Click chọn tất cả
                    checkbox_all_locator.click()
                    page.wait_for_timeout(500)

                    # B. KIỂM TRA NÚT XÓA VÀ THỰC HIỆN XÓA
                    if delete_button_locator.is_visible():
                        delete_button_locator.click()
                        page.wait_for_timeout(1000)

                        # C. CLICK NÚT XÁC NHẬN TRONG HỘP THOẠI
                        confirm_delete_button = page.locator("xpath=//button[@type='submit']")

                        if confirm_delete_button.is_visible():
                            confirm_delete_button.click()
                            page.wait_for_timeout(3000)
                            logging.error(f"[{idx}] - Đã XÓA thành công dữ liệu tháng {current_month_str}")
                        else:
                            logging.error(f"[{idx}] - LỖI: Không tìm thấy nút XÁC NHẬN XÓA.")
                    else:
                        logging.error(
                            f"[{idx}] - CẢNH BÁO: Đã chọn nhưng nút XÓA không hiển thị. Bỏ qua tháng {current_month_str}.")


                else:
                    logging.error(f"[{idx}] - KHÔNG TÌM THẤY dữ liệu Đã Thanh Toán cho tháng {current_month_str}.")

            except Exception as e:
                # Bắt lỗi chung trong quá trình thao tác hoặc xóa
                logging.error(
                    f"[{idx}] - Lỗi bất ngờ trong vòng lặp tháng {current_month_str}: {e}. Chuyển sang tháng trước.")

            # 5. CHUYỂN SANG THÁNG TRƯỚC
            current_month_str = get_previous_month(current_month_str)
            if current_month_str is None: break

            page.wait_for_timeout(1000)

    # Không đóng page ở đây để wrapper bên ngoài quản lý hoặc đóng nếu chạy độc lập
    pass

if __name__ == "__main__":
    from playwright.sync_api import sync_playwright
    
    # Cấu hình logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler()]
    )

    logging.info(">>> BẮT ĐẦU TOOL TỰ ĐỘNG XÓA DỮ LIỆU ĐÃ THANH TOÁN (DOCKER VERSION) <<<")
    
    with sync_playwright() as p:
        # Chạy headless=True trong môi trường Docker
        # args=['--no-sandbox'] là bắt buộc đối với Docker
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-setuid-sandbox'])
        page = browser.new_page()
        
        try:
            test_xoa_du_lieu_bao_phi_da_thanh_toan(page)
        except Exception as e:
            logging.error(f"Lỗi Fatal trong quá trình chạy: {e}")
        finally:
            browser.close()
            logging.info(">>> KẾT THÚC <<<")