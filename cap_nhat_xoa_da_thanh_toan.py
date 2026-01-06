from playwright.sync_api import Page
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime
import logging

# --- VÍ DỤ CẤU HÌNH VÀ HÀM HỖ TRỢ LÙI THÁNG ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
def get_previous_month(month_str):
    """Chuyển đổi chuỗi MM/YYYY thành đối tượng datetime và lùi lại 1 tháng."""
    try:
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
    ws = wb["BaoCao"]

    # 🌟 LẤY THÁNG HIỆN TẠI ĐỂ BẮT ĐẦU VÒNG LẶP
    now = pd.Timestamp.now()
    start_date = now - pd.DateOffset(months=3)
    start_month_str = start_date.strftime("%m/%Y")
    logging.info(f"Tháng bắt đầu vòng lặp: {start_month_str}")

    # 1. ĐĂNG NHẬP
    logging.info("Đang đăng nhập...")
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

    for idx, project_val in enumerate(project_list, start=2):
        print(f"\n[{idx}] Project={project_val}")
        logging.info(f"[{idx}] - Project={project_val}")

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
        try:
            # Thử tìm nút phân trang cuối cùng (thường là số trang lớn nhất hoặc nút >>)
            # Locator cũ: li[8] -> Rủi ro nếu số trang thay đổi.
            # Cải thiện: Tìm nút phân trang có số lớn nhất hoặc nút Last Page nếu có.
            # Tạm thời giữ nguyên nhưng bọc try-catch
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
            page.wait_for_timeout(1000)
        except Exception:
            logging.warning(f"[{idx}] - Không thể chuyển đến trang cuối để tìm tháng cũ nhất. Dùng mặc định.")

        try:
            # Lấy tháng cũ nhất từ cột Tháng của hàng đầu tiên (Giả sử td[5])
            thangcunhat_locator = page.locator(
                'xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
            thangcunhat = thangcunhat_locator.text_content().strip()
            logging.info(f"[{idx}] - Tháng cũ nhất được tìm thấy: {thangcunhat}")
        except Exception:
            thangcunhat = "01/2000"  # Giá trị mặc định an toàn
            logging.warning(f"[{idx}] - Lỗi khi tìm tháng cũ nhất. Đặt mặc định: {thangcunhat}")

        # Quay lại trang đầu
        try:
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[2]/button").click()
        except:
             page.reload() # Fallback nếu không click được
        page.wait_for_timeout(1000)

        # Click để mở rộng danh sách hiển thị
        try:
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[2]/button").click()
            page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[8]").click()
            page.wait_for_timeout(2000)
        except:
            logging.warning(f"[{idx}] - Không thể mở rộng danh sách hiển thị.")

        current_month_str = start_month_str  # BẮT ĐẦU TỪ THÁNG HIỆN TẠI

        # 4. VÒNG LẶP XÓA NGƯỢC THÁNG
        while True:
            # 🌟 ĐIỀU KIỆN DỪNG VÒNG LẶP
            try:
                date_current = datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y')
                date_oldest = datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y')

                # Dừng nếu tháng hiện tại nhỏ hơn tháng cũ nhất
                if date_current < date_oldest:
                    logging.info(f"[{idx}] - Đã lùi quá tháng cũ nhất ({thangcunhat}). THOÁT VÒNG LẶP.")
                    # Reset view size
                    try:
                        page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[2]/button").click()
                        page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[1]").click()
                        page.wait_for_timeout(1000)
                    except: pass
                    break
            except ValueError:
                logging.error(f"[{idx}] - Lỗi định dạng tháng. THOÁT VÒNG LẶP.")
                break

            print(f"[{idx}] Đang xử lý tháng: {current_month_str}")
            logging.info(f"[{idx}] - Đang xử lý tháng: {current_month_str}")

            try:
                # LOCATORs CHUNG
                filter_button = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                checkbox_all_locator = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input")
                
                # CẬP NHẬT: Locator nút xóa linh hoạt hơn
                # Tìm button có chứa icon delete hoặc nằm ở vị trí toolbar
                # Thử tìm theo aria-label hoặc SVG icon nếu có thể, hoặc fallback về xpath ngắn hơn
                # Ở đây ta thử tìm tất cả button trong toolbar và lọc
                delete_button_locator = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/div[2]/div/div[2]/button')

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
                    logging.info(
                        f"[{idx}] - TÌM THẤY dữ liệu Đã Thanh Toán cho tháng {current_month_str}. Bắt đầu xóa.")

                    # A. Click chọn tất cả
                    checkbox_all_locator.click()
                    
                    # QUAN TRỌNG: Chờ UI cập nhật trạng thái "Đã chọn" -> Nút xóa mới hiện
                    page.wait_for_timeout(2000) 

                    # B. KIỂM TRA NÚT XÓA VÀ THỰC HIỆN XÓA
                    if delete_button_locator.is_visible():
                        delete_button_locator.click()
                        page.wait_for_timeout(1000)

                        # C. CLICK NÚT XÁC NHẬN TRONG HỘP THOẠI
                        confirm_delete_button = page.locator("xpath=//button[@type='submit']")

                        if confirm_delete_button.is_visible():
                            confirm_delete_button.click()
                            # Chờ mạng xử lý xong request xóa (QUAN TRỌNG)
                            try:
                                page.wait_for_load_state("networkidle", timeout=5000)
                            except:
                                page.wait_for_timeout(5000)
                            
                            logging.info(f"[{idx}] - Đã XÓA thành công dữ liệu tháng {current_month_str}")
                        else:
                            logging.error(f"[{idx}] - LỖI: Không tìm thấy nút XÁC NHẬN XÓA.")
                    else:
                        # DEBUG: Nút xóa không hiện
                        logging.error(f"[{idx}] - CẢNH BÁO: Đã chọn nhưng nút XÓA không hiển thị. Bỏ qua tháng {current_month_str}.")
                        
                        # CHỤP MÀN HÌNH ĐỂ DEBUG
                        screenshot_path = f"debug_no_delete_{idx}_{current_month_str.replace('/', '_')}.png"
                        page.screenshot(path=screenshot_path)
                        logging.info(f"Đã chụp ảnh màn hình debug: {screenshot_path}")

                else:
                    logging.info(f"[{idx}] - KHÔNG TÌM THẤY dữ liệu Đã Thanh Toán cho tháng {current_month_str}.")

            except Exception as e:
                logging.error(f"[{idx}] - Lỗi bất ngờ trong vòng lặp tháng {current_month_str}: {e}")
                # Chụp ảnh lỗi ngoại lệ
                page.screenshot(path=f"error_exception_{idx}.png")

            # 5. CHUYỂN SANG THÁNG TRƯỚC
            current_month_str = get_previous_month(current_month_str)
            if current_month_str is None: break

            page.wait_for_timeout(1000)
    pass

if __name__ == "__main__":
    from playwright.sync_api import sync_playwright
    
    # Cấu hình logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("run.log", mode='w', encoding='utf-8')
        ]
    )

    logging.info(">>> BẮT ĐẦU TOOL TỰ ĐỘNG XÓA DỮ LIỆU ĐÃ THANH TOÁN (DOCKER VERSION) <<<")
    
    with sync_playwright() as p:
        # CẤU HÌNH CHO SERVER (GITHUB ACTIONS)
        # 1. Headless = True (Bắt buộc vì server không có màn hình)
        # 2. Viewport = 1920x1080 (Quan trọng: Ép giao diện hiển thị như Desktop để nút bấm không bị chạy)
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        # Tạo context với độ phân giải màn hình Full HD
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()
        
        try:
            test_xoa_du_lieu_bao_phi_da_thanh_toan(page)
        except Exception as e:
            logging.error(f"Lỗi Fatal trong quá trình chạy: {e}")
        finally:
            browser.close()
            logging.info(">>> KẾT THÚC <<<")
