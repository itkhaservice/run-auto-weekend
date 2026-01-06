from playwright.sync_api import Page, sync_playwright
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime
import logging
import time

# --- CẤU HÌNH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = "run.log"

# Cấu hình Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
    ]
)

# MÃ MÀU CHO GITHUB ACTIONS (ANSI CODES)
class Colors:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    BLUE = "\033[1;34m"   # Xanh lam đậm
    GREEN = "\033[1;32m"  # Xanh lá đậm
    RED = "\033[1;31m"    # Đỏ đậm


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

def process_single_project(project_name, project_idx, start_month_str):
    """
    Hàm xử lý trọn gói cho 1 dự án duy nhất.
    Mở browser -> Login -> Xử lý -> Đóng browser tự động (khi hết with).
    """
    logging.info(f"{Colors.BLUE}--- BẮT ĐẦU XỬ LÝ DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
    
    with sync_playwright() as p:
        # Cấu hình Browser cho GitHub Actions (Server)
        # headless=True: Chạy ẩn
        # Viewport cố định 1920x1080: Giả lập màn hình Desktop chuẩn để tránh lỗi Responsive
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        
        # Ép độ phân giải Full HD (Quan trọng hơn start-maximized khi chạy headless)
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            # 1. ĐĂNG NHẬP (Thực hiện lại cho mỗi phiên để đảm bảo session tươi mới)
            page.goto("https://qlvh.khaservice.com.vn/login")
            page.locator("input[name='email']").fill("admin@khaservice.com.vn")
            page.locator("input[name='password']").fill("Kha@@123")
            page.locator("button[type='submit']").click()
            
            # Chờ đăng nhập thành công (Linh hoạt: Chỉ cần thoát khỏi trang login)
            try:
                page.wait_for_url(lambda u: "login" not in u, timeout=30000)
                page.wait_for_load_state("networkidle")
                logging.info(f"[{project_idx}] Đăng nhập thành công. URL: {page.url}")
            except Exception as e:
                logging.warning(f"[{project_idx}] Cảnh báo: Hết thời gian chờ chuyển trang, nhưng vẫn thử tiếp tục. Lỗi: {e}")

            # 2. CHỌN DỰ ÁN
            try:
                page.locator("#combo-box-demo").click()
                page.locator("#combo-box-demo").fill(str(project_name))
                page.locator("#combo-box-demo-option-0").click()
                page.wait_for_timeout(2000) # Chờ dự án load context
            except Exception as e:
                logging.error(f"[{project_idx}] - Lỗi khi chọn dự án {project_name}: {e}")
                return # Bỏ qua dự án này

            # 3. CHUYỂN ĐẾN TRANG BÁO PHÍ
            page.locator("//a[@href='/fee-reports']").click()
            page.wait_for_load_state("networkidle")

            # --- [NEW] PRE-FILTER: Lọc sẵn trạng thái "Đã thanh toán" để xác định tháng chính xác ---
            try:
                # Mở Filter
                filter_btn = page.locator(
                    "xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                
                # CẬP NHẬT: Chờ nút hiện thay vì kiểm tra tức thì
                try:
                    filter_btn.wait_for(state="visible", timeout=5000)
                    filter_btn.click()
                except:
                    logging.warning(f"[{project_idx}] - Không thấy nút Filter sau 5s chờ.")
                    # Nếu không mở được filter thì skip đoạn dưới
                    raise Exception("Skip Filter interaction")

                page.wait_for_timeout(500)

                # Set điều kiện lọc
                page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                page.locator("xpath=//*[@data-value='1']").click()  # Đã thanh toán
                page.keyboard.press("Escape")

                logging.info(f"[{project_idx}] - Đã Pre-Filter trạng thái 'Đã thanh toán'.")
                # Chờ load dữ liệu sau lọc
                try:
                    page.wait_for_load_state("networkidle", timeout=5000)
                except: pass
                page.wait_for_timeout(6000)
            except Exception as e:
                logging.warning(f"[{project_idx}] - Pre-Filter chưa hoàn tất: {e}")

            # --- TÌM THÁNG CŨ NHẤT ---
            # Nếu không tìm thấy, mặc định gán bằng Tháng bắt đầu -> Để vòng lặp không chạy
            thangcunhat = start_month_str 
            
            try:
                # Click sang trang cuối
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
                page.wait_for_timeout(2000)
                
                # Lấy dữ liệu
                thangcunhat_locator = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
                if thangcunhat_locator.is_visible():
                    thangcunhat_web = thangcunhat_locator.text_content().strip()
                    # Kiểm tra xem text lấy được có đúng định dạng ngày tháng không
                    datetime.strptime(f"01/{thangcunhat_web}", '%d/%m/%Y')
                    thangcunhat = thangcunhat_web
                    logging.info(f"[{project_idx}] - Tháng cũ nhất (từ Web): {thangcunhat}")
                else:
                    logging.info(f"[{project_idx}] - Không thấy dữ liệu. Có thể dự án chưa có phiếu Đã thanh toán nào.")
            except Exception:
                logging.warning(f"[{project_idx}] - Lỗi khi tìm tháng cũ nhất. Sẽ bỏ qua dự án này.")

            # Quay lại trang đầu
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[2]/button").click()
                page.wait_for_timeout(1000)
            except: 
                page.reload()

            # Mở rộng danh sách hiển thị
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[2]/button").click()
                page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[8]").click()
                page.wait_for_timeout(2000)
            except: pass

            # --- VÒNG LẶP XÓA THÁNG ---
            current_month_str = start_month_str
            
            while True:
                # Kiểm tra điều kiện dừng
                try:
                    date_current = datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y')
                    date_oldest = datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y')
                    if date_current < date_oldest:
                        logging.info(f"[{project_idx}] - Đã xử lý xong đến tháng cũ nhất ({thangcunhat}).")
                        break
                except ValueError:
                    break

                # Bắt đầu xây dựng nội dung log cho tháng này
                status_msg = f"[{project_idx}] - Tháng {current_month_str}: "

                try:
                    # Mở Filter
                    filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                    if filter_btn.is_visible():
                        filter_btn.click()
                        page.wait_for_timeout(500)
                        
                        # Set điều kiện lọc
                        page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                        page.locator("xpath=//*[@data-value='1']").click() # Đã thanh toán
                        page.locator("xpath=//*[@placeholder='MM/YYYY']").fill(current_month_str)
                        page.keyboard.press("Escape")
                        
                        # Chờ load dữ liệu sau lọc
                        try:
                            page.wait_for_load_state("networkidle", timeout=5000)
                        except: pass
                        page.wait_for_timeout(6000)
                        
                        checkbox_all = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input")
                        
                        if checkbox_all.is_visible():
                            status_msg += "Đang xóa... "
                            checkbox_all.click()
                            page.wait_for_timeout(2000) # Chờ nút xóa hiện
                            
                            delete_btn = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/div[2]/div/div[2]/button')
                            
                            if delete_btn.is_visible():
                                delete_btn.click()
                                page.wait_for_timeout(1000)
                                
                                confirm_btn = page.locator("xpath=//button[@type='submit']")
                                if confirm_btn.is_visible():
                                    confirm_btn.click()
                                    try:
                                        page.wait_for_load_state("networkidle", timeout=5000)
                                    except:
                                        page.wait_for_timeout(4000)
                                    status_msg += f"{Colors.RED}XÓA THÀNH CÔNG{Colors.RESET}"
                                else:
                                    status_msg += "Lỗi (Không thấy nút Xác nhận)"
                            else:
                                status_msg += "Không tìm thấy nút Xóa... Bỏ qua"
                        else:
                            status_msg += "Không có dữ liệu."
                    
                    logging.info(status_msg)
                    
                except Exception as inner_e:
                    logging.error(f"[{project_idx}] - Lỗi tháng {current_month_str}: {inner_e}")

                # Lùi tháng
                current_month_str = get_previous_month(current_month_str)
                if current_month_str is None: break
                page.wait_for_timeout(500)

        except Exception as e:
            logging.error(f"[{project_idx}] - Lỗi Fatal dự án {project_name}: {e}")
        finally:
            # BƯỚC QUAN TRỌNG NHẤT: Đóng Browser để giải phóng RAM
            browser.close()
            logging.info(f"{Colors.GREEN}--- Đã hoàn thành và đóng Browser cho dự án {project_name} ---{Colors.RESET}\n")

def main_orchestrator():
    """Hàm điều phối chính: Đọc Excel và gọi xử lý từng dự án"""
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    if not os.path.exists(excel_path):
        logging.error("Không tìm thấy file data.xlsx")
        return

    # Tính toán tháng bắt đầu
    now = pd.Timestamp.now()
    start_date = now - pd.DateOffset(months=3)
    start_month_str = start_date.strftime("%m/%Y")
    logging.info(f">>> TOOL STARTED. Start Month: {start_month_str} <<<")

    try:
        project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
        project_list = project_df.iloc[1:, 0].tolist()
        
        logging.info(f"Tổng số dự án cần xử lý: {len(project_list)}")
        
        for idx, project_val in enumerate(project_list, start=1):
            # Gọi hàm xử lý riêng biệt cho từng dự án
            process_single_project(project_val, idx, start_month_str)
            
            # Nghỉ ngắn giữa các dự án để CPU server "thở"
            time.sleep(2)
            
    except Exception as e:
        logging.error(f"Lỗi khi đọc file Excel hoặc khởi tạo: {e}")

if __name__ == "__main__":
    main_orchestrator()
    logging.info(">>> HOÀN TẤT TOÀN BỘ CÔNG VIỆC <<<")