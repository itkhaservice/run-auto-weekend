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

def empty_trash_module(page: Page, project_idx, url, label):
    """
    Hàm dùng chung để dọn dẹp thùng rác cho các module.
    Sử dụng XPath chuẩn đã kiểm tra ổn định.
    """
    logging.info(f"[{project_idx}] - --- ĐANG DỌN DẸP: {label} ---")
    try:
        page.goto(f"https://qlvh.khaservice.com.vn{url}")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        while True:
            # 1. Kiểm tra sự tồn tại của checkbox và dòng dữ liệu
            checkbox_list = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table//input[@type='checkbox']")
            count = checkbox_list.count()

            if count <= 1:
                logging.info(f"[{project_idx}] - Thùng rác {label} trống.")
                break

            logging.info(f"[{project_idx}] - Tìm thấy {count - 1} dòng. Đang tiến hành xóa...")

            # 2. Chọn hiển thị 100 dòng
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button").click()
                page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[4]").click()
                page.wait_for_timeout(2000)
            except: pass

            # 3. Bấm chọn tất cả các dòng
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input").click()
            page.wait_for_timeout(1000)

            # 4. Bấm nút Xóa tất cả các dòng đã chọn
            delete_all_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/div[2]/div/div[3]/button")
            if delete_all_btn.is_visible():
                delete_all_btn.click()
                page.wait_for_timeout(1000)

                # 5. Bấm nút Đồng ý (Xác nhận xóa)
                confirm_btn = page.locator("xpath=/html/body/div[2]/div[3]/div/div[2]/button[2]")
                if confirm_btn.is_visible():
                    confirm_btn.click()
                    logging.info(f"[{project_idx}] - Đã xóa xong 1 đợt dữ liệu...")
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(5000)
                else: break
            else: break

    except Exception as e:
        logging.error(f"[{project_idx}] - Lỗi dọn dẹp {label}: {e}")
def process_single_project(project_name, project_idx, start_month_str):
    """
    Hàm xử lý trọn gói cho 1 dự án duy nhất.
    Mở browser -> Login -> Xử lý -> Đóng browser tự động (khi hết with).
    """
    if str(project_name).strip().upper() == "CHUNG CƯ SEN HỒNG BC":
        logging.info(f"{Colors.BLUE}--- BỎ QUA DỰ ÁN [{project_idx}]: {project_name} (Theo yêu cầu) ---{Colors.RESET}")
        return

    logging.info(f"{Colors.BLUE}--- BẮT ĐẦU XỬ LÝ DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
    
    with sync_playwright() as p:
        # Cấu hình Browser cho GitHub Actions (Server)
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            # 1. ĐĂNG NHẬP
            page.goto("https://qlvh.khaservice.com.vn/login")
            page.locator("input[name='email']").fill("admin@khaservice.com.vn")
            page.locator("input[name='password']").fill("Kha@@123")
            page.locator("button[type='submit']").click()
            
            try:
                page.wait_for_url(lambda u: "login" not in u, timeout=30000)
                page.wait_for_load_state("networkidle")
                logging.info(f"[{project_idx}] Đăng nhập thành công.")
            except Exception as e:
                logging.warning(f"[{project_idx}] Cảnh báo: Đăng nhập chậm hoặc lỗi: {e}")

            # 2. CHỌN DỰ ÁN
            try:
                page.locator("#combo-box-demo").click()
                page.locator("#combo-box-demo").fill(str(project_name))
                page.locator("#combo-box-demo-option-0").click()
                page.wait_for_timeout(2000)
            except Exception as e:
                logging.error(f"[{project_idx}] - Lỗi khi chọn dự án {project_name}: {e}")
                return

            # 3. DỌN DẸP BÁO PHÍ CŨ (ĐỊNH KỲ)
            page.locator("//a[@href='/fee-reports']").click()
            page.wait_for_load_state("networkidle")

            # [Đoạn mã xóa báo phí cũ đã có sẵn...]
            # (Tôi sẽ giữ nguyên logic xóa báo phí cũ ở đây)
            # --- [NEW] PRE-FILTER ---
            try:
                filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                filter_btn.wait_for(state="visible", timeout=5000)
                filter_btn.click()
                page.wait_for_timeout(500)
                page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                page.locator("xpath=//*[@data-value='1']").click()  # Đã thanh toán
                page.keyboard.press("Escape")
                page.wait_for_timeout(3000)
            except: pass

            thangcunhat = start_month_str 
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
                page.wait_for_timeout(2000)
                thangcunhat_locator = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
                if thangcunhat_locator.is_visible():
                    thangcunhat = thangcunhat_locator.text_content().strip()
            except: pass

            current_month_str = start_month_str
            while True:
                try:
                    if datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y') < datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y'):
                        break
                except: break

                try:
                    # Tái sử dụng Filter cho từng tháng
                    filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                    if filter_btn.is_visible():
                        filter_btn.click()
                        page.locator("xpath=//*[@placeholder='MM/YYYY']").fill(current_month_str)
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(4000)
                        
                        checkbox_all = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input")
                        if checkbox_all.is_visible():
                            checkbox_all.click()
                            page.wait_for_timeout(1000)
                            delete_btn = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/div[2]/div/div[2]/button')
                            if delete_btn.is_visible():
                                delete_btn.click()
                                page.locator("xpath=//button[@type='submit']").click()
                                page.wait_for_timeout(3000)
                                logging.info(f"[{project_idx}] - Đã xóa báo phí tháng {current_month_str}")
                except: pass
                current_month_str = get_previous_month(current_month_str)
                if not current_month_str: break

            # --- 4. DỌN DẸP THÙNG RÁC PHƯƠNG TIỆN ---
            empty_trash_module(page, project_idx, "/vehicles/trash", "THÙNG RÁC PHƯƠNG TIỆN")

            # --- 5. DỌN DẸP THÙNG RÁC BÁO PHÍ ---
            empty_trash_module(page, project_idx, "/fee-reports/trash", "THÙNG RÁC BÁO PHÍ")

        except Exception as e:
            logging.error(f"[{project_idx}] - Lỗi Fatal: {e}")
        finally:
            browser.close()
            logging.info(f"{Colors.GREEN}--- Hoàn tất dự án {project_name} ---{Colors.RESET}\n")

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