from playwright.sync_api import Page, sync_playwright
import pandas as pd
import os
from datetime import datetime
import logging
import time

# --- CẤU HÌNH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = "xoa_bao_phi.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')]
)

class Colors:
    RESET = "\033[0m"
    BLUE = "\033[1;34m"
    GREEN = "\033[1;32m"
    RED = "\033[1;31m"

def get_previous_month(month_str):
    try:
        date_obj = datetime.strptime(f"01/{month_str}", '%d/%m/%Y')
        new_month = date_obj.month - 1
        new_year = date_obj.year
        if new_month == 0:
            new_month = 12
            new_year -= 1
        return f"{new_month:02d}/{new_year}"
    except ValueError: return None

def process_xoa_bao_phi(project_name, project_idx, start_month_str):
    if str(project_name).strip().upper() == "CHUNG CƯ SEN HỒNG BC": return

    logging.info(f"{Colors.BLUE}--- [XÓA BÁO PHÍ] DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox'])
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            # 1. Đăng nhập
            page.goto("https://qlvh.khaservice.com.vn/login")
            page.locator("input[name='email']").fill("admin@khaservice.com.vn")
            page.locator("input[name='password']").fill("Kha@@123")
            page.locator("button[type='submit']").click()
            page.wait_for_load_state("networkidle")

            # 2. Chọn dự án
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_name))
            page.locator("#combo-box-demo-option-0").click()
            page.wait_for_timeout(2000)

            # 3. Vào trang Báo phí
            page.goto("https://qlvh.khaservice.com.vn/fee-reports")
            page.wait_for_load_state("networkidle")

            # Pre-filter tìm tháng cũ nhất
            thangcunhat = start_month_str
            try:
                filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                filter_btn.click()
                page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                page.locator("xpath=//*[@data-value='1']").click() # Đã thanh toán
                page.keyboard.press("Escape")
                page.wait_for_timeout(3000)
                
                # Click trang cuối
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
                page.wait_for_timeout(2000)
                thangcunhat = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div').text_content().strip()
            except: pass

            # Vòng lặp xóa
            current_month_str = start_month_str
            while True:
                try:
                    if datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y') < datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y'): break
                except: break

                logging.info(f"   -> Đang kiểm tra tháng {current_month_str}...")
                try:
                    filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
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
                            logging.info(f"      {Colors.RED}[OK] Đã xóa thành công.{Colors.RESET}")
                    else:
                        logging.info("      Không có dữ liệu.")
                except: pass
                
                current_month_str = get_previous_month(current_month_str)
                if not current_month_str: break

        except Exception as e:
            logging.error(f"Lỗi Fatal: {e}")
        finally:
            browser.close()

def main():
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    now = pd.Timestamp.now()
    start_month_str = (now - pd.DateOffset(months=3)).strftime("%m/%Y")
    
    df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    projects = df.iloc[1:, 0].tolist()
    
    for idx, name in enumerate(projects, 1):
        process_xoa_bao_phi(name, idx, start_month_str)

if __name__ == "__main__":
    main()
