from playwright.sync_api import Page, sync_playwright
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime
import logging
import time

import json

# --- CẤU HÌNH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = "run.log"
JSON_LOG_FILE = "cleanup_results.json"

# Biến toàn cục để lưu kết quả chạy
execution_results = []

def save_json_log():
    """Lưu danh sách kết quả vào file JSON"""
    with open(JSON_LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(execution_results, f, ensure_ascii=False, indent=4)
    logging.info(f"--- ĐÃ CẬP NHẬT KẾT QUẢ VÀO FILE: {JSON_LOG_FILE} ---")

def empty_trash_module(page: Page, project_idx, url, label):
    """
    Hàm dùng chung để dọn dẹp thùng rác cho các module.
    Đã tối ưu hóa tốc độ load dữ liệu lớn.
    """
    logging.info(f"[{project_idx}] - --- ĐANG DỌN DẸP: {label} ---")
    batches_count = 0
    try:
        page.goto(f"https://qlvh.khaservice.com.vn{url}")
        # Đợi bảng hoặc thông báo trống xuất hiện
        page.wait_for_selector("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table", timeout=30000)
        
        while True:
            # Kiểm tra nhanh số lượng checkbox hiện có
            checkbox_list = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table//input[@type='checkbox']")
            count = checkbox_list.count()

            if count <= 1:
                logging.info(f"[{project_idx}] - Thùng rác {label} trống.")
                break
            
            batches_count += 1
            logging.info(f"[{project_idx}] - Tìm thấy {count - 1} dòng. Đợt xóa {batches_count}...")

            # 2. Chọn hiển thị 1000 dòng (Chỉ thực hiện ở đợt đầu tiên để tiết kiệm thời gian)
            if batches_count == 1:
                try:
                    page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button").click()
                    page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[6]").click()
                    # Đợi bảng cập nhật lại số dòng (đợi mạng rảnh hoặc bảng hiển thị đủ)
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(1000) # Nghỉ ngắn 1s cho chắc chắn UI cập nhật
                except: pass

            # 3. Bấm chọn tất cả các dòng
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input").click()
            
            # 4. Bấm nút Xóa tất cả các dòng đã chọn
            delete_all_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/div[2]/div/div[3]/button")
            # Đợi nút xóa hiện ra thực sự trước khi click
            delete_all_btn.wait_for(state="visible", timeout=5000)
            
            if delete_all_btn.is_visible():
                delete_all_btn.click()
                
                # 5. Bấm nút Đồng ý (Xác nhận xóa)
                confirm_xpath = "/html/body/div[2]/div[3]/div/div[2]/button[2]"
                page.wait_for_selector(f"xpath={confirm_xpath}", state="visible", timeout=5000)
                page.locator(f"xpath={confirm_xpath}").click()
                
                logging.info(f"[{project_idx}] - Đã gửi lệnh xóa đợt {batches_count}. Đang đợi nạp lại dữ liệu...")
                
                # TỐI ƯU: Đợi bảng nạp lại xong thay vì nghỉ 5s
                page.wait_for_load_state("networkidle")
                # Đợi cho đến khi checkbox "Chọn tất cả" có thể click được trở lại (dấu hiệu bảng đã load xong)
                try:
                    page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input").wait_for(state="visible", timeout=10000)
                except: pass
                page.wait_for_timeout(1000) # Nghỉ 1s đệm cuối cùng cho server ổn định
            else:
                break
        return {"status": "Completed", "batches": batches_count}
    except Exception as e:
        logging.error(f"[{project_idx}] - Lỗi dọn dẹp {label}: {e}")
        return {"status": f"Error: {str(e)}", "batches": batches_count}

def process_single_project(project_name, project_idx, start_month_str):
    """
    Hàm xử lý trọn gói cho 1 dự án duy nhất.
    """
    project_result = {
        "project_name": project_name,
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "fee_reports_cleaned": [],
        "vehicle_trash": {"status": "Skipped", "batches": 0},
        "fee_trash": {"status": "Skipped", "batches": 0}
    }

    if str(project_name).strip().upper() == "CHUNG CƯ SEN HỒNG BC":
        logging.info(f"{Colors.BLUE}--- BỎ QUA DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
        return

    logging.info(f"{Colors.BLUE}--- BẮT ĐẦU XỬ LÝ DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-setuid-sandbox'])
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()

        try:
            # 1. ĐĂNG NHẬP
            page.goto("https://qlvh.khaservice.com.vn/login")
            page.locator("input[name='email']").fill("admin@khaservice.com.vn")
            page.locator("input[name='password']").fill("Kha@@123")
            page.locator("button[type='submit']").click()
            page.wait_for_load_state("networkidle")

            # 2. CHỌN DỰ ÁN
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_name))
            page.locator("#combo-box-demo-option-0").click()
            page.wait_for_timeout(2000)

            # 3. DỌN DẸP BÁO PHÍ CŨ
            page.locator("//a[@href='/fee-reports']").click()
            page.wait_for_load_state("networkidle")

            # Lọc trạng thái "Đã thanh toán"
            try:
                filter_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]")
                filter_btn.click()
                page.locator("xpath=//*[@id='demo-simple-select-helper']").click()
                page.locator("xpath=//*[@data-value='1']").click()
                page.keyboard.press("Escape")
                page.wait_for_timeout(3000)
            except: pass

            thangcunhat = start_month_str 
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[4]/div/div[1]/nav/ul/li[8]/button").click()
                page.wait_for_timeout(2000)
                thang_loc = page.locator('xpath=//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
                if thang_loc.is_visible():
                    thangcunhat = thang_loc.text_content().strip()
            except: pass

            current_month_str = start_month_str
            while True:
                try:
                    if datetime.strptime(f"01/{current_month_str}", '%d/%m/%Y') < datetime.strptime(f"01/{thangcunhat}", '%d/%m/%Y'):
                        break
                except: break

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
                            project_result["fee_reports_cleaned"].append(current_month_str)
                            logging.info(f"[{project_idx}] - Đã xóa báo phí {current_month_str}")
                except: pass
                current_month_str = get_previous_month(current_month_str)
                if not current_month_str: break

            # --- 4. DỌN DẸP THÙNG RÁC PHƯƠNG TIỆN ---
            project_result["vehicle_trash"] = empty_trash_module(page, project_idx, "/vehicles/trash", "THÙNG RÁC PHƯƠNG TIỆN")

            # --- 5. DỌN DẸP THÙNG RÁC BÁO PHÍ ---
            project_result["fee_trash"] = empty_trash_module(page, project_idx, "/fee-reports/trash", "THÙNG RÁC BÁO PHÍ")

        except Exception as e:
            logging.error(f"[{project_idx}] - Lỗi Fatal: {e}")
        finally:
            browser.close()
            execution_results.append(project_result)
            save_json_log() # Lưu JSON sau mỗi dự án để đảm bảo không mất dữ liệu
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