from playwright.sync_api import Page, sync_playwright
import pandas as pd
import os
import logging
import time

# --- CẤU HÌNH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = "don_dep_thung_rac.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')]
)

class Colors:
    RESET = "\033[0m"
    BLUE = "\033[1;34m"
    GREEN = "\033[1;32m"

def empty_trash_logic(page: Page, project_idx, url, label):
    logging.info(f"[{project_idx}] - --- ĐANG DỌN DẸP: {label} ---")
    try:
        page.goto(f"https://qlvh.khaservice.com.vn{url}")
        page.wait_for_load_state("networkidle")
        
        while True:
            # Đếm số lượng thẻ p trong tbody - đây là chỉ số dòng dữ liệu thực tế
            p_xpath = "//*[@id='root']/div[2]/main/div/div/div[2]/table/tbody/tr/td/div/p"
            p_count = page.locator(f"xpath={p_xpath}").count()
            
            # Nếu p_count <= 1 (chỉ còn 1 dòng thông báo trống hoặc không có gì), kết thúc
            if p_count <= 1:
                logging.info(f"[{project_idx}] - Thùng rác {label} đã sạch (P-count: {p_count}).")
                break
            
            logging.info(f"[{project_idx}] - Còn dữ liệu ({p_count} dòng). Đang tiến hành xóa...")

            # 2. Chọn hiển thị 1000 dòng (li[6]) nếu là đợt đầu
            try:
                page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button").click()
                page.locator("xpath=//*[@id='menu-apartment-list-style1']/div[3]/ul/li[6]").click()
                page.wait_for_timeout(3000)
            except: pass

            # 3. Chọn tất cả và xóa
            page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/table/thead/tr/th[1]/span/input").click()
            page.wait_for_timeout(500)
            
            delete_all_btn = page.locator("xpath=//*[@id='root']/div[2]/main/div/div/div[2]/div[2]/div/div[3]/button")
            if delete_all_btn.is_visible():
                delete_all_btn.click()
                
                # 5. Bấm nút Đồng ý (Xác nhận xóa)
                confirm_xpath = "/html/body/div[2]/div[3]/div/div[2]/button[2]"
                page.locator(f"xpath={confirm_xpath}").click()
                
                logging.info(f"[{project_idx}] - Đã bấm Xác nhận xóa. Đang đợi hệ thống thực thi và nạp lại giao diện...")
                
                # Đợi nút xác nhận biến mất (modal đóng)
                try:
                    page.locator(f"xpath={confirm_xpath}").wait_for(state="hidden", timeout=5000)
                except: pass

                # QUAN TRỌNG: Đợi cho đến khi giao diện bảng nạp lại hoàn toàn
                # Đợi Combobox chọn số dòng xuất hiện trở lại
                combobox_xpath = "//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button"
                # Đợi Thẻ p nội dung xuất hiện trở lại
                p_xpath = "//*[@id='root']/div[2]/main/div/div/div[2]/table/tbody/tr/td/div/p"
                
                try:
                    page.wait_for_selector(f"xpath={combobox_xpath}", state="visible", timeout=20000)
                    page.wait_for_selector(f"xpath={p_xpath}", state="visible", timeout=20000)
                    logging.info(f"[{project_idx}] - Giao diện đã nạp lại. Đang kiểm tra dữ liệu còn lại...")
                except Exception as e:
                    logging.warning(f"[{project_idx}] - Hết thời gian đợi giao diện nạp lại: {e}")
                
                # Nghỉ thêm 1s cho ổn định hẳn trước khi quay lại đầu vòng lặp
                page.wait_for_timeout(1000) 
            else:
                break
    except Exception as e:
        logging.error(f"Lỗi dọn dẹp {label}: {e}")

    except Exception as e:
        logging.error(f"Lỗi trong quá trình dọn dẹp {label}: {e}")
def process_don_dep(project_name, project_idx):
    if str(project_name).strip().upper() == "CHUNG CƯ SEN HỒNG BC": return

    logging.info(f"{Colors.BLUE}--- [KIỂM TRA TRỰC QUAN] DỰ ÁN [{project_idx}]: {project_name} ---{Colors.RESET}")
    
    with sync_playwright() as p:
        # CẤU HÌNH ĐỂ KIỂM TRA (HẾT SỨC TRỰC QUAN)
        browser = p.chromium.launch(
            headless=False,        # HIỆN CỬA SỔ TRÌNH DUYỆT
            slow_mo=1000,          # CHẬM LẠI 1 GIÂY GIỮA CÁC THAO TÁC
            args=['--start-maximized'] # MỞ TO MÀN HÌNH
        )
        
        # Không đặt viewport cố định để dùng toàn màn hình của trình duyệt
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        try:
            # Đăng nhập
            page.goto("https://qlvh.khaservice.com.vn/login")
            page.locator("input[name='email']").fill("admin@khaservice.com.vn")
            page.locator("input[name='password']").fill("Kha@@123")
            page.locator("button[type='submit']").click()
            page.wait_for_load_state("networkidle")

            # Chọn dự án
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_name))
            page.locator("#combo-box-demo-option-0").click()
            page.wait_for_timeout(2000)

            # Dọn dẹp
            empty_trash_logic(page, project_idx, "/vehicles/trash", "THÙNG RÁC PHƯƠNG TIỆN")
            empty_trash_logic(page, project_idx, "/fee-reports/trash", "THÙNG RÁC BÁO PHÍ")


        except Exception as e:
            logging.error(f"Lỗi Fatal: {e}")
        finally:
            browser.close()

def main():
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    projects = df.iloc[1:, 0].tolist()
    
    for idx, name in enumerate(projects, 1):
        process_don_dep(name, idx)

if __name__ == "__main__":
    main()
