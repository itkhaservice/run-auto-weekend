import os
import sys
import subprocess
import logging
import pandas as pd
from playwright.sync_api import sync_playwright, Page
import pytest
from openpyxl import load_workbook
from datetime import datetime
from tabulate import tabulate

# Cấu hình BASE_DIR
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as p:
        is_ci = os.environ.get("CI") == "true"
        browser = p.chromium.launch(
            headless=is_ci,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-blink-features=AutomationControlled"]
        )
        yield browser
        browser.close()

@pytest.fixture
def page(browser):
    context = browser.new_context(viewport={'width': 1920, 'height': 1080})
    page = context.new_page()
    yield page
    context.close()

def login(page: Page):
    if "login" in page.url or "qlvh.khaservice.com.vn" not in page.url:
        page.goto("https://qlvh.khaservice.com.vn/login")
        page.locator("input[name='email']").fill("admin@khaservice.com.vn")
        page.locator("input[name='password']").fill("Kha@@123")
        page.locator("button[type='submit']").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

def select_project(page: Page, project_name):
    """Hàm trợ giúp chọn dự án an toàn, tránh lỗi Timeout trên GitHub Actions"""
    combo = page.locator("#combo-box-demo")
    combo.click()
    # Xóa sạch nội dung cũ trong ô tìm kiếm
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    combo.fill(str(project_name))
    
    try:
        # Đợi option đầu tiên xuất hiện thực sự
        option0 = page.locator("#combo-box-demo-option-0")
        option0.wait_for(state="visible", timeout=10000)
        option0.click()
    except:
        # Nếu lag không hiện dropdown, thử nhấn Enter
        page.keyboard.press("Enter")
    
    page.wait_for_timeout(2000) # Chờ hệ thống nạp context dự án

# --- 1. LẤY OVERVIEW (Cột B, C, D, E) ---
def test_lay_thong_tin_du_an(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()
    wb = load_workbook(excel_path)
    ws = wb["BaoCao2"]
    login(page)
    
    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Overview: {project_val}")
            select_project(page, project_val)
            
            page.locator("a[href='/statistics/overview']").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            ws[f"B{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[2]/div/div[1]/p[1]').inner_text()
            ws[f"C{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[3]/div/div[1]/p[1]').inner_text()
            ws[f"D{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[5]/div/div[1]/p[1]').inner_text()
            ws[f"E{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[6]/div/div[1]/p[1]').inner_text()
        except Exception as e:
            print(f"Lỗi Overview {project_val}: {e}")
            
    wb.save(excel_path)

def set_max_rows(page: Page):
    """Hàm hỗ trợ chọn hiển thị tối đa số dòng (1000 dòng) để lấy đủ dữ liệu"""
    try:
        # 1. Cuộn xuống cuối trang để nút phân trang hiển thị (tránh bị che bởi header/footer)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(1000)

        # 2. Tìm nút chọn số dòng (thường chứa text số dòng hiện tại, ví dụ '10')
        # Selector này nhắm vào div chứa số dòng trong Material UI Table Pagination
        pagination_trigger = page.locator("xpath=//div[contains(@class, 'MuiTablePagination-selectRoot') or contains(@class, 'MuiSelect-select')]")

        if pagination_trigger.count() > 0:
            pagination_trigger.first.scroll_into_view_if_needed()
            pagination_trigger.first.click(force=True, timeout=10000)

            # 3. Đợi menu dropdown xuất hiện và tìm option '1000'
            # Menu của MUI thường nằm trong portal ở cuối <body>
            option_1000 = page.locator("xpath=//li[contains(@class, 'MuiMenuItem-root') and (text()='1000' or .='1000')]")

            # Nếu không tìm thấy '1000', thử tìm '100' hoặc giá trị lớn nhất
            if option_1000.count() == 0:
                option_1000 = page.locator("xpath=//li[contains(@class, 'MuiMenuItem-root')]").last

            print(f"   -> Đang chọn: {option_1000.inner_text()}")
            option_1000.click(force=True)

            # 4. Đợi dữ liệu tải lại
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(3000) # Nghỉ thêm để bảng render lại xong
        else:
            print("   [!] Không tìm thấy nút chọn số dòng (có thể trang này không phân trang)")

    except Exception as e:
        print(f"   [!] Lỗi khi chọn 1000 dòng: {e}")

# --- 2. LẤY SỐ LƯỢNG BÀI VIẾT (Cột F, G) ---
def test_lay_so_luong_bai_viet(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()
    wb = load_workbook(excel_path)
    ws = wb["BaoCao2"]
    login(page)
    
    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Posts: {project_val}")
            select_project(page, project_val)

            # --- TIN TỨC ---
            page.goto("https://qlvh.khaservice.com.vn/posts/news")
            page.wait_for_load_state("networkidle")
            set_max_rows(page)
            
            count_news = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr').count()
            ws[f"F{idx}"] = count_news

            # --- THÔNG BÁO ---
            page.goto("https://qlvh.khaservice.com.vn/posts/notification")
            page.wait_for_load_state("networkidle")
            set_max_rows(page)

            count_notif = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr').count()
            ws[f"G{idx}"] = count_notif

        except Exception as e:
            print(f"Lỗi Posts {project_val}: {e}")
            
    wb.save(excel_path)

# --- 3. LẤY NGÀY BÀI VIẾT CUỐI (Cột H) ---
def test_lay_thong_tin_bai_viet_ngay_cuoi(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()
    wb = load_workbook(excel_path)
    ws = wb["BaoCao2"]
    login(page)
    
    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Ngày cuối: {project_val}")
            select_project(page, project_val)
            
            dates = []
            
            # Kiểm tra Thông báo
            page.goto("https://qlvh.khaservice.com.vn/posts/notification")
            page.wait_for_load_state("networkidle")
            loc = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[8]/div')
            if loc.is_visible():
                text = loc.inner_text().strip()
                try:
                    date_str = text.split()[0]
                    dates.append(datetime.strptime(date_str, '%d/%m/%Y'))
                except: pass
            
            # Kiểm tra Tin tức
            page.goto("https://qlvh.khaservice.com.vn/posts/news")
            page.wait_for_load_state("networkidle")
            loc = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[8]/div')
            if loc.is_visible():
                text = loc.inner_text().strip()
                try:
                    date_str = text.split()[0]
                    dates.append(datetime.strptime(date_str, '%d/%m/%Y'))
                except: pass
            
            if dates:
                max_date = max(dates).strftime('%d/%m/%Y')
                ws[f"H{idx}"] = max_date
            else:
                ws[f"H{idx}"] = "N/A"
                
        except Exception as e:
            print(f"Lỗi Date {project_val}: {e}")
            
    wb.save(excel_path)

# --- 4. LẤY BÁO PHÍ MỚI NHẤT (Cột I) ---
def test_lay_thong_tin_bao_phi_moi_nhat(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()
    wb = load_workbook(excel_path)
    ws = wb["BaoCao2"]
    login(page)
    
    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Fee Report: {project_val}")
            select_project(page, project_val)
            
            # Đảm bảo chuyển trang sau khi context dự án đã ổn định
            page.goto("https://qlvh.khaservice.com.vn/fee-reports")
            page.wait_for_load_state("networkidle", timeout=20000)
            
            # Sử dụng selector linh hoạt: Tìm td thứ 5 trong hàng đầu tiên của tbody
            # Selector này ít bị ảnh hưởng bởi các thẻ div bọc bên trong
            cell_selector = "table tbody tr:first-child td:nth-child(5)"
            
            # Thử đợi dữ liệu xuất hiện trong tối đa 20 giây
            found_data = False
            for retry in range(3): # Thử lại 3 lần nếu thấy trang trắng
                try:
                    page.wait_for_selector(cell_selector, state="visible", timeout=7000)
                    loc = page.locator(cell_selector)
                    text = loc.inner_text().strip()
                    
                    if text and text != "":
                        ws[f"I{idx}"] = text
                        print(f"   -> Phí mới nhất: {text}")
                        found_data = True
                        break
                except:
                    print(f"   [!] Thử lại lần {retry+1} cho {project_val}...")
                    page.reload()
                    page.wait_for_load_state("networkidle")

            if not found_data:
                # Kiểm tra xem có phải do 'Không có dữ liệu' thực sự không
                if "không có dữ liệu" in page.content().lower():
                    ws[f"I{idx}"] = "N/A (No Data)"
                    print("   -> Hệ thống báo: Không có dữ liệu.")
                else:
                    ws[f"I{idx}"] = "N/A (Timeout)"
                    print("   -> Lỗi: Không load được bảng dữ liệu.")
                    
        except Exception as e:
            print(f"Lỗi Fee {project_val}: {e}")

    wb.save(excel_path)

# --- XUẤT BÁO CÁO RA GITHUB ---
def test_z_summary_report():
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    if not os.path.exists(excel_path): return
    df = pd.read_excel(excel_path, sheet_name="BaoCao", keep_default_na=False)
    df = df.iloc[:, :9]
    
    # Đặt lại tên cột
    df.columns = [
        "Dự án", "Tổng căn hộ", "Tổng cư dân sử dụng APP", 
        "Tổng số căn hộ sử dụng APP", "Tổng số cư dân", 
        "Tin tức", "Thông báo", "Ngày mới nhất", "Báo phí"
    ]
    
    json_path = os.path.join(BASE_DIR, "report.json")
    df.to_json(json_path, orient='records', force_ascii=False, indent=4)
    
    table = tabulate(df, headers='keys', tablefmt='github', showindex=False)
    output = f"## 📊 Báo Cáo Tổng Hợp Dữ Liệu\n"
    output += f"*Thời gian cập nhật: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}*\n\n"
    output += table
    
    if 'GITHUB_STEP_SUMMARY' in os.environ:
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as f:
            f.write(output)
