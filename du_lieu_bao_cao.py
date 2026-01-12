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
    BASE_DIR = os.path.dirname(os.path.abspath(__file__)))

@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as p:
        # Chạy headless=True để phù hợp với GitHub Actions
        browser = p.chromium.launch(
            headless=True,
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
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

# --- Các hàm lấy thông tin dữ liệu ---

def test_lay_thong_tin_du_an(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    login(page)

    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Overview: {project_val}")
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_val))
            page.locator("#combo-box-demo-option-0").click()

            page.locator("a[href='/statistics/overview']").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)

            ws[f"B{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[2]/div/div[1]/p[1]').inner_text()
            ws[f"C{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[3]/div/div[1]/p[1]').inner_text()
            ws[f"D{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[5]/div/div[1]/p[1]').inner_text()
            ws[f"E{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[6]/div/div[1]/p[1]').inner_text()
        except Exception as e:
            print(f"Lỗi Overview tại {project_val}: {e}")

    wb.save(excel_path)

def test_lay_so_luong_bai_viet(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    login(page)

    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Posts: {project_val}")
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_val))
            page.locator("#combo-box-demo-option-0").click()
            page.wait_for_timeout(1000)

            # Lấy tin tức
            page.goto("https://qlvh.khaservice.com.vn/posts/news")
            page.wait_for_load_state("networkidle")
            ws[f"F{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr').count()

            # Lấy thông báo
            page.goto("https://qlvh.khaservice.com.vn/posts/notification")
            page.wait_for_load_state("networkidle")
            ws[f"G{idx}"] = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr').count()
        except Exception as e:
            print(f"Lỗi Posts tại {project_val}: {e}")

    wb.save(excel_path)

def test_lay_thong_tin_bao_phi_moi_nhat(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    login(page)

    for idx, project_val in enumerate(project_list, start=2):
        try:
            print(f"[{idx}] Đang lấy Fee Report: {project_val}")
            page.locator("#combo-box-demo").click()
            page.locator("#combo-box-demo").fill(str(project_val))
            page.locator("#combo-box-demo-option-0").click()
            page.wait_for_timeout(2000)

            page.goto("https://qlvh.khaservice.com.vn/fee-reports")
            page.wait_for_load_state("networkidle")
            
            thangmoinhat_locator = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')
            if thangmoinhat_locator.is_visible():
                text = thangmoinhat_locator.text_content().strip()
                ws[f"I{idx}"] = text
        except Exception as e:
            print(f"Lỗi Fee tại {project_val}: {e}")

    wb.save(excel_path)

# --- BƯỚC CUỐI: XUẤT BÁO CÁO RA GITHUB ---
def test_z_summary_report():
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    if not os.path.exists(excel_path):
        return

    df = pd.read_excel(excel_path, sheet_name="BaoCao")
    
    # Định dạng bảng Markdown
    table = tabulate(df, headers='keys', tablefmt='github', showindex=False)
    
    output_content = f"## 📊 Kết Quả Báo Cáo Dữ Liệu\n"
    output_content += f"*Cập nhật lúc: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}*\n\n"
    output_content += table
    output_content += "\n\n---\n💡 **Mẹo:** Bạn có thể bôi đen bảng trên, nhấn `Ctrl+C` và dán trực tiếp vào Excel."

    # Xuất ra GitHub Job Summary
    if 'GITHUB_STEP_SUMMARY' in os.environ:
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as f:
            f.write(output_content)
    else:
        print("\n" + output_content + "\n")