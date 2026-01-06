import os
import sys
import subprocess
import logging
import pandas as pd
from playwright.sync_api import sync_playwright, Page
import pytest
from openpyxl import load_workbook
from datetime import datetime

# Phần code cài đặt trình duyệt và fixtures Pytest giữ nguyên
try:
    from playwright._impl._installer import install

    install("chromium")
except Exception:
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True
        )
    except Exception as e:
        print("Không thể tải Chromium:", e)
        sys.exit(1)

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--disable-animations", "--start-maximized"]
        )
        yield browser
        browser.close()


@pytest.fixture
def page(browser):
    context = browser.new_context(no_viewport=True)
    page = context.new_page()
    yield page
    context.close()


# --- Test Case Chính đã sửa ---
def test_lay_thong_tin_du_an(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")

    # Sửa: Đọc file Excel, bỏ qua header để lấy danh sách project từ hàng 2
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    # Sửa: Lấy danh sách từ hàng thứ 2 (chỉ số 1) trở đi của cột đầu tiên (chỉ số 0)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

    # 2. Vòng lặp cập nhật danh mục
    for idx, project_val in enumerate(project_list, start=2):  # Sửa: Bắt đầu idx từ 2
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        page.locator("#combo-box-demo").click()
        page.locator("#combo-box-demo").fill(str(project_val))
        page.locator("#combo-box-demo-option-0").click()

        page.locator("a[href='/statistics/overview']").click()
        page.wait_for_timeout(500)

        # Sửa lỗi cú pháp XPath
        tong_can_ho = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[2]/div/div[1]/p[1]').inner_text()
        tong_cu_dan = page.locator('//*[@id="root"]/div[2]/main/div/div/div/div[3]/div/div[1]/p[1]').inner_text()
        tong_cu_dan_su_dung_app = page.locator(
            '//*[@id="root"]/div[2]/main/div/div/div/div[5]/div/div[1]/p[1]').inner_text()
        tong_can_ho_su_dung_app = page.locator(
            '//*[@id="root"]/div[2]/main/div/div/div/div[6]/div/div[1]/p[1]').inner_text()

        # Ghi các giá trị vào các cột B, C, D, E của hàng tương ứng với idx
        ws[f"B{idx}"] = tong_can_ho
        ws[f"C{idx}"] = tong_cu_dan
        ws[f"D{idx}"] = tong_cu_dan_su_dung_app
        ws[f"E{idx}"] = tong_can_ho_su_dung_app

    # Lưu file
    wb.save(excel_path)
    print("Đã ghi xong dữ liệu vào file Excel.")
    page.close()

def test_lay_so_luong_bai_viet_loai_tin_tuc(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    base_url = "https://qlvh.khaservice.com.vn"
    # Sửa: Đọc file Excel, bỏ qua header để lấy danh sách project từ hàng 2
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    # Sửa: Lấy danh sách từ hàng thứ 2 (chỉ số 1) trở đi của cột đầu tiên (chỉ số 0)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

    page.goto(f"{base_url}/posts/news")
    page.locator("//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button").click()
    page.locator("//*[@id='menu-apartment-list-style1']/div[3]/ul/li[6]").click()
    page.wait_for_timeout(2000)

    # 2. Vòng lặp cập nhật danh mục
    for idx, project_val in enumerate(project_list, start=2):  # Sửa: Bắt đầu idx từ 2
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        page.locator("#combo-box-demo").click()
        page.locator("#combo-box-demo").fill(str(project_val))
        page.locator("#combo-box-demo-option-0").click()

        page.wait_for_timeout(1000)
        rows = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr')
        tin_tuc_count = rows.count()

        logging.error(f"[{idx}] - Project:{project_val} - Tin tuc:{tin_tuc_count}")

        ws[f"F{idx}"] = tin_tuc_count
    wb.save(excel_path)
    print("Đã ghi xong dữ liệu vào file Excel.")
    page.close()

def test_lay_so_luong_bai_viet_loai_thong_bao(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    base_url = "https://qlvh.khaservice.com.vn"
    # Sửa: Đọc file Excel, bỏ qua header để lấy danh sách project từ hàng 2
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    # Sửa: Lấy danh sách từ hàng thứ 2 (chỉ số 1) trở đi của cột đầu tiên (chỉ số 0)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)

    page.goto(f"{base_url}/posts/notification")
    page.locator("//*[@id='root']/div[2]/main/div/div/div[3]/div/div[2]/button").click()
    page.locator("//*[@id='menu-apartment-list-style1']/div[3]/ul/li[6]").click()
    page.wait_for_timeout(2000)

    # 2. Vòng lặp cập nhật danh mục
    for idx, project_val in enumerate(project_list, start=2):  # Sửa: Bắt đầu idx từ 2
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        page.locator("#combo-box-demo").click()
        page.locator("#combo-box-demo").fill(str(project_val))
        page.locator("#combo-box-demo-option-0").click()

        page.wait_for_timeout(1000)
        rows = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr')
        notification_count = rows.count()

        logging.error(f"[{idx}] - Project:{project_val} - Tin tuc:{notification_count}")

        ws[f"G{idx}"] = notification_count
    wb.save(excel_path)
    print("Đã ghi xong dữ liệu vào file Excel.")
    page.close()

def test_lay_thong_tin_bai_viet_ngay_cuoi(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    base_url = "https://qlvh.khaservice.com.vn"
    # Sửa: Đọc file Excel, bỏ qua header để lấy danh sách project từ hàng 2
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    # Sửa: Lấy danh sách từ hàng thứ 2 (chỉ số 1) trở đi của cột đầu tiên (chỉ số 0)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)
    page.goto(f"{base_url}/posts/notification")
    page.wait_for_timeout(2000)

    # # 2. Vòng lặp cập nhật danh mục
    # Vòng lặp
    for idx, project_val in enumerate(project_list, start=2):
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        page.locator("#combo-box-demo").click()
        page.locator("#combo-box-demo").fill(str(project_val))
        page.locator("#combo-box-demo-option-0").click()
        page.wait_for_timeout(1000)  # Chờ 1 giây để trang cập nhật dữ liệu

        # Khởi tạo giá trị ban đầu là None
        ngay_trang1 = None
        ngay_trang2 = None

        # Lấy giá trị ngày giờ trên trang thông báo
        try:
            page.goto(f"{base_url}/posts/notification")
            locator_thong_bao = page.locator(
                '//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[8]/div')
            locator_thong_bao.wait_for(timeout=2000)
            ngay_trang1_str = locator_thong_bao.inner_text()
            ngay_trang1 = datetime.strptime(ngay_trang1_str.strip(), '%d/%m/%Y %H:%M')
            logging.error(f"[{idx}] - Ngày trang thông báo: {ngay_trang1_str}")
        except Exception:
            logging.error(f"[{idx}] - Không tìm thấy ngày trên trang thông báo. Bỏ qua.")

        # Lấy giá trị ngày giờ trên trang tin tức
        try:
            page.goto(f"{base_url}/posts/news")
            locator_tin_tuc = page.locator(
                '//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[8]/div'
            )
            locator_tin_tuc.wait_for(timeout=2000)
            ngay_trang2_str = locator_tin_tuc.inner_text().strip()

            # --- Chỉ lấy phần ngày tháng năm ---
            # cách 1: tách chuỗi
            ngay_trang2_date_str = ngay_trang2_str.split()[0]  # ví dụ '16/09/2025'

            # parse thành datetime để dễ xử lý
            ngay_trang2 = datetime.strptime(ngay_trang2_date_str, '%d/%m/%Y')

            logging.error(f"[{idx}] - Ngày trang tin tức: {ngay_trang2.strftime('%d/%m/%Y')}")
        except Exception:
            logging.error(f"[{idx}] - Không tìm thấy ngày trên trang tin tức. Bỏ qua.")

        # So sánh và ghi vào Excel
        if ngay_trang1 and ngay_trang2:
            # so sánh theo date thôi
            ngay_moi_nhat = max(ngay_trang1, ngay_trang2)
            ws[f"H{idx}"] = ngay_moi_nhat.strftime('%d/%m/%Y')
            logging.error(f"[{idx}] - Ngày mới nhất: {ngay_moi_nhat.strftime('%d/%m/%Y')}")
        elif ngay_trang1:
            ws[f"H{idx}"] = ngay_trang1.strftime('%d/%m/%Y')
            logging.error(f"[{idx}] - Chỉ có ngày trên trang thông báo: {ngay_trang1.strftime('%d/%m/%Y')}")
        elif ngay_trang2:
            ws[f"H{idx}"] = ngay_trang2.strftime('%d/%m/%Y')
            logging.error(f"[{idx}] - Chỉ có ngày trên trang tin tức: {ngay_trang2.strftime('%d/%m/%Y')}")
        else:
            ws[f"H{idx}"] = "Không có dữ liệu"
            logging.error(f"[{idx}] - Không có dữ liệu ngày nào được tìm thấy.")

        wb.save(excel_path)
    print("Đã ghi xong dữ liệu vào file Excel.")
    page.close()

def test_lay_thong_tin_bao_phi_moi_nhat(page: Page):
    excel_path = os.path.join(BASE_DIR, "data.xlsx")
    base_url = "https://qlvh.khaservice.com.vn"
    # Sửa: Đọc file Excel, bỏ qua header để lấy danh sách project từ hàng 2
    project_df = pd.read_excel(excel_path, sheet_name="BaoCao", header=None)
    # Sửa: Lấy danh sách từ hàng thứ 2 (chỉ số 1) trở đi của cột đầu tiên (chỉ số 0)
    project_list = project_df.iloc[1:, 0].tolist()

    wb = load_workbook(excel_path)
    ws = wb["BaoCao"]

    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Kha@@123")
    page.locator("button[type='submit']").click()
    page.wait_for_timeout(2000)
    page.goto(f"{base_url}/fee-reports")
    page.wait_for_timeout(2000)

    # # 2. Vòng lặp cập nhật danh mục
    # Vòng lặp
    for idx, project_val in enumerate(project_list, start=2):
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")

        page.locator("#combo-box-demo").click()
        page.locator("#combo-box-demo").fill(str(project_val))
        page.locator("#combo-box-demo-option-0").click()
        page.wait_for_timeout(2000)  # Chờ 1 giây để trang cập nhật dữ liệu

        # Lấy giá trị ngày giờ trên trang thông báo
        from datetime import datetime
        # ... (các import khác)

        # Đặt thangmoinhat_text = "" trước try/except để tránh lỗi khi dùng trong except nếu cần
        thangmoinhat_text = ""

        try:
            thangmoinhat_locator = page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/table/tbody/tr[1]/td[5]/div')  # Đã sửa td[5] thành td[4]
            thangmoinhat_text = thangmoinhat_locator.text_content().strip()
            logging.error(f"[{idx}] - Báo phí mới nhất: {thangmoinhat_text}")
            date_object = datetime.strptime(f"01/{thangmoinhat_text}", '%d/%m/%Y')
            ws[f"I{idx}"] = date_object.strftime('%d/%m/%Y')
            wb.save(excel_path)
        except Exception as e:
            logging.error(f"[{idx}] - Lỗi xảy ra khi xử lý/lưu phí: {e}. Bỏ qua.")
            continue
    print("Đã ghi xong dữ liệu vào file Excel.")
    page.close()

