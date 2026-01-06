==================================================
HƯỚNG DẪN CHẠY PLAYWRIGHT (PYTHON) HÀNG TUẦN
FREE – KHÔNG CẦN BẬT PC
(Docker + cron + Oracle Cloud Free)

MỤC TIÊU

Project Playwright + Python đơn giản (cap_nhat_xoa_da_thanh_toan.py)

Chạy tự động mỗi CHỦ NHẬT hàng tuần

Không cần bật máy tính cá nhân

Sử dụng server FREE

Cài 1 lần → chạy mãi

I. TỔNG QUAN CÁCH HOẠT ĐỘNG

Code Playwright được đóng gói vào Docker

Docker chạy trên server Oracle Cloud Free

Bên trong Docker có cron

Cron tự động chạy script mỗi Chủ nhật

PC bạn chỉ dùng để:

Viết code

Upload code lên server
Sau đó có thể tắt PC

II. CẤU TRÚC PROJECT (LOCAL)

Tạo 1 thư mục, ví dụ:

playwright_auto/

Bên trong có:

main.py

requirements.txt

Dockerfile

crontab

III. NỘI DUNG CÁC FILE

File requirements.txt

Nội dung:

playwright

File main.py (ví dụ tối thiểu)

from playwright.sync_api import sync_playwright

def run():
with sync_playwright() as p:
browser = p.chromium.launch(headless=True)
page = browser.new_page()
page.goto("https://example.com
")
print("PLAYWRIGHT DONE")
browser.close()

if name == "main":
run()

GHI CHÚ:

headless=True (bắt buộc khi chạy server)

Có thể thay code này bằng script thật của bạn

File crontab (chạy mỗi Chủ nhật)

Ví dụ: chạy 1h sáng (UTC)

0 1 * * 0 python /app/main.py >> /app/log.txt 2>&1

Giải thích:

0 1 : 01:00

: mọi ngày trong tháng

0 : Chủ nhật

log.txt dùng để kiểm tra kết quả

LƯU Ý GIỜ:

Oracle Cloud dùng UTC

Giờ Việt Nam = UTC +7

Nếu muốn chạy 8h sáng CN VN → để 1h UTC

File Dockerfile

FROM mcr.microsoft.com/playwright/python:v1.42.0-jammy

WORKDIR /app
COPY . .

RUN pip install -r requirements.txt

RUN apt-get update && apt-get install -y cron
COPY crontab /etc/cron.d/playwright-cron
RUN chmod 0644 /etc/cron.d/playwright-cron
RUN crontab /etc/cron.d/playwright-cron

CMD ["cron", "-f"]

GHI CHÚ:

Image này đã có sẵn Chromium + Playwright

Không cần cài browser thủ công

cron chạy nền liên tục

IV. TẠO SERVER ORACLE CLOUD FREE

Đăng ký Oracle Cloud (Free Tier)

Tạo Compute Instance:

OS: Ubuntu 22.04

Shape: VM.Standard.A1.Flex

CPU: 1 core

RAM: 6 GB

Tải SSH key về máy

V. KẾT NỐI SSH VÀ CÀI DOCKER

SSH vào server:

ssh ubuntu@IP_SERVER

Cài Docker:

sudo apt update
sudo apt install -y docker.io
sudo systemctl enable docker
sudo systemctl start docker

Kiểm tra:

docker --version

VI. UPLOAD PROJECT LÊN SERVER

Cách đơn giản nhất: SCP

scp -r playwright_auto ubuntu@IP_SERVER:/home/ubuntu/

Sau đó SSH vào server:

cd ~/playwright_auto

VII. BUILD & CHẠY DOCKER

Build image:

docker build -t playwright-auto .

Chạy container:

docker run -d --restart unless-stopped playwright-auto

Giải thích:

-d : chạy nền

--restart unless-stopped : server reboot vẫn tự chạy

VIII. KIỂM TRA LOG

Xem container đang chạy:

docker ps

Xem log Docker:

docker logs <container_id>

Hoặc vào log file:

docker exec -it <container_id> cat /app/log.txt

IX. CÁCH CHỈNH GIỜ CHẠY SAU NÀY

Sửa file crontab (local)

Upload lại project

Build lại image

Xóa container cũ

Run container mới

X. LƯU Ý QUAN TRỌNG

Playwright phải chạy headless

Không dùng sleep / display

Không cần domain

Không cần web server

Server free đủ mạnh cho automation

XI. TỔNG KẾT

Giải pháp này:

FREE

Ổn định

Chạy hàng tuần

Không cần bật PC

Dùng được lâu dài

CHỈ CẦN:

Docker

Cron

Oracle Cloud Free

==================================================