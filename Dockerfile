# Sử dụng image chính thức của Playwright (đã bao gồm Python, Node, và Browsers)
FROM mcr.microsoft.com/playwright/python:v1.49.0-jammy

# Thiết lập múi giờ Việt Nam
ENV TZ=Asia/Ho_Chi_Minh
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# Thiết lập thư mục làm việc
WORKDIR /app

# Copy file requirements trước để tận dụng Docker cache
COPY requirements.txt .

# Cài đặt các thư viện Python
RUN pip install --no-cache-dir -r requirements.txt

# Cài đặt Cron (Trình lập lịch)
RUN apt-get update && apt-get install -y cron

# Copy toàn bộ code và file data vào container
COPY cap_nhat_xoa_da_thanh_toan.py .
COPY data.xlsx .

# Copy và cấu hình file crontab
COPY crontab /etc/cron.d/scheduler
RUN chmod 0644 /etc/cron.d/scheduler
RUN crontab /etc/cron.d/scheduler

# Tạo file log trống để cron ghi vào
RUN touch /app/cron.log

# Chạy Cron ở chế độ foreground (chạy mãi mãi)
CMD ["cron", "-f"]
