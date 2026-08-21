name: Daily Scraper

on:
  schedule:
    - cron: '0 0 * * *' # 每天午夜运行
  workflow_dispatch:     # 允许手动触发

jobs:
  scrape:
    runs-on: ubuntu-latest
    
    steps:
    - uses: actions/checkout@v4
    
    - name: Setup Python
      uses: actions/setup-python@v5
      with:
        python-version: '3.10'
        cache: 'pip'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install playwright pandas openpyxl  # 确保安装了 openpyxl 以支持 Excel 保存
    
    - name: Install Playwright Browser
      run: |
        # 这个命令会同时解决 libnss3 等所有 Linux 系统依赖问题，完美规避 apt-get 卡死
        python -m playwright install --with-deps chromium
    
    - name: Run scraper
      env: # 将 GitHub Secrets 映射到环境变量供 Python 读取
        AKUBELA_USERNAME: ${{ secrets.AKUBELA_USERNAME }}
        AKUBELA_PASSWORD: ${{ secrets.AKUBELA_PASSWORD }}
        AKUBELA_ENV: 'prod-cn-hz'
        SMTP_SERVER: 'smtp.exmail.qq.com'
        SMTP_PORT: '465'
        SENDER_EMAIL: ${{ secrets.SENDER_EMAIL }}
        SENDER_PASSWORD: ${{ secrets.SENDER_PASSWORD }}
        RECEIVER_EMAIL: ${{ secrets.RECEIVER_EMAIL }}
      run: python scraper.py
    
    - name: Upload data
      if: always()
      uses: actions/upload-artifact@v4
      with:
        name: scraper-data
        path: data/ # 上传 data 文件夹里的所有 JSON 和 Excel 结果
