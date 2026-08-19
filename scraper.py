# -*- coding: utf-8 -*-
"""
Akubela 数据爬虫 - 每天自动运行
"""

import sys
import io
import json
import time
import os
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from playwright.sync_api import sync_playwright

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class AkubelaScraper:
    def __init__(self):
        # 从环境变量读取配置
        self.username = os.getenv('AKUBELA_USERNAME', 'super')
        self.password = os.getenv('AKUBELA_PASSWORD', 'Akubela@super2024&')
        self.environment = os.getenv('AKUBELA_ENV', 'prod-cn-hz')
        
        # 邮件配置
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.exmail.qq.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '465'))
        self.sender_email = os.getenv('SENDER_EMAIL', 'enwade.chen@akuvox.com')
        self.sender_password = os.getenv('SENDER_PASSWORD', 'cyw110120')
        self.receiver_email = os.getenv('RECEIVER_EMAIL', 'enwade.chen@akuvox.com')
        
        # 数据存储
        self.device_data = []
        self.cross_region_data = []
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 创建数据目录
        Path('./data').mkdir(exist_ok=True)
        
        # 文件路径
        self.device_json = f'data/devices_{self.timestamp}.json'
        self.cross_region_json = f'data/cross_region_{self.timestamp}.json'
        self.device_excel = f'data/devices_{self.timestamp}.xlsx'
        self.cross_region_excel = f'data/cross_region_{self.timestamp}.xlsx'
    
    def run(self):
        logger.info("=" * 50)
        logger.info("开始爬取数据...")
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                
                # 1. 登录
                logger.info("正在登录...")
                page.goto("https://super.akubela.com/#/login")
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                
                page.fill('input[type="text"]:visible', self.username)
                page.fill('input[type="password"]', self.password)
                page.click('.el-button--primary')
                page.wait_for_load_state("networkidle", timeout=15000)
                time.sleep(3)
                logger.info("登录成功")
                
                # 2. 选择环境
                try:
                    if page.is_visible('div[role="combobox"]'):
                        page.click('div[role="combobox"]')
                        time.sleep(1)
                        page.click(f'text={self.environment}')
                        logger.info(f"已选择环境: {self.environment}")
                except:
                    pass
                
                # 3. 抓取设备数据
                logger.info("抓取设备数据...")
                page.goto("https://super.akubela.com/#/distributor/device")
                page.wait_for_load_state("networkidle")
                time.sleep(3)
                self.device_data = self.get_table_data(page)
                logger.info(f"设备数据: {len(self.device_data)} 条")
                
                # 4. 抓取跨区管控数据
                logger.info("抓取跨区管控数据...")
                page.goto("https://super.akubela.com/#/distributor/cross-region")
                page.wait_for_load_state("networkidle")
                time.sleep(3)
                self.cross_region_data = self.get_table_data(page)
                logger.info(f"跨区管控数据: {len(self.cross_region_data)} 条")
                
                browser.close()
            
            # 5. 保存数据
            self.save_data()
            
            # 6. 发送邮件
            self.send_email()
            
            logger.info("爬取完成！")
            logger.info("=" * 50)
            return True
            
        except Exception as e:
            logger.error(f"出错: {e}")
            return False
    
    def get_table_data(self, page):
        """提取表格数据"""
        data = []
        try:
            time.sleep(3)
            table = page.locator('table').first
            rows = table.locator('tbody tr').all()
            
            for row in rows:
                cells = row.locator('td').all()
                row_data = {}
                for i, cell in enumerate(cells):
                    row_data[f'列{i+1}'] = cell.text_content().strip()
                data.append(row_data)
        except:
            pass
        return data
    
    def save_data(self):
        """保存数据"""
        with open(self.device_json, 'w', encoding='utf-8') as f:
            json.dump(self.device_data, f, ensure_ascii=False, indent=2)
        
        with open(self.cross_region_json, 'w', encoding='utf-8') as f:
            json.dump(self.cross_region_data, f, ensure_ascii=False, indent=2)
        
        if self.device_data:
            pd.DataFrame(self.device_data).to_excel(self.device_excel, index=False)
        if self.cross_region_data:
            pd.DataFrame(self.cross_region_data).to_excel(self.cross_region_excel, index=False)
        
        logger.info("数据已保存")
    
    def send_email(self):
        """发送邮件"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email
            msg['Subject'] = f'跨区管控数据报告 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            
            body = f"""
            <h2>跨区管控数据报告</h2>
            <p>爬取时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
            <p>环境: {self.environment}</p>
            <table border="1">
                <tr><th>数据源</th><th>记录数</th></tr>
                <tr><td>设备数据</td><td>{len(self.device_data)}</td></tr>
                <tr><td>跨区管控数据</td><td>{len(self.cross_region_data)}</td></tr>
            </table>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # 添加附件
            for file in [self.device_excel, self.cross_region_excel]:
                if Path(file).exists():
                    with open(file, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{Path(file).name}"')
                        msg.attach(part)
            
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            server.login(self.sender_email, self.sender_password)
            server.sendmail(self.sender_email, self.receiver_email, msg.as_string())
            server.quit()
            logger.info("邮件发送成功")
        except Exception as e:
            logger.error(f"邮件发送失败: {e}")


if __name__ == "__main__":
    scraper = AkubelaScraper()
    scraper.run()
