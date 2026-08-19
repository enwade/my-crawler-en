# -*- coding: utf-8 -*-
"""
Akubela 数据爬虫 - 每天自动运行 (融合2.py强大抓取逻辑)
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
        
        page = None
        browser = None
        
        try:
            with sync_playwright() as p:
                # 核心改动1：沿用2.py的稳定浏览器配置
                browser = p.chromium.launch(
                    headless=True, 
                    slow_mo=500  # 增加操作间隔，模拟真人，减少因为快导致的定位失败
                )
                context = browser.new_context(viewport={'width': 1920, 'height': 1080})
                page = context.new_page()
                
                # 1. 登录 (采用2.py的容错机制)
                logger.info("正在登录...")
                self.login(page)
                
                # 2. 选择环境 (采用2.py的容错机制)
                logger.info("选择环境...")
                self.select_environment(page)
                
                # 3. 抓取设备数据 (授权状态=已限制)
                logger.info("抓取设备数据 (筛选: 已限制)...")
                self.scrape_data1(page)
                
                # 4. 抓取跨区管控数据 (风险等级=高)
                logger.info("抓取跨区管控数据 (筛选: 高)...")
                self.scrape_data2(page)
                
                browser.close()
            
            # 5. 保存数据
            self.save_data()
            
            # 6. 发送邮件
            self.send_email()
            
            logger.info("爬取完成！")
            logger.info("=" * 50)
            return True
            
        except Exception as e:
            logger.error(f"运行出错: {e}")
            if page:
                try:
                    page.screenshot(path="error_screenshot.png")
                    logger.info("已保存错误截图: error_screenshot.png")
                except:
                    pass
            return False

    # ---------------------------------------------------------
    # 以下核心逻辑直接融合 2.py 的强大实现
    # ---------------------------------------------------------

    def login(self, page):
        page.goto("https://super.akubela.com/#/login")
        page.wait_for_load_state("networkidle")
        time.sleep(3)
        
        # 尝试多种方式定位用户名
        username_selectors = ['input[placeholder*="用户名"]', 'input[placeholder*="账号"]', 'input[type="text"]:visible']
        for selector in username_selectors:
            try:
                if page.is_visible(selector):
                    page.fill(selector, self.username)
                    break
            except:
                continue
        
        # 尝试多种方式定位密码
        password_selectors = ['input[placeholder*="密码"]', 'input[type="password"]']
        for selector in password_selectors:
            try:
                if page.is_visible(selector):
                    page.fill(selector, self.password)
                    break
            except:
                continue
        
        time.sleep(0.5)
        # 多种方式登录
        login_selectors = ['button:has-text("登录")', 'button:has-text("Login")', 'button[type="submit"]', '.el-button--primary']
        clicked = False
        for selector in login_selectors:
            try:
                if page.is_visible(selector):
                    page.click(selector)
                    clicked = True
                    break
            except:
                continue
        if not clicked:
            page.keyboard.press('Enter') # 回车兜底
        
        page.wait_for_load_state("networkidle", timeout=15000)
        time.sleep(3)
        logger.info("登录成功")

    def select_environment(self, page):
        time.sleep(2)
        selectors = ['select:has-text("选择环境")', '.el-select .el-input__inner', 'div[role="combobox"]']
        clicked = False
        for selector in selectors:
            try:
                if page.is_visible(selector):
                    page.click(selector)
                    time.sleep(1)
                    clicked = True
                    break
            except:
                continue
        
        if clicked:
            try:
                page.click(f'text={self.environment}')
                time.sleep(1)
                logger.info(f"已选择环境：{self.environment}")
            except:
                logger.warning("未找到目标环境选项")
            
            try:
                page.click('button:has-text("确认"), button:has-text("确定")')
                time.sleep(1)
            except:
                pass
        page.wait_for_load_state("networkidle")

    def scrape_data1(self, page):
        # 进入代理限制 -> 设备
        page.goto("https://super.akubela.com/#/distributor/device")
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        
        # 1. 点击授权状态下拉框
        try:
            status_selectors = ['text=授权状态', 'div:has-text("授权状态")', 'span:has-text("授权状态")']
            for selector in status_selectors:
                if page.is_visible(selector):
                    page.click(selector)
                    time.sleep(1)
                    break
            # 2. 选择"已限制"
            for option in ['已限制', 'restricted', 'Restricted']:
                if page.is_visible(f'text={option}'):
                    page.click(f'text={option}')
                    time.sleep(1)
                    break
        except Exception as e:
            logger.warning(f"选择授权状态失败: {e}")

        # 3. 点击搜索按钮
        search_btns = ['button:has-text("搜索")', 'button:has-text("查询")', '.el-button:has-text("搜索")']
        for btn in search_btns:
            try:
                if page.is_visible(btn):
                    page.click(btn)
                    time.sleep(3)
                    break
            except:
                continue
        
        page.wait_for_load_state("networkidle")
        time.sleep(3)
        self.device_data = self.extract_table_data(page)
        logger.info(f"设备数据(已限制): {len(self.device_data)} 条")

    def scrape_data2(self, page):
        # 进入跨区管控页面
        page.goto("https://super.akubela.com/#/distributor/cross-region")
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        
        # 1. 点击风险等级下拉框
        try:
            risk_selectors = ['text=风险等级', 'div:has-text("风险等级")', 'span:has-text("风险等级")']
            for selector in risk_selectors:
                if page.is_visible(selector):
                    page.click(selector)
                    time.sleep(1)
                    break
            # 2. 选择"高"
            for option in ['高', 'high', 'High']:
                if page.is_visible(f'text={option}'):
                    page.click(f'text={option}')
                    time.sleep(1)
                    break
        except Exception as e:
            logger.warning(f"选择风险等级失败: {e}")

        # 3. 点击搜索按钮
        search_btns = ['button:has-text("搜索")', 'button:has-text("查询")', '.el-button:has-text("搜索")']
        for btn in search_btns:
            try:
                if page.is_visible(btn):
                    page.click(btn)
                    time.sleep(3)
                    break
            except:
                continue
        
        page.wait_for_load_state("networkidle")
        time.sleep(3)
        self.cross_region_data = self.extract_table_data(page)
        logger.info(f"跨区管控数据(高风险): {len(self.cross_region_data)} 条")

    def extract_table_data(self, page):
        """2.py 的核心提取逻辑，极其稳健"""
        data = []
        try:
            time.sleep(3)
            tables = page.locator('table').all()
            if not tables:
                return data
            
            for table in tables:
                # 提取表头
                headers = []
                try:
                    header_cells = table.locator('thead th').all()
                    if header_cells:
                        headers = [cell.text_content().strip() for cell in header_cells]
                except:
                    try:
                        first_row = table.locator('tr').first
                        first_cells = first_row.locator('th, td').all()
                        if first_cells:
                            headers = [cell.text_content().strip() for cell in first_cells]
                    except:
                        pass
                
                # 提取行
                rows = table.locator('tbody tr').all()
                if not rows:
                    all_rows = table.locator('tr').all()
                    if len(all_rows) > 1:
                        rows = all_rows[1:]
                
                for row in rows:
                    cells = row.locator('td').all()
                    row_data = {}
                    for i, cell in enumerate(cells):
                        cell_text = cell.text_content().strip()
                        if i < len(headers):
                            row_data[headers[i]] = cell_text
                        else:
                            row_data[f'列{i+1}'] = cell_text
                    if row_data:
                        data.append(row_data)
                
                if data:
                    break
        except Exception as e:
            logger.error(f"提取数据异常: {e}")
        
        return data

    # ---------------------------------------------------------
    # 以下沿用 1.py 原有的保存和发送邮件逻辑
    # ---------------------------------------------------------

    def save_data(self):
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
                <tr><td>设备数据 (已限制)</td><td>{len(self.device_data)}</td></tr>
                <tr><td>跨区管控数据 (高风险)</td><td>{len(self.cross_region_data)}</td></tr>
            </table>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
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
