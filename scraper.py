# -*- coding: utf-8 -*-
"""
数据爬虫 - 抓取所有数据（优化版，增加等待时间）
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
from playwright.sync_api import sync_playwright, TimeoutError

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
    
    def wait_for_page_load(self, page, timeout=60000):
        """等待页面完全加载，增加超时时间"""
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except:
            logger.warning("网络空闲等待超时，继续执行")
        time.sleep(5)  # 额外等待
    
    def safe_click(self, page, selector, timeout=30000, retry=3):
        """安全点击，带重试"""
        for i in range(retry):
            try:
                page.wait_for_selector(selector, timeout=timeout)
                page.click(selector)
                return True
            except:
                logger.warning(f"点击失败 {selector}，第 {i+1} 次重试...")
                time.sleep(3)
        return False
    
    def run(self):
        logger.info("=" * 50)
        logger.info("开始爬取数据...")
        logger.info("注意：由于服务器在美国，访问国内服务可能较慢，请耐心等待")
        
        try:
            with sync_playwright() as p:
                # 使用更慢的操作速度
                browser = p.chromium.launch(headless=True, slow_mo=1000)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                
                # 设置更长的超时时间
                page.set_default_timeout(60000)
                
                # 1. 登录 - 增加等待
                logger.info("正在登录...")
                page.goto("https://super.akubela.com/#/login", timeout=60000)
                self.wait_for_page_load(page)
                
                # 等待输入框出现
                logger.info("等待输入框加载...")
                page.wait_for_selector('input[type="text"]:visible', timeout=60000)
                time.sleep(3)
                
                page.fill('input[type="text"]:visible', self.username)
                page.fill('input[type="password"]', self.password)
                
                # 点击登录
                logger.info("点击登录按钮...")
                page.click('.el-button--primary')
                
                # 等待登录完成 - 加长等待
                logger.info("等待登录完成...")
                time.sleep(10)  # 给服务器足够时间响应
                
                # 截图检查登录状态
                page.screenshot(path='data/login_result.png')
                logger.info("登录截图已保存")
                
                if "login" in page.url:
                    logger.error("登录失败，仍然在登录页面")
                    browser.close()
                    return False
                logger.info(f"登录成功，当前URL: {page.url}")
                
                # 2. 选择环境 - 增加等待
                logger.info("等待环境选择器加载...")
                time.sleep(8)  # 额外等待下拉框加载
                
                try:
                    # 先截图看页面状态
                    page.screenshot(path='data/env_page.png')
                    
                    # 等待环境选择器出现
                    env_selector = 'div[role="combobox"]'
                    if page.wait_for_selector(env_selector, timeout=30000):
                        logger.info("环境选择器已加载")
                        page.click(env_selector)
                        time.sleep(3)  # 等待下拉选项加载
                        
                        # 点击目标环境
                        page.click(f'text={self.environment}')
                        logger.info(f"已选择环境: {self.environment}")
                        time.sleep(3)
                    else:
                        logger.warning("环境选择器未加载，跳过")
                except Exception as e:
                    logger.warning(f"选择环境失败: {e}")
                    page.screenshot(path='data/env_error.png')
                
                # 3. 抓取设备数据 - 大幅增加等待
                logger.info("跳转到设备页面...")
                page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                logger.info("等待设备页面加载...")
                time.sleep(15)  # 给更多时间加载数据
                
                # 截图
                page.screenshot(path='data/device_page.png')
                logger.info("设备页面截图已保存")
                
                self.device_data = self.get_table_data(page)
                logger.info(f"设备数据: {len(self.device_data)} 条")
                
                # 4. 抓取跨区管控数据
                logger.info("跳转到跨区管控页面...")
                page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
                logger.info("等待跨区管控页面加载...")
                time.sleep(15)
                
                # 截图
                page.screenshot(path='data/cross_region_page.png')
                logger.info("跨区管控页面截图已保存")
                
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
        """提取表格数据 - 增强版"""
        data = []
        try:
            # 等待表格加载
            logger.info("等待表格数据加载...")
            time.sleep(5)
            
            rows = []
            
            # 方法1：通过 table 标签
            try:
                table = page.locator('table').first
                if table.count() > 0:
                    rows = table.locator('tbody tr').all()
                    logger.info(f"方法1 - 通过table找到 {len(rows)} 行")
            except:
                pass
            
            # 方法2：通过 el-table 类
            if not rows:
                try:
                    table = page.locator('.el-table')
                    if table.count() > 0:
                        rows = table.locator('tbody tr').all()
                        logger.info(f"方法2 - 通过el-table找到 {len(rows)} 行")
                except:
                    pass
            
            # 方法3：通过 xpath
            if not rows:
                try:
                    rows = page.locator('//table//tbody//tr').all()
                    logger.info(f"方法3 - 通过xpath找到 {len(rows)} 行")
                except:
                    pass
            
            # 检查是否有"暂无数据"
            if not rows:
                try:
                    no_data = page.locator('text=暂无数据')
                    if no_data.count() > 0:
                        logger.warning("页面显示：暂无数据")
                    else:
                        no_data = page.locator('.el-table__empty-text')
                        if no_data.count() > 0:
                            logger.warning("页面显示：暂无数据（el-table空状态）")
                except:
                    pass
            
            for row in rows:
                try:
                    cells = row.locator('td').all()
                    if cells:
                        row_data = {}
                        for i, cell in enumerate(cells):
                            text = cell.text_content().strip()
                            row_data[f'列{i+1}'] = text
                        data.append(row_data)
                except:
                    continue
                
        except Exception as e:
            logger.error(f"提取表格数据时出错: {e}")
        
        return data
    
    def save_data(self):
        """保存数据"""
        logger.info("保存数据...")
        
        # 保存 JSON
        try:
            with open(self.device_json, 'w', encoding='utf-8') as f:
                json.dump(self.device_data, f, ensure_ascii=False, indent=2)
            logger.info(f"设备JSON已保存: {self.device_json}")
        except Exception as e:
            logger.error(f"保存设备JSON失败: {e}")
        
        try:
            with open(self.cross_region_json, 'w', encoding='utf-8') as f:
                json.dump(self.cross_region_data, f, ensure_ascii=False, indent=2)
            logger.info(f"跨区JSON已保存: {self.cross_region_json}")
        except Exception as e:
            logger.error(f"保存跨区JSON失败: {e}")
        
        # 保存 Excel
        try:
            if self.device_data:
                df = pd.DataFrame(self.device_data)
                df.to_excel(self.device_excel, index=False, engine='openpyxl')
                logger.info(f"设备Excel已保存: {self.device_excel}")
            else:
                logger.warning("设备数据为空，不生成Excel")
        except Exception as e:
            logger.error(f"保存设备Excel失败: {e}")
        
        try:
            if self.cross_region_data:
                df = pd.DataFrame(self.cross_region_data)
                df.to_excel(self.cross_region_excel, index=False, engine='openpyxl')
                logger.info(f"跨区Excel已保存: {self.cross_region_excel}")
            else:
                logger.warning("跨区数据为空，不生成Excel")
        except Exception as e:
            logger.error(f"保存跨区Excel失败: {e}")
    
    def send_email(self):
        """发送邮件"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email
            msg['Subject'] = f'Akubela数据报告 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            
            body = f"""
            <h2>Akubela 数据报告</h2>
            <p>爬取时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
            <p>环境: {self.environment}</p>
            <table border="1">
                <tr><th>数据源</th><th>记录数</th><th>状态</th></tr>
                <tr>
                    <td>设备数据</td>
                    <td>{len(self.device_data)}</td>
                    <td>{'✅ 有数据' if len(self.device_data) > 0 else '⚠️ 无数据'}</td>
                </tr>
                <tr>
                    <td>跨区管控数据</td>
                    <td>{len(self.cross_region_data)}</td>
                    <td>{'✅ 有数据' if len(self.cross_region_data) > 0 else '⚠️ 无数据'}</td>
                </tr>
            </table>
            <p>附件为数据文件（JSON + Excel）</p>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # 添加附件
            files_to_attach = [
                self.device_excel if Path(self.device_excel).exists() else None,
                self.cross_region_excel if Path(self.cross_region_excel).exists() else None,
                self.device_json if Path(self.device_json).exists() else None,
                self.cross_region_json if Path(self.cross_region_json).exists() else None,
            ]
            
            attached_count = 0
            for file_path in files_to_attach:
                if file_path:
                    try:
                        with open(file_path, 'rb') as f:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="{Path(file_path).name}"')
                            msg.attach(part)
                            attached_count += 1
                    except Exception as e:
                        logger.error(f"添加附件失败 {file_path}: {e}")
            
            logger.info(f"共添加 {attached_count} 个附件")
            
            # 发送
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
