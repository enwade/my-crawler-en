# -*- coding: utf-8 -*-
"""
抓取所有数据
"""

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
        
        # 标记执行状态
        self.login_success = False
        self.device_success = False
        self.cross_region_success = False
    
    def run(self):
        logger.info("=" * 60)
        logger.info("开始爬取 Akubela 数据")
        logger.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 60)
        
        try:
            with sync_playwright() as p:
                # 使用慢速模式，给服务器足够响应时间
                browser = p.chromium.launch(headless=True, slow_mo=1000)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                page.set_default_timeout(60000)
                
                # ============ 第1步：登录 ============
                logger.info("【第1步】开始登录...")
                try:
                    logger.info("  正在加载登录页面...")
                    page.goto("https://super.akubela.com/#/login", timeout=60000)
                    page.wait_for_load_state("networkidle", timeout=60000)
                    time.sleep(5)
                    
                    logger.info("  等待输入框加载...")
                    page.wait_for_selector('input[type="text"]:visible', timeout=30000)
                    time.sleep(2)
                    
                    logger.info("  填写用户名...")
                    page.fill('input[type="text"]:visible', self.username)
                    
                    logger.info("  填写密码...")
                    page.fill('input[type="password"]', self.password)
                    
                    logger.info("  点击登录按钮...")
                    page.click('.el-button--primary')
                    
                    logger.info("  等待登录完成（等待10秒）...")
                    time.sleep(10)
                    
                    # 截图
                    page.screenshot(path='data/01_login_result.png')
                    logger.info("  登录截图已保存: data/01_login_result.png")
                    
                    # 检查是否登录成功
                    current_url = page.url
                    logger.info(f"  当前URL: {current_url}")
                    
                    if "login" in current_url:
                        logger.error("  ❌ 登录失败：仍然在登录页面")
                        browser.close()
                        # 即使登录失败也发送邮件通知
                        self.send_email(success=False, error="登录失败")
                        return False
                    
                    self.login_success = True
                    logger.info("  ✅ 登录成功！")
                    
                except Exception as e:
                    logger.error(f"  ❌ 登录过程出错: {e}")
                    browser.close()
                    self.send_email(success=False, error=f"登录出错: {e}")
                    return False
                
                # ============ 第2步：选择环境 ============
                logger.info("【第2步】选择环境...")
                try:
                    logger.info("  等待页面完全加载...")
                    time.sleep(8)
                    
                    # 截图
                    page.screenshot(path='data/02_before_env.png')
                    logger.info("  选择环境前截图: data/02_before_env.png")
                    
                    # 等待环境选择器
                    env_selector = 'div[role="combobox"]'
                    logger.info(f"  等待选择器出现: {env_selector}")
                    
                    if page.wait_for_selector(env_selector, timeout=30000):
                        logger.info("  ✅ 环境选择器已出现")
                        time.sleep(2)
                        
                        logger.info("  点击环境选择器...")
                        page.click(env_selector)
                        time.sleep(5)  # 等待下拉选项加载
                        
                        logger.info(f"  查找选项: {self.environment}")
                        if page.wait_for_selector(f'text={self.environment}', timeout=10000):
                            logger.info(f"  ✅ 找到选项: {self.environment}")
                            page.click(f'text={self.environment}')
                            time.sleep(3)
                            logger.info(f"  ✅ 已选择环境: {self.environment}")
                        else:
                            logger.warning(f"  ⚠️ 未找到选项: {self.environment}")
                    else:
                        logger.warning("  ⚠️ 环境选择器未出现，跳过")
                    
                except Exception as e:
                    logger.warning(f"  ⚠️ 选择环境出错: {e}，继续执行")
                    page.screenshot(path='data/02_env_error.png')
                
                # ============ 第3步：抓取设备数据 ============
                logger.info("【第3步】抓取设备数据...")
                try:
                    logger.info("  跳转到设备页面...")
                    page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                    logger.info("  等待设备页面加载（等待15秒）...")
                    time.sleep(15)
                    
                    page.screenshot(path='data/03_device_page.png')
                    logger.info("  设备页面截图: data/03_device_page.png")
                    
                    logger.info("  开始提取设备数据...")
                    self.device_data = self.get_table_data(page)
                    logger.info(f"  ✅ 设备数据: {len(self.device_data)} 条")
                    self.device_success = True
                    
                except Exception as e:
                    logger.error(f"  ❌ 抓取设备数据出错: {e}")
                    page.screenshot(path='data/03_device_error.png')
                
                # ============ 第4步：抓取跨区管控数据 ============
                logger.info("【第4步】抓取跨区管控数据...")
                try:
                    logger.info("  跳转到跨区管控页面...")
                    page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
                    logger.info("  等待跨区管控页面加载（等待15秒）...")
                    time.sleep(15)
                    
                    page.screenshot(path='data/04_cross_region_page.png')
                    logger.info("  跨区管控页面截图: data/04_cross_region_page.png")
                    
                    logger.info("  开始提取跨区管控数据...")
                    self.cross_region_data = self.get_table_data(page)
                    logger.info(f"  ✅ 跨区管控数据: {len(self.cross_region_data)} 条")
                    self.cross_region_success = True
                    
                except Exception as e:
                    logger.error(f"  ❌ 抓取跨区管控数据出错: {e}")
                    page.screenshot(path='data/04_cross_region_error.png')
                
                browser.close()
            
            # ============ 第5步：保存数据 ============
            logger.info("【第5步】保存数据...")
            self.save_data()
            
            # ============ 第6步：发送邮件 ============
            logger.info("【第6步】发送邮件...")
            self.send_email(success=True)
            
            logger.info("=" * 60)
            logger.info("✅ 爬取完成！")
            logger.info("=" * 60)
            return True
            
        except Exception as e:
            logger.error(f"❌ 程序运行出错: {e}")
            try:
                self.send_email(success=False, error=str(e))
            except:
                logger.error("发送失败通知邮件也出错了")
            return False
    
    def get_table_data(self, page):
        """提取表格数据"""
        data = []
        try:
            # 额外等待
            time.sleep(5)
            
            rows = []
            
            # 方法1：通过 table 标签
            try:
                table = page.locator('table').first
                if table.count() > 0:
                    rows = table.locator('tbody tr').all()
                    logger.info(f"    方法1(table)找到 {len(rows)} 行")
            except Exception as e:
                logger.info(f"    方法1失败: {e}")
            
            # 方法2：通过 el-table
            if not rows:
                try:
                    table = page.locator('.el-table')
                    if table.count() > 0:
                        rows = table.locator('tbody tr').all()
                        logger.info(f"    方法2(el-table)找到 {len(rows)} 行")
                except Exception as e:
                    logger.info(f"    方法2失败: {e}")
            
            # 方法3：通过 xpath
            if not rows:
                try:
                    rows = page.locator('//table//tbody//tr').all()
                    logger.info(f"    方法3(xpath)找到 {len(rows)} 行")
                except Exception as e:
                    logger.info(f"    方法3失败: {e}")
            
            # 检查是否无数据
            if not rows:
                try:
                    no_data = page.locator('text=暂无数据')
                    if no_data.count() > 0:
                        logger.warning("    页面显示：暂无数据")
                except:
                    pass
                return data
            
            # 提取数据
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
            
            logger.info(f"    成功提取 {len(data)} 行数据")
                
        except Exception as e:
            logger.error(f"    提取表格数据出错: {e}")
        
        return data
    
    def save_data(self):
        """保存数据"""
        try:
            # 保存 JSON
            with open(self.device_json, 'w', encoding='utf-8') as f:
                json.dump(self.device_data, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 设备JSON: {self.device_json}")
        except Exception as e:
            logger.error(f"  ❌ 保存设备JSON失败: {e}")
        
        try:
            with open(self.cross_region_json, 'w', encoding='utf-8') as f:
                json.dump(self.cross_region_data, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 跨区JSON: {self.cross_region_json}")
        except Exception as e:
            logger.error(f"  ❌ 保存跨区JSON失败: {e}")
        
        # 保存 Excel
        try:
            if self.device_data:
                df = pd.DataFrame(self.device_data)
                df.to_excel(self.device_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 设备Excel: {self.device_excel}")
            else:
                logger.warning("  ⚠️ 设备数据为空，不生成Excel")
        except Exception as e:
            logger.error(f"  ❌ 保存设备Excel失败: {e}")
        
        try:
            if self.cross_region_data:
                df = pd.DataFrame(self.cross_region_data)
                df.to_excel(self.cross_region_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 跨区Excel: {self.cross_region_excel}")
            else:
                logger.warning("  ⚠️ 跨区数据为空，不生成Excel")
        except Exception as e:
            logger.error(f"  ❌ 保存跨区Excel失败: {e}")
    
    def send_email(self, success=True, error=""):
        """发送邮件"""
        logger.info("正在发送邮件...")
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email
            
            if success:
                status = "✅ 成功"
                subject = f'Akubela数据报告 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            else:
                status = "❌ 失败"
                subject = f'Akubela爬虫异常 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            
            msg['Subject'] = subject
            
            # 构建邮件正文
            body = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    h2 {{ color: #333; }}
                    .success {{ color: green; }}
                    .fail {{ color: red; }}
                    table {{ border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                </style>
            </head>
            <body>
                <h2>Akubela 数据爬虫报告</h2>
                
                <p><strong>状态：</strong><span class="{'success' if success else 'fail'}">{status}</span></p>
                <p><strong>时间：</strong>{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><strong>环境：</strong>{self.environment}</p>
                
                <h3>执行结果</h3>
                <table>
                    <tr>
                        <th>步骤</th>
                        <th>状态</th>
                        <th>记录数</th>
                    </tr>
                    <tr>
                        <td>登录</td>
                        <td>{'✅' if self.login_success else '❌'}</td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td>设备数据</td>
                        <td>{'✅' if self.device_success else '❌'}</td>
                        <td>{len(self.device_data)}</td>
                    </tr>
                    <tr>
                        <td>跨区管控数据</td>
                        <td>{'✅' if self.cross_region_success else '❌'}</td>
                        <td>{len(self.cross_region_data)}</td>
                    </tr>
                </table>
                
                {f'<p><strong>错误信息：</strong>{error}</p>' if error else ''}
                
                <p>附件为数据文件。</p>
                <hr>
                <p style="color: #666; font-size: 12px;">此邮件由自动化系统发送，请勿回复。</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # 添加附件
            files_to_attach = []
            if Path(self.device_excel).exists():
                files_to_attach.append(self.device_excel)
            if Path(self.cross_region_excel).exists():
                files_to_attach.append(self.cross_region_excel)
            if Path(self.device_json).exists():
                files_to_attach.append(self.device_json)
            if Path(self.cross_region_json).exists():
                files_to_attach.append(self.cross_region_json)
            
            for file_path in files_to_attach:
                try:
                    with open(file_path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{Path(file_path).name}"')
                        msg.attach(part)
                        logger.info(f"  添加附件: {Path(file_path).name}")
                except Exception as e:
                    logger.error(f"  添加附件失败 {file_path}: {e}")
            
            # 发送
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            server.login(self.sender_email, self.sender_password)
            server.sendmail(self.sender_email, self.receiver_email, msg.as_string())
            server.quit()
            
            logger.info(f"✅ 邮件发送成功到: {self.receiver_email}")
            
        except Exception as e:
            logger.error(f"❌ 邮件发送失败: {e}")


if __name__ == "__main__":
    scraper = AkubelaScraper()
    scraper.run()
