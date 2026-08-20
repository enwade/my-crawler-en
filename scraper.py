# -*- coding: utf-8 -*-
"""
Akubela 数据爬虫 - 强化环境选择版
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
        self.username = os.getenv('AKUBELA_USERNAME', 'super')
        self.password = os.getenv('AKUBELA_PASSWORD', 'Akubela@super2024&')
        self.environment = os.getenv('AKUBELA_ENV', 'prod-cn-hz')
        
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.exmail.qq.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '465'))
        self.sender_email = os.getenv('SENDER_EMAIL', 'enwade.chen@akuvox.com')
        self.sender_password = os.getenv('SENDER_PASSWORD', 'cyw110120')
        self.receiver_email = os.getenv('RECEIVER_EMAIL', 'enwade.chen@akuvox.com')
        
        self.device_data = []
        self.cross_region_data = []
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        Path('./data').mkdir(exist_ok=True)
        
        self.device_json = f'data/devices_{self.timestamp}.json'
        self.cross_region_json = f'data/cross_region_{self.timestamp}.json'
        self.device_excel = f'data/devices_{self.timestamp}.xlsx'
        self.cross_region_excel = f'data/cross_region_{self.timestamp}.xlsx'
        
        self.login_success = False
        self.env_selected = False
        self.device_success = False
        self.cross_region_success = False
        self.error_message = ""

    def wait_for_any(self, page, selectors, timeout=60000):
        """等待任意一个选择器出现"""
        start = time.time()
        while time.time() - start < timeout / 1000:
            for sel in selectors:
                try:
                    if page.locator(sel).count() > 0:
                        logger.info(f"  ✅ 找到元素: {sel}")
                        return sel
                except:
                    pass
            time.sleep(1)
        return None

    def run(self):
        logger.info("=" * 70)
        logger.info("开始爬取 Akubela 数据 (强化环境选择版)")
        logger.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 70)

        try:
            with sync_playwright() as p:
                # 超慢速，超长超时
                browser = p.chromium.launch(headless=True, slow_mo=1500)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                page.set_default_timeout(120000)

                # ========== 1. 登录 ==========
                logger.info("【步骤1】登录")
                try:
                    logger.info("  加载登录页...")
                    page.goto("https://super.akubela.com/#/login", timeout=60000)
                    time.sleep(5)

                    logger.info("  等待输入框...")
                    page.wait_for_selector('input[type="text"]:visible', timeout=30000)
                    time.sleep(2)

                    logger.info("  输入用户名...")
                    page.fill('input[type="text"]:visible', self.username)
                    logger.info("  输入密码...")
                    page.fill('input[type="password"]', self.password)

                    logger.info("  点击登录按钮...")
                    page.click('.el-button--primary')
                    logger.info("  等待登录响应...")

                    # 等待登录成功标志
                    login_indicators = [
                        'text=Super',
                        'text=super',
                        '.el-menu',
                        '.sidebar',
                        '.header-user'
                    ]
                    found = self.wait_for_any(page, login_indicators, timeout=60000)
                    if found:
                        logger.info(f"  ✅ 登录成功，检测到: {found}")
                        self.login_success = True
                    else:
                        logger.warning("  ⚠️ 未检测到登录标志，继续")

                    time.sleep(5)
                    page.screenshot(path='data/01_after_login.png')

                    if "login" in page.url and not self.login_success:
                        logger.error("  ❌ 登录失败")
                        self.error_message = "登录失败"
                        browser.close()
                        self.send_email(success=False)
                        return False

                    self.login_success = True
                    logger.info("  ✅ 登录完成")

                except Exception as e:
                    logger.error(f"  ❌ 登录异常: {e}")
                    self.error_message = f"登录异常: {e}"
                    browser.close()
                    self.send_email(success=False)
                    return False

                # ========== 2. 选择环境（核心加强） ==========
                logger.info("【步骤2】选择环境")
                try:
                    # 等待页面完全加载
                    logger.info("  等待页面稳定（10秒）...")
                    time.sleep(10)

                    page.screenshot(path='data/02_before_env.png')
                    logger.info("  截图保存: data/02_before_env.png")

                    # 多种方式定位环境选择器
                    env_selectors = [
                        'div[role="combobox"]',
                        '.el-select',
                        'select',
                        '[placeholder*="环境"]',
                        '[placeholder*="Environment"]',
                        '//div[contains(@class, "el-select")]',
                        '//*[text()="环境"]/following::div[1]',  # 通过文本"环境"定位
                    ]

                    # 先尝试点击可能包含"环境"标签的容器
                    try:
                        env_label = page.locator('text=环境')
                        if env_label.count() > 0:
                            logger.info("  找到文本'环境'，点击其后面的选择器")
                            env_label.click()
                            time.sleep(2)
                    except:
                        pass

                    # 遍历选择器
                    clicked = False
                    for sel in env_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                logger.info(f"  尝试点击选择器: {sel}")
                                page.click(sel)
                                time.sleep(5)
                                clicked = True
                                break
                        except:
                            continue

                    if not clicked:
                        logger.warning("  ⚠️ 未找到环境选择器，尝试强制点击第一个可点击的div")
                        # 尝试点击任何可见的div[role="combobox"]或class包含select的元素
                        try:
                            page.locator('div[role="combobox"]').first.click()
                            time.sleep(5)
                            clicked = True
                        except:
                            pass

                    if not clicked:
                        logger.warning("  ⚠️ 实在找不到环境选择器，跳过")
                    else:
                        # 等待下拉选项出现
                        logger.info("  等待下拉选项出现...")
                        option_selector = f'text={self.environment}'
                        if page.wait_for_selector(option_selector, timeout=30000):
                            logger.info(f"  找到选项: {self.environment}")
                            page.click(option_selector)
                            time.sleep(3)
                            logger.info(f"  ✅ 已选择环境: {self.environment}")
                            self.env_selected = True
                        else:
                            # 尝试点击下拉框的第一个选项
                            logger.warning("  未找到目标选项，尝试点击第一个")
                            try:
                                page.locator('.el-select-dropdown__item').first.click()
                                time.sleep(3)
                                logger.info("  已点击第一个选项")
                                self.env_selected = True
                            except:
                                pass

                    if not self.env_selected:
                        logger.warning("  ⚠️ 环境选择可能失败，但继续执行")

                    page.screenshot(path='data/02_after_env.png')
                    logger.info("  截图保存: data/02_after_env.png")

                except Exception as e:
                    logger.warning(f"  ⚠️ 选择环境异常: {e}")
                    page.screenshot(path='data/02_env_error.png')

                # ========== 3. 直接跳转设备页面 ==========
                logger.info("【步骤3】跳转设备页面")
                try:
                    logger.info("  直接跳转到设备页面...")
                    page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                    time.sleep(10)

                    # 检查表格是否出现
                    table_selector = 'table, .el-table'
                    if page.wait_for_selector(table_selector, timeout=60000):
                        logger.info("  ✅ 表格已出现")
                    else:
                        logger.warning("  ⚠️ 表格未出现，可能页面未加载成功")

                    page.screenshot(path='data/03_device_page.png')
                except Exception as e:
                    logger.warning(f"  跳转设备页面异常: {e}")

                # ========== 4. 抓取设备数据 ==========
                logger.info("【步骤4】抓取设备数据")
                try:
                    self.device_data = self.get_table_data(page)
                    logger.info(f"  ✅ 设备数据: {len(self.device_data)} 条")
                    self.device_success = True
                except Exception as e:
                    logger.error(f"  ❌ 抓取设备数据失败: {e}")
                    self.error_message += "设备数据失败;"

                # ========== 5. 抓取跨区管控数据 ==========
                logger.info("【步骤5】抓取跨区管控数据")
                try:
                    logger.info("  跳转到跨区管控页面...")
                    page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
                    time.sleep(10)
                    page.screenshot(path='data/04_cross_region_page.png')
                    self.cross_region_data = self.get_table_data(page)
                    logger.info(f"  ✅ 跨区管控数据: {len(self.cross_region_data)} 条")
                    self.cross_region_success = True
                except Exception as e:
                    logger.error(f"  ❌ 抓取跨区管控数据失败: {e}")
                    self.error_message += "跨区管控失败;"

                browser.close()

            # ========== 6. 保存 ==========
            self.save_data()
            self.send_email(success=True)
            logger.info("=" * 70)
            logger.info("✅ 爬取流程结束")
            return True

        except Exception as e:
            logger.error(f"❌ 主流程异常: {e}")
            self.error_message += f"主流程异常:{e}"
            try:
                self.send_email(success=False)
            except:
                pass
            return False

    def get_table_data(self, page):
        """提取表格数据"""
        data = []
        for attempt in range(3):
            try:
                logger.info(f"    尝试提取表格 (第{attempt+1}次)")
                rows = []
                for method in ['table', '.el-table', '//table']:
                    try:
                        if method.startswith('//'):
                            rows = page.locator(method + '//tbody//tr').all()
                        else:
                            table = page.locator(method).first
                            if table.count() > 0:
                                rows = table.locator('tbody tr').all()
                        if rows:
                            logger.info(f"      方法 {method} 找到 {len(rows)} 行")
                            break
                    except:
                        continue
                if not rows:
                    logger.warning("      无数据行")
                    return data
                for row in rows:
                    cells = row.locator('td').all()
                    if cells:
                        row_data = {}
                        for i, cell in enumerate(cells):
                            row_data[f'列{i+1}'] = cell.text_content().strip()
                        data.append(row_data)
                logger.info(f"      成功提取 {len(data)} 行")
                break
            except Exception as e:
                logger.error(f"      尝试失败: {e}")
                time.sleep(3)
        return data

    def save_data(self):
        """保存数据"""
        try:
            with open(self.device_json, 'w', encoding='utf-8') as f:
                json.dump(self.device_data, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 设备JSON: {self.device_json}")
        except Exception as e:
            logger.error(f"  ❌ 设备JSON保存失败: {e}")

        try:
            with open(self.cross_region_json, 'w', encoding='utf-8') as f:
                json.dump(self.cross_region_data, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 跨区JSON: {self.cross_region_json}")
        except Exception as e:
            logger.error(f"  ❌ 跨区JSON保存失败: {e}")

        try:
            if self.device_data:
                pd.DataFrame(self.device_data).to_excel(self.device_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 设备Excel: {self.device_excel}")
            else:
                logger.warning("  ⚠️ 设备数据为空")
        except Exception as e:
            logger.error(f"  ❌ 设备Excel保存失败: {e}")

        try:
            if self.cross_region_data:
                pd.DataFrame(self.cross_region_data).to_excel(self.cross_region_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 跨区Excel: {self.cross_region_excel}")
            else:
                logger.warning("  ⚠️ 跨区数据为空")
        except Exception as e:
            logger.error(f"  ❌ 跨区Excel保存失败: {e}")

    def send_email(self, success=True):
        """发送邮件"""
        logger.info("正在发送邮件...")
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email

            if success and self.device_success and self.cross_region_success:
                status = "✅ 完全成功"
                subject = f'Akubela数据报告 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            elif success:
                status = "⚠️ 部分成功"
                subject = f'Akubela数据报告（部分成功） - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            else:
                status = "❌ 失败"
                subject = f'Akubela爬虫异常 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'

            msg['Subject'] = subject

            body = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    .success {{ color: green; }}
                    .partial {{ color: orange; }}
                    .fail {{ color: red; }}
                    table {{ border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                </style>
            </head>
            <body>
                <h2>Akubela 数据爬虫报告</h2>
                <p><strong>状态：</strong><span class="{'success' if success and self.device_success and self.cross_region_success else 'partial' if success else 'fail'}">{status}</span></p>
                <p><strong>时间：</strong>{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><strong>环境：</strong>{self.environment}</p>
                <h3>执行结果</h3>
                <table>
                    <tr><th>步骤</th><th>状态</th><th>记录数</th></tr>
                    <tr><td>登录</td><td>{'✅' if self.login_success else '❌'}</td><td>-</td></tr>
                    <tr><td>选择环境</td><td>{'✅' if self.env_selected else '⚠️'}</td><td>-</td></tr>
                    <tr><td>设备数据</td><td>{'✅' if self.device_success else '❌'}</td><td>{len(self.device_data)}</td></tr>
                    <tr><td>跨区管控</td><td>{'✅' if self.cross_region_success else '❌'}</td><td>{len(self.cross_region_data)}</td></tr>
                </table>
                {f'<p><strong>错误信息：</strong>{self.error_message}</p>' if self.error_message else ''}
                <hr><p style="color:#666;font-size:12px;">此邮件由自动化系统发送，请勿回复。</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))

            files_to_attach = []
            for f in [self.device_excel, self.cross_region_excel, self.device_json, self.cross_region_json]:
                if Path(f).exists():
                    files_to_attach.append(f)
            for file_path in files_to_attach:
                with open(file_path, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{Path(file_path).name}"')
                    msg.attach(part)

            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            server.login(self.sender_email, self.sender_password)
            server.sendmail(self.sender_email, self.receiver_email, msg.as_string())
            server.quit()
            logger.info(f"✅ 邮件已发送到: {self.receiver_email}")
        except Exception as e:
            logger.error(f"❌ 邮件发送失败: {e}")


if __name__ == "__main__":
    scraper = AkubelaScraper()
    scraper.run()
