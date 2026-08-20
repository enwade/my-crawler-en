# -*- coding: utf-8 -*-
"""
所有数据（含菜单导航）
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

    # ========== 导航函数（来自2.py） ==========
    def navigate_to_agent_restriction(self, page):
        """导航到代理限制菜单"""
        logger.info("  导航到代理限制...")
        try:
            page.click('text=代理限制')
            time.sleep(2)
            logger.info("  已点击'代理限制'")
        except:
            try:
                page.click('a:has-text("代理限制")')
                time.sleep(2)
                logger.info("  已点击'代理限制'链接")
            except:
                logger.warning("  无法找到'代理限制'菜单，尝试直接跳转")
                page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                return
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(2)

    def navigate_to_device(self, page):
        """导航到设备页面（通过菜单）"""
        logger.info("  导航到设备页面...")
        # 先进入代理限制
        self.navigate_to_agent_restriction(page)
        # 点击设备子菜单
        try:
            page.click('text=设备')
            time.sleep(2)
            logger.info("  已点击'设备'子菜单")
        except:
            try:
                page.click('a:has-text("设备")')
                time.sleep(2)
                logger.info("  已点击'设备'链接")
            except:
                logger.warning("  无法找到'设备'子菜单，直接跳转")
                page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)

    def navigate_to_cross_region(self, page):
        """导航到跨区管控页面"""
        logger.info("  导航到跨区管控...")
        # 先进入代理限制
        self.navigate_to_agent_restriction(page)
        # 点击跨区管控子菜单
        try:
            page.click('text=跨区管控')
            time.sleep(2)
            logger.info("  已点击'跨区管控'子菜单")
        except:
            try:
                page.click('a:has-text("跨区管控")')
                time.sleep(2)
                logger.info("  已点击'跨区管控'链接")
            except:
                logger.warning("  无法找到'跨区管控'子菜单，直接跳转")
                page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)

    # ========== 主流程 ==========
    def run(self):
        logger.info("=" * 70)
        logger.info("开始爬取 Akubela 数据 (菜单导航版)")
        logger.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 70)

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True, slow_mo=1200)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                page.set_default_timeout(120000)

                # ========== 1. 登录 ==========
                logger.info("【步骤1】登录")
                try:
                    page.goto("https://super.akubela.com/#/login", timeout=60000)
                    time.sleep(5)
                    page.wait_for_selector('input[type="text"]:visible', timeout=30000)
                    time.sleep(2)
                    page.fill('input[type="text"]:visible', self.username)
                    page.fill('input[type="password"]', self.password)
                    page.click('.el-button--primary')
                    time.sleep(10)
                    
                    login_indicators = ['text=Super', 'text=super', '.el-menu', '.sidebar']
                    found = self.wait_for_any(page, login_indicators, timeout=60000)
                    if found:
                        logger.info(f"  ✅ 登录成功")
                        self.login_success = True
                    else:
                        logger.warning("  ⚠️ 登录状态未确认")
                    
                    page.screenshot(path='data/01_after_login.png')
                    if "login" in page.url and not self.login_success:
                        self.error_message = "登录失败"
                        browser.close()
                        self.send_email(success=False)
                        return False
                except Exception as e:
                    logger.error(f"  ❌ 登录异常: {e}")
                    self.error_message = f"登录异常: {e}"
                    browser.close()
                    self.send_email(success=False)
                    return False

                # ========== 2. 选择环境 ==========
                logger.info("【步骤2】选择环境")
                try:
                    time.sleep(10)
                    page.screenshot(path='data/02_before_env.png')
                    
                    env_selectors = [
                        'div[role="combobox"]',
                        '.el-select',
                        'select',
                        '//div[contains(@class, "el-select")]',
                    ]
                    
                    clicked = False
                    for sel in env_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                logger.info(f"  点击选择器: {sel}")
                                page.click(sel)
                                time.sleep(5)
                                clicked = True
                                break
                        except:
                            continue
                    
                    if clicked:
                        option_selector = f'text={self.environment}'
                        if page.wait_for_selector(option_selector, timeout=30000):
                            page.click(option_selector)
                            time.sleep(3)
                            logger.info(f"  ✅ 已选择环境: {self.environment}")
                            self.env_selected = True
                        else:
                            try:
                                page.locator('.el-select-dropdown__item').first.click()
                                logger.info("  已选择第一个选项")
                                self.env_selected = True
                            except:
                                pass
                    
                    page.screenshot(path='data/02_after_env.png')
                except Exception as e:
                    logger.warning(f"  ⚠️ 选择环境异常: {e}")

                # ========== 3. 导航到设备页面 ==========
                logger.info("【步骤3】导航到设备页面")
                try:
                    self.navigate_to_device(page)
                    page.screenshot(path='data/03_device_page.png')
                except Exception as e:
                    logger.warning(f"  导航到设备页面异常: {e}")

                # ========== 4. 抓取设备数据 ==========
                logger.info("【步骤4】抓取设备数据")
                try:
                    self.device_data = self.get_table_data_enhanced(page)
                    logger.info(f"  ✅ 设备数据: {len(self.device_data)} 条")
                    self.device_success = True
                except Exception as e:
                    logger.error(f"  ❌ 抓取设备数据失败: {e}")
                    self.error_message += "设备数据失败;"

                # ========== 5. 导航到跨区管控 ==========
                logger.info("【步骤5】导航到跨区管控")
                try:
                    self.navigate_to_cross_region(page)
                    page.screenshot(path='data/04_cross_region_page.png')
                except Exception as e:
                    logger.warning(f"  导航到跨区管控异常: {e}，尝试直接跳转")
                    page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
                    time.sleep(10)

                # ========== 6. 抓取跨区管控数据 ==========
                logger.info("【步骤6】抓取跨区管控数据")
                try:
                    self.cross_region_data = self.get_table_data_enhanced(page)
                    logger.info(f"  ✅ 跨区管控数据: {len(self.cross_region_data)} 条")
                    self.cross_region_success = True
                except Exception as e:
                    logger.error(f"  ❌ 抓取跨区管控数据失败: {e}")
                    self.error_message += "跨区管控失败;"

                browser.close()

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

    # ========== 数据提取（增强版） ==========
    def get_table_data_enhanced(self, page):
        """增强版数据提取（来自2.py的优化）"""
        data = []
        try:
            # 等待表格出现
            logger.info("    等待表格加载...")
            for attempt in range(3):
                try:
                    table_selectors = ['table', '.el-table', '.el-table__body']
                    table_found = False
                    for sel in table_selectors:
                        if page.locator(sel).count() > 0:
                            logger.info(f"    ✅ 找到表格: {sel}")
                            table_found = True
                            break
                    if table_found:
                        break
                    time.sleep(3)
                except:
                    time.sleep(3)
            
            time.sleep(5)
            
            # 提取行
            rows = []
            methods = [
                ('table', 'table tbody tr'),
                ('el-table', '.el-table tbody tr'),
                ('xpath', '//table//tbody//tr'),
            ]
            
            for method_name, selector in methods:
                try:
                    rows = page.locator(selector).all()
                    if rows and len(rows) > 0:
                        logger.info(f"    方法 {method_name} 找到 {len(rows)} 行")
                        break
                except:
                    continue
            
            if not rows:
                if page.locator('text=暂无数据').count() > 0:
                    logger.warning("    页面显示：暂无数据")
                    return data
                if page.locator('.el-table__empty-text').count() > 0:
                    logger.warning("    页面显示：空状态")
                    return data
                logger.warning("    未找到任何数据行")
                return data
            
            # 提取每行数据
            for row in rows:
                try:
                    cells = row.locator('td').all()
                    if cells:
                        row_data = {}
                        for i, cell in enumerate(cells):
                            text = cell.text_content().strip()
                            if text:
                                row_data[f'列{i+1}'] = text
                        if row_data:
                            data.append(row_data)
                except:
                    continue
            
            logger.info(f"    ✅ 成功提取 {len(data)} 行")
            
        except Exception as e:
            logger.error(f"    ❌ 提取失败: {e}")
        
        return data

    # ========== 保存和发送邮件 ==========
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

            for f in [self.device_excel, self.cross_region_excel, self.device_json, self.cross_region_json]:
                if Path(f).exists():
                    with open(f, 'rb') as fp:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(fp.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{Path(f).name}"')
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
