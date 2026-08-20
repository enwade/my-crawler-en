# -*- coding: utf-8 -*-
"""
抓取所有数据（含环境选择）
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

    def wait_for_any_selector(self, page, selectors, timeout=60000, step=2):
        """等待多个选择器中的任意一个出现"""
        start = time.time()
        while time.time() - start < timeout:
            for selector in selectors:
                try:
                    if page.locator(selector).count() > 0:
                        logger.info(f"  找到元素: {selector}")
                        return selector
                except:
                    pass
            time.sleep(step)
        return None

    def run(self):
        logger.info("=" * 70)
        logger.info("开始爬取 Akubela 数据 (含环境选择)")
        logger.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("注意：服务器在美国访问国内服务，可能会非常慢，请耐心等待")
        logger.info("=" * 70)

        try:
            with sync_playwright() as p:
                # 使用慢速模式，超长超时
                browser = p.chromium.launch(headless=True, slow_mo=1200)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                page.set_default_timeout(120000)  # 2分钟

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

                    # 等待登录后的标志（用户信息或菜单）
                    login_indicators = [
                        'text=Super',
                        'text=super',
                        '.el-menu',
                        '.sidebar',
                        '.nav-bar',
                        '.header-user'
                    ]
                    found = self.wait_for_any_selector(page, login_indicators, timeout=60000)
                    if found:
                        logger.info(f"  ✅ 登录成功，检测到: {found}")
                        self.login_success = True
                    else:
                        logger.warning("  ⚠️ 未检测到登录成功标志，继续尝试")

                    time.sleep(5)
                    page.screenshot(path='data/01_after_login.png')
                    logger.info("  截图: data/01_after_login.png")

                    if "login" in page.url:
                        logger.error("  ❌ 登录失败：仍停留在登录页")
                        page.screenshot(path='data/01_login_failed.png')
                        self.error_message = "登录失败"
                        browser.close()
                        self.send_email(success=False)
                        return False

                    logger.info(f"  当前URL: {page.url}")
                    self.login_success = True

                except Exception as e:
                    logger.error(f"  ❌ 登录异常: {e}")
                    page.screenshot(path='data/01_login_exception.png')
                    self.error_message = f"登录异常: {e}"
                    browser.close()
                    self.send_email(success=False)
                    return False

                # ========== 2. 选择环境 ==========
                logger.info("【步骤2】选择环境")
                try:
                    # 等待一段时间让页面完全加载
                    logger.info("  等待页面完全加载（10秒）...")
                    time.sleep(10)

                    # 截图看当前状态
                    page.screenshot(path='data/02_before_env.png')
                    logger.info("  截图: data/02_before_env.png")

                    # 尝试多种方式定位环境选择器
                    env_selectors = [
                        'div[role="combobox"]',
                        '.el-select .el-input__inner',
                        'select',
                        '[placeholder*="环境"]',
                        '[placeholder*="Environment"]'
                    ]
                    env_found = None
                    for sel in env_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                env_found = sel
                                break
                        except:
                            continue

                    if env_found:
                        logger.info(f"  找到环境选择器: {env_found}")
                        # 点击展开下拉
                        logger.info("  点击环境选择器展开下拉...")
                        page.click(env_found)
                        time.sleep(5)  # 等待下拉选项加载

                        # 等待目标选项出现
                        option_selector = f'text={self.environment}'
                        logger.info(f"  等待选项: {self.environment}")
                        if page.wait_for_selector(option_selector, timeout=30000):
                            logger.info("  选项已出现，点击选择...")
                            page.click(option_selector)
                            time.sleep(3)
                            logger.info(f"  ✅ 已选择环境: {self.environment}")
                            self.env_selected = True
                        else:
                            logger.warning(f"  ⚠️ 未找到选项: {self.environment}")
                            # 尝试点击下拉框中的第一个选项
                            try:
                                page.locator('.el-select-dropdown__item').first.click()
                                logger.info("  已选择第一个选项")
                                self.env_selected = True
                            except:
                                pass
                    else:
                        logger.warning("  ⚠️ 未找到环境选择器，可能已默认环境或无需选择")

                    if not self.env_selected:
                        logger.warning("  ⚠️ 环境选择可能未完成，继续执行")

                    page.screenshot(path='data/02_after_env.png')
                    logger.info("  截图: data/02_after_env.png")

                except Exception as e:
                    logger.warning(f"  ⚠️ 选择环境异常: {e}，继续执行")
                    page.screenshot(path='data/02_env_error.png')

                # ========== 3. 进入设备页面 ==========
                logger.info("【步骤3】进入设备页面")
                try:
                    # 方法A: 直接跳转
                    logger.info("  尝试直接跳转到设备页面...")
                    page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                    time.sleep(10)

                    # 检查是否加载
                    device_indicators = [
                        '.el-table',
                        'table',
                        'text=设备列表',
                        'text=设备管理'
                    ]
                    found = self.wait_for_any_selector(page, device_indicators, timeout=30000)
                    if found:
                        logger.info(f"  ✅ 设备页面已加载，检测到: {found}")
                    else:
                        logger.warning("  ⚠️ 设备页面可能未加载，尝试点击菜单")

                    page.screenshot(path='data/03_device_page_attempt.png')

                except Exception as e:
                    logger.warning(f"  跳转设备页面异常: {e}，尝试点击菜单")

                # ========== 4. 如果直接跳转不成功，点击菜单 ==========
                logger.info("【步骤4】尝试点击侧边菜单")
                try:
                    menu_selectors = [
                        'text=设备管理',
                        'text=设备',
                        'a:has-text("设备")',
                        '.el-menu-item:has-text("设备")'
                    ]
                    clicked = False
                    for sel in menu_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                logger.info(f"  点击菜单: {sel}")
                                page.click(sel)
                                time.sleep(8)
                                clicked = True
                                break
                        except:
                            continue
                    if clicked:
                        logger.info("  ✅ 已点击设备菜单")
                    else:
                        logger.info("  ⚠️ 未找到设备菜单，继续")

                    # 再次尝试跳转
                    if not clicked:
                        logger.info("  再次尝试直接跳转...")
                        page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
                        time.sleep(10)

                    page.screenshot(path='data/04_after_menu_click.png')

                except Exception as e:
                    logger.warning(f"  点击菜单异常: {e}")

                # ========== 5. 抓取设备数据 ==========
                logger.info("【步骤5】抓取设备数据")
                try:
                    # 等待表格
                    table_selector = 'table, .el-table'
                    logger.info("  等待表格出现...")
                    if page.wait_for_selector(table_selector, timeout=60000):
                        logger.info("  ✅ 表格已出现")
                        time.sleep(5)
                    else:
                        logger.warning("  ⚠️ 表格未出现，尝试提取")

                    page.screenshot(path='data/05_device_table.png')

                    self.device_data = self.get_table_data(page)
                    logger.info(f"  ✅ 设备数据: {len(self.device_data)} 条")
                    self.device_success = True

                except Exception as e:
                    logger.error(f"  ❌ 抓取设备数据失败: {e}")
                    page.screenshot(path='data/05_device_error.png')
                    self.error_message += f"设备数据抓取失败: {e}"

                # ========== 6. 抓取跨区管控数据 ==========
                logger.info("【步骤6】抓取跨区管控数据")
                try:
                    logger.info("  跳转到跨区管控页面...")
                    page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
                    time.sleep(15)

                    table_selector = 'table, .el-table'
                    logger.info("  等待表格出现...")
                    if page.wait_for_selector(table_selector, timeout=60000):
                        logger.info("  ✅ 表格已出现")
                        time.sleep(5)
                    else:
                        logger.warning("  ⚠️ 表格未出现")

                    page.screenshot(path='data/06_cross_region_table.png')

                    self.cross_region_data = self.get_table_data(page)
                    logger.info(f"  ✅ 跨区管控数据: {len(self.cross_region_data)} 条")
                    self.cross_region_success = True

                except Exception as e:
                    logger.error(f"  ❌ 抓取跨区管控数据失败: {e}")
                    page.screenshot(path='data/06_cross_region_error.png')
                    self.error_message += f"跨区管控数据抓取失败: {e}"

                browser.close()

            # ========== 7. 保存数据 ==========
            logger.info("【步骤7】保存数据")
            self.save_data()

            # ========== 8. 发送邮件 ==========
            logger.info("【步骤8】发送邮件")
            self.send_email(success=True)

            logger.info("=" * 70)
            logger.info("✅ 爬取流程结束")
            logger.info("=" * 70)
            return True

        except Exception as e:
            logger.error(f"❌ 程序主流程异常: {e}")
            self.error_message += f"主流程异常: {e}"
            try:
                self.send_email(success=False)
            except:
                logger.error("发送失败邮件也出错了")
            return False

    def get_table_data(self, page):
        """提取表格数据（增强版，带重试）"""
        data = []
        for attempt in range(3):
            try:
                logger.info(f"    尝试提取表格数据 (第{attempt+1}次)")
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
                    logger.warning("      未找到行数据，可能页面无数据")
                    if page.locator('text=暂无数据').count() > 0:
                        logger.warning("      页面显示：暂无数据")
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
                logger.error(f"      提取尝试失败: {e}")
                time.sleep(3)
                if attempt == 2:
                    logger.error("      三次尝试均失败")
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
        """发送邮件（强制发送）"""
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
                    h2 {{ color: #333; }}
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
                        <td>选择环境</td>
                        <td>{'✅' if self.env_selected else '⚠️'}</td>
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
                
                {f'<p><strong>错误信息：</strong>{self.error_message}</p>' if self.error_message else ''}
                
                <p>附件为数据文件（如有）。</p>
                <hr>
                <p style="color: #666; font-size: 12px;">此邮件由自动化系统发送，请勿回复。</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))

            # 添加附件
            files_to_attach = []
            for f in [self.device_excel, self.cross_region_excel, self.device_json, self.cross_region_json]:
                if Path(f).exists():
                    files_to_attach.append(f)

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
