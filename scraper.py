# -*- coding: utf-8 -*-
"""
抓取数据1: 代理限制 -> 设备 -> 授权状态="已限制"
抓取数据2: 代理限制 -> 跨区管控 -> 风险事件="高"
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
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ==================== 配置日志 ====================
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
        # 登录信息（从环境变量读取，更安全）
        self.username = os.getenv('AKUBELA_USERNAME', 'super')
        self.password = os.getenv('AKUBELA_PASSWORD', 'Akubela@super2024&')
        self.environment = os.getenv('AKUBELA_ENV', 'prod-cn-hz')

        # 邮件配置（从环境变量读取）
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.exmail.qq.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '465'))
        self.sender_email = os.getenv('SENDER_EMAIL', 'enwade.chen@akuvox.com')
        self.sender_password = os.getenv('SENDER_PASSWORD', 'cyw110120')
        self.receiver_email = os.getenv('RECEIVER_EMAIL', 'enwade.chen@akuvox.com')

        # 数据存储
        self.data1 = []  # 设备-已限制
        self.data2 = []  # 跨区管控-高
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 创建数据目录
        Path('./data').mkdir(exist_ok=True)

        # 文件路径
        self.data1_json = f'data/data1_devices_restricted_{self.timestamp}.json'
        self.data2_json = f'data/data2_cross_region_high_{self.timestamp}.json'
        self.data1_excel = f'data/data1_devices_restricted_{self.timestamp}.xlsx'
        self.data2_excel = f'data/data2_cross_region_high_{self.timestamp}.xlsx'

        # 状态标记
        self.login_success = False
        self.env_selected = False
        self.data1_success = False
        self.data2_success = False
        self.error_message = ""

    # ==================== 辅助函数 ====================
    def wait_for_any(self, page, selectors, timeout=60000):
        """等待多个选择器中的任意一个出现"""
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

    def safe_click(self, page, selector, timeout=30000, retry=3):
        """安全点击，带重试机制"""
        for i in range(retry):
            try:
                page.wait_for_selector(selector, timeout=timeout)
                page.click(selector)
                logger.info(f"  ✅ 成功点击: {selector}")
                return True
            except Exception as e:
                logger.warning(f"  点击失败 (尝试 {i+1}/{retry}): {selector}, 错误: {e}")
                time.sleep(2)
        logger.error(f"  ❌ 点击失败，已重试 {retry} 次: {selector}")
        return False

    def select_dropdown_option(self, page, label_text, option_text, timeout=30000):
        """
        通用下拉框选择函数
        :param label_text: 下拉框前面的标签文字，如 "授权状态"
        :param option_text: 要选择的选项文字，如 "已限制"
        """
        logger.info(f"  正在选择下拉框: {label_text} -> {option_text}")
        try:
            # 1. 找到并点击下拉框
            # 先尝试通过标签定位关联的下拉框
            label_selector = f'text={label_text}'
            if page.locator(label_selector).count() > 0:
                # 点击标签旁边的下拉框（父级或兄弟元素）
                # 这里使用更通用的方法：找到包含标签的容器，再找下拉框
                container = page.locator(label_selector).locator('xpath=ancestor::*[1]')
                # 在容器内找下拉框
                dropdown = container.locator('div[role="combobox"], .el-select, select').first
                if dropdown.count() > 0:
                    dropdown.click()
                    time.sleep(2)
                else:
                    # 如果没找到，尝试直接找可见的下拉框
                    page.locator('div[role="combobox"], .el-select, select').first.click()
                    time.sleep(2)
            else:
                # 如果没找到标签，直接找可见的下拉框
                page.locator('div[role="combobox"], .el-select, select').first.click()
                time.sleep(2)

            # 2. 等待选项出现并点击
            option_selector = f'text={option_text}'
            if page.wait_for_selector(option_selector, timeout=timeout):
                page.click(option_selector)
                logger.info(f"  ✅ 已选择: {option_text}")
                time.sleep(1)
                return True
            else:
                logger.warning(f"  ⚠️ 未找到选项: {option_text}")
                return False

        except Exception as e:
            logger.error(f"  ❌ 选择下拉框失败: {e}")
            return False

    # ==================== 导航函数 ====================
    def navigate_to_agent_restriction(self, page):
        """导航到代理限制菜单"""
        logger.info("  导航到代理限制...")
        time.sleep(3)
        # 尝试多种方式点击菜单
        menu_selectors = [
            'text=代理限制',
            'a:has-text("代理限制")',
            '.el-menu-item:has-text("代理限制")',
            'span:has-text("代理限制")',
            '//*[contains(text(), "代理限制")]'
        ]
        for selector in menu_selectors:
            try:
                if page.locator(selector).count() > 0:
                    page.click(selector)
                    logger.info(f"  ✅ 已点击代理限制菜单")
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(2)
                    return True
            except:
                continue
        logger.warning("  ⚠️ 无法找到代理限制菜单")
        return False

    def navigate_to_device(self, page):
        """导航到设备页面"""
        logger.info("  导航到设备页面...")
        if not self.navigate_to_agent_restriction(page):
            # 如果菜单点击失败，尝试直接跳转
            logger.warning("  尝试直接跳转到设备页面")
            page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
            page.wait_for_load_state("networkidle", timeout=30000)
            time.sleep(5)
            return

        # 点击设备子菜单
        device_selectors = [
            'text=设备',
            'a:has-text("设备")',
            '.el-menu-item:has-text("设备")',
            'span:has-text("设备")',
            '//*[contains(text(), "设备")]'
        ]
        for selector in device_selectors:
            try:
                if page.locator(selector).count() > 0:
                    page.click(selector)
                    logger.info(f"  ✅ 已点击设备子菜单")
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(3)
                    return
            except:
                continue
        logger.warning("  ⚠️ 无法找到设备子菜单，尝试直接跳转")
        page.goto("https://super.akubela.com/#/distributor/device", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)

    def navigate_to_cross_region(self, page):
        """导航到跨区管控页面"""
        logger.info("  导航到跨区管控...")
        if not self.navigate_to_agent_restriction(page):
            logger.warning("  尝试直接跳转到跨区管控页面")
            page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
            page.wait_for_load_state("networkidle", timeout=30000)
            time.sleep(5)
            return

        # 点击跨区管控子菜单
        cross_selectors = [
            'text=跨区管控',
            'a:has-text("跨区管控")',
            '.el-menu-item:has-text("跨区管控")',
            'span:has-text("跨区管控")',
            '//*[contains(text(), "跨区管控")]'
        ]
        for selector in cross_selectors:
            try:
                if page.locator(selector).count() > 0:
                    page.click(selector)
                    logger.info(f"  ✅ 已点击跨区管控子菜单")
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(3)
                    return
            except:
                continue
        logger.warning("  ⚠️ 无法找到跨区管控子菜单，尝试直接跳转")
        page.goto("https://super.akubela.com/#/distributor/cross-region", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)

    # ==================== 数据提取 ====================
    def get_table_data(self, page):
        """提取表格数据"""
        data = []
        try:
            logger.info("    等待表格加载...")
            # 等待表格出现
            table_selectors = ['table', '.el-table', '.el-table__body']
            table_found = False
            for _ in range(5):  # 最多等待15秒
                for sel in table_selectors:
                    if page.locator(sel).count() > 0:
                        logger.info(f"    ✅ 找到表格: {sel}")
                        table_found = True
                        break
                if table_found:
                    break
                time.sleep(3)

            if not table_found:
                # 检查是否为空状态
                if page.locator('text=暂无数据').count() > 0:
                    logger.warning("    页面显示：暂无数据")
                elif page.locator('.el-table__empty-text').count() > 0:
                    logger.warning("    页面显示：空状态")
                else:
                    logger.warning("    未找到表格")
                return data

            time.sleep(3)

            # 提取行数据
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
                logger.warning("    未找到任何数据行")
                return data

            # 解析每一行
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

            logger.info(f"    ✅ 成功提取 {len(data)} 行数据")

        except Exception as e:
            logger.error(f"    ❌ 提取表格数据失败: {e}")

        return data

    # ==================== 主流程 ====================
    def run(self):
        logger.info("=" * 70)
        logger.info("开始爬取 Akubela 数据")
        logger.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 70)

        try:
            with sync_playwright() as p:
                # 启动浏览器（无头模式，慢速操作以适应网络延迟）
                browser = p.chromium.launch(headless=True, slow_mo=1200)
                page = browser.new_page(viewport={'width': 1920, 'height': 1080})
                page.set_default_timeout(120000)  # 全局超时2分钟

                # ========== 步骤1: 登录 ==========
                logger.info("【步骤1】登录系统")
                try:
                    page.goto("https://super.akubela.com/#/login", timeout=60000)
                    time.sleep(5)

                    # 填写用户名
                    logger.info("  填写用户名...")
                    page.wait_for_selector('input[type="text"]', timeout=30000)
                    page.fill('input[type="text"]', self.username)
                    time.sleep(1)

                    # 填写密码
                    logger.info("  填写密码...")
                    page.fill('input[type="password"]', self.password)
                    time.sleep(1)

                    # 点击登录按钮
                    logger.info("  点击登录按钮...")
                    page.click('button:has-text("登陆")')
                    time.sleep(10)

                    # 验证登录是否成功
                    login_indicators = ['text=Super', 'text=super', '.el-menu', '.sidebar']
                    found = self.wait_for_any(page, login_indicators, timeout=60000)
                    if found:
                        logger.info(f"  ✅ 登录成功，检测到: {found}")
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

                # ========== 步骤2: 选择环境 ==========
                logger.info("【步骤2】选择环境: prod-cn-hz")
                try:
                    time.sleep(8)

                    # 查找并点击环境选择器
                    env_selectors = [
                        'div[role="combobox"]',
                        '.el-select',
                        'select',
                        '//div[contains(@class, "el-select")]'
                    ]
                    clicked = False
                    for sel in env_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                logger.info(f"  点击选择器: {sel}")
                                page.click(sel)
                                time.sleep(3)
                                clicked = True
                                break
                        except:
                            continue

                    if clicked:
                        # 选择目标环境
                        option_selector = f'text={self.environment}'
                        if page.wait_for_selector(option_selector, timeout=30000):
                            page.click(option_selector)
                            time.sleep(3)
                            logger.info(f"  ✅ 已选择环境: {self.environment}")
                            self.env_selected = True
                        else:
                            # 尝试选择第一个选项
                            try:
                                page.locator('.el-select-dropdown__item').first.click()
                                logger.info("  已选择第一个选项")
                                self.env_selected = True
                            except:
                                pass
                    else:
                        logger.warning("  ⚠️ 未找到环境选择器")

                    page.screenshot(path='data/02_after_env.png')

                except Exception as e:
                    logger.warning(f"  ⚠️ 选择环境异常: {e}")

                # ========== 步骤3: 抓取数据1 ==========
                logger.info("【步骤3】抓取数据1: 设备 -> 授权状态=已限制")
                try:
                    # 导航到设备页面
                    self.navigate_to_device(page)
                    page.screenshot(path='data/03_device_page.png')

                    # 等待页面稳定
                    time.sleep(5)

                    # 选择授权状态 = 已限制
                    self.select_dropdown_option(page, "授权状态", "已限制")

                    # 点击搜索按钮
                    logger.info("  点击搜索按钮...")
                    search_selectors = [
                        'button:has-text("搜索")',
                        'button:has-text("查询")',
                        'button:has-text("Search")',
                        '.el-button--primary:has-text("搜索")'
                    ]
                    searched = False
                    for sel in search_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                page.click(sel)
                                logger.info(f"  ✅ 已点击搜索按钮")
                                searched = True
                                break
                        except:
                            continue

                    if not searched:
                        logger.warning("  ⚠️ 未找到搜索按钮")

                    # 等待搜索结果
                    time.sleep(5)
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(3)

                    page.screenshot(path='data/04_device_search_result.png')

                    # 抓取数据
                    self.data1 = self.get_table_data(page)
                    logger.info(f"  ✅ 数据1: 共抓取 {len(self.data1)} 条记录")
                    self.data1_success = True

                except Exception as e:
                    logger.error(f"  ❌ 抓取数据1失败: {e}")
                    self.error_message += "数据1失败;"

                # ========== 步骤4: 抓取数据2 ==========
                logger.info("【步骤4】抓取数据2: 跨区管控 -> 风险事件=高")
                try:
                    # 导航到跨区管控页面
                    self.navigate_to_cross_region(page)
                    page.screenshot(path='data/05_cross_region_page.png')

                    # 等待页面稳定
                    time.sleep(5)

                    # 选择风险事件 = 高
                    self.select_dropdown_option(page, "风险事件", "高")

                    # 点击搜索按钮
                    logger.info("  点击搜索按钮...")
                    search_selectors = [
                        'button:has-text("搜索")',
                        'button:has-text("查询")',
                        'button:has-text("Search")',
                        '.el-button--primary:has-text("搜索")'
                    ]
                    searched = False
                    for sel in search_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                page.click(sel)
                                logger.info(f"  ✅ 已点击搜索按钮")
                                searched = True
                                break
                        except:
                            continue

                    if not searched:
                        logger.warning("  ⚠️ 未找到搜索按钮")

                    # 等待搜索结果
                    time.sleep(5)
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(3)

                    page.screenshot(path='data/06_cross_region_search_result.png')

                    # 抓取数据
                    self.data2 = self.get_table_data(page)
                    logger.info(f"  ✅ 数据2: 共抓取 {len(self.data2)} 条记录")
                    self.data2_success = True

                except Exception as e:
                    logger.error(f"  ❌ 抓取数据2失败: {e}")
                    self.error_message += "数据2失败;"

                browser.close()

            # ========== 步骤5: 保存数据 ==========
            logger.info("【步骤5】保存数据")
            self.save_data()

            # ========== 步骤6: 发送邮件 ==========
            logger.info("【步骤6】发送邮件")
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

    # ==================== 保存数据 ====================
    def save_data(self):
        """保存数据到 JSON 和 Excel"""
        try:
            with open(self.data1_json, 'w', encoding='utf-8') as f:
                json.dump(self.data1, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 数据1 JSON: {self.data1_json}")
        except Exception as e:
            logger.error(f"  ❌ 数据1 JSON保存失败: {e}")

        try:
            with open(self.data2_json, 'w', encoding='utf-8') as f:
                json.dump(self.data2, f, ensure_ascii=False, indent=2)
            logger.info(f"  ✅ 数据2 JSON: {self.data2_json}")
        except Exception as e:
            logger.error(f"  ❌ 数据2 JSON保存失败: {e}")

        try:
            if self.data1:
                pd.DataFrame(self.data1).to_excel(self.data1_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 数据1 Excel: {self.data1_excel}")
            else:
                logger.warning("  ⚠️ 数据1为空")
        except Exception as e:
            logger.error(f"  ❌ 数据1 Excel保存失败: {e}")

        try:
            if self.data2:
                pd.DataFrame(self.data2).to_excel(self.data2_excel, index=False, engine='openpyxl')
                logger.info(f"  ✅ 数据2 Excel: {self.data2_excel}")
            else:
                logger.warning("  ⚠️ 数据2为空")
        except Exception as e:
            logger.error(f"  ❌ 数据2 Excel保存失败: {e}")

    # ==================== 发送邮件 ====================
    def send_email(self, success=True):
        """发送邮件报告"""
        logger.info("正在发送邮件...")
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email

            if success and self.data1_success and self.data2_success:
                status = "✅ 完全成功"
                subject = f'Akubela数据报告 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            elif success:
                status = "⚠️ 部分成功"
                subject = f'Akubela数据报告（部分成功） - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
            else:
                status = "❌ 失败"
                subject = f'Akubela爬虫异常 - {datetime.now().strftime("%Y-%m-%d %H:%M")}'

            msg['Subject'] = subject

            # 邮件正文
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
                <p><strong>状态：</strong><span class="{'success' if success and self.data1_success and self.data2_success else 'partial' if success else 'fail'}">{status}</span></p>
                <p><strong>时间：</strong>{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><strong>环境：</strong>{self.environment}</p>

                <h3>执行结果</h3>
                <table>
                    <tr><th>数据源</th><th>筛选条件</th><th>状态</th><th>记录数</th></tr>
                    <tr>
                        <td>数据1 (设备)</td>
                        <td>授权状态 = 已限制</td>
                        <td>{'✅' if self.data1_success else '❌'}</td>
                        <td>{len(self.data1)}</td>
                    </tr>
                    <tr>
                        <td>数据2 (跨区管控)</td>
                        <td>风险事件 = 高</td>
                        <td>{'✅' if self.data2_success else '❌'}</td>
                        <td>{len(self.data2)}</td>
                    </tr>
                </table>

                {f'<p><strong>错误信息：</strong>{self.error_message}</p>' if self.error_message else ''}

                <p>附件为数据文件（Excel + JSON）。</p>
                <hr>
                <p style="color:#666;font-size:12px;">此邮件由自动化系统发送，请勿回复。</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(body, 'html', 'utf-8'))

            # 添加附件
            for f in [self.data1_excel, self.data2_excel, self.data1_json, self.data2_json]:
                if Path(f).exists():
                    with open(f, 'rb') as fp:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(fp.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{Path(f).name}"')
                        msg.attach(part)
                        logger.info(f"  添加附件: {Path(f).name}")

            # 发送邮件
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
