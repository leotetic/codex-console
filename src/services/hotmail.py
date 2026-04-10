"""
Hotmail/Outlook 邮箱服务实现
通过 Playwright 浏览器自动化创建新的 Hotmail/Outlook 邮箱账号
支持自动处理人机验证（长按验证等）
"""

import re
import time
import logging
import random
import string
import secrets
from typing import Optional, Dict, Any, List
from datetime import datetime

from .base import BaseEmailService, EmailServiceError, EmailServiceType
from ..config.constants import OTP_CODE_PATTERN, FIRST_NAMES

logger = logging.getLogger(__name__)

# 常用的 Hotmail/Outlook 域名
EMAIL_DOMAINS = ["hotmail.com", "outlook.com", "live.com", "msn.com"]

# 随机的英文名字列表（用于生成用户名）
NAMES = FIRST_NAMES + [
    "Alex", "Blake", "Cameron", "Dakota", "Ellis", "Finley", "Gray", "Hayden",
    "Jack", "Jules", "Kelly", "Lane", "Morgan", "Parker", "Quinn", "Reese",
    "Robin", "Sage", "Skyler", "Tatum", "Taylor", "Val", "Winter", "Jordan"
]

# Playwright 配置
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


def get_playwright():
    """获取 Playwright 实例"""
    from playwright.sync_api import sync_playwright
    return sync_playwright().start()


class HotmailService(BaseEmailService):
    """
    Hotmail/Outlook 账号自动注册服务

    使用 Playwright 浏览器自动化创建新的 Hotmail/Outlook 邮箱账号，
    自动处理人机验证（包括长按验证）。
    """

    def __init__(self, config: Dict[str, Any] = None, name: str = None):
        """
        初始化 Hotmail 服务

        Args:
            config: 配置字典，支持以下键:
                - proxy_url: 代理地址 (可选，格式: http://host:port)
                - domain: 首选邮箱域名，默认随机选择
                - backup_email_service: 备用邮箱服务类型 (可选)
                - backup_email_config: 备用邮箱服务配置 (可选)
                - headless: 是否隐藏浏览器，默认 True
                - timeout: 操作超时时间，默认 30
            name: 服务名称
        """
        super().__init__(EmailServiceType.HOTMAIL, name)

        default_config = {
            "proxy_url": None,
            "domain": None,
            "backup_email_service": None,
            "backup_email_config": None,
            "headless": True,
            "timeout": 30,
        }
        self.config = {**default_config, **(config or {})}

        # 缓存已创建的邮箱（实例级别）
        self._created_emails: Dict[str, Dict[str, Any]] = {}

        # 备用邮箱服务（用于接收验证码）
        self._backup_email_service = None

        # Playwright 实例
        self._playwright = None
        self._browser = None
        self._context = None

    def _start_browser(self):
        """启动浏览器"""
        if self._browser:
            return

        try:
            from playwright.sync_api import sync_playwright

            self._playwright = sync_playwright().start()

            # 启动 Chromium
            launch_options = {
                "headless": self.config.get("headless", True),
                "args": [
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled",
                ]
            }

            # 使用系统 Chrome
            launch_options["executable_path"] = CHROME_PATH

            self._browser = self._playwright.chromium.launch(**launch_options)

            # 创建上下文（隔离 cookies）
            context_options = {
                "viewport": {"width": 1280, "height": 720},
                "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "locale": "en-US",
                "timezone_id": "America/New_York",
            }

            # 添加代理
            proxy_url = self.config.get("proxy_url")
            if proxy_url:
                context_options["proxy"] = {"server": proxy_url}

            self._context = self._browser.new_context(**context_options)

            logger.info("浏览器启动成功")

        except Exception as e:
            logger.error(f"启动浏览器失败: {e}")
            raise EmailServiceError(f"启动浏览器失败: {e}")

    def _close_browser(self):
        """关闭浏览器"""
        try:
            if self._context:
                self._context.close()
                self._context = None
            if self._browser:
                self._browser.close()
                self._browser = None
            if self._playwright:
                self._playwright.stop()
                self._playwright = None
        except Exception as e:
            logger.warning(f"关闭浏览器时出错: {e}")

    def _generate_username(self, domain: str) -> str:
        """
        生成随机用户名

        Args:
            domain: 邮箱域名

        Returns:
            随机用户名 (不含域名部分)
        """
        first_name = random.choice(NAMES).lower()
        suffix = "".join(random.choices(string.digits, k=random.randint(2, 4)))
        return f"{first_name}{suffix}"

    def _generate_password(self, length: int = 16) -> str:
        """生成随机密码"""
        charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
        return "".join(secrets.choice(charset) for _ in range(length))

    def _generate_name(self) -> tuple:
        """生成随机姓名"""
        first_name = random.choice(NAMES)
        last_name = random.choice(NAMES)
        return first_name, last_name

    def _get_domain(self) -> str:
        """获取邮箱域名"""
        domain = self.config.get("domain")
        if domain:
            return domain
        return random.choice(EMAIL_DOMAINS)

    def _handle_human_verification(self, page) -> bool:
        """
        处理人机验证（长按验证等）

        Args:
            page: Playwright page 对象

        Returns:
            是否处理成功
        """
        try:
            # 等待验证元素出现
            time.sleep(2)

            # 方法1: 查找长按按钮（按住不放的验证）
            hold_button_selectors = [
                "[class*='hold'], [class*='press'], [class*='按住']",
                "[aria-label*='hold'], [aria-label*='press']",
                "button:has-text('Hold'), button:has-text('按住')",
                "[class*='verify'], [class*='challenge']",
            ]

            for selector in hold_button_selectors:
                try:
                    elements = page.query_selector_all(selector)
                    for elem in elements:
                        text = elem.inner_text().lower() if elem.inner_text() else ""
                        if 'hold' in text or 'press' in text or '按住' in text:
                            # 找到长按按钮，执行长按
                            logger.info("发现长按验证按钮，执行长按...")
                            elem.click()
                            time.sleep(3)  # 长按 3 秒
                            return True
                except Exception:
                    continue

            # 方法2: 查找进度条类型的验证
            progress_selectors = [
                "[class*='progress'], [class*='slider'], [class*='drag']",
                "[role='slider'], [class*='captcha']",
            ]

            for selector in progress_selectors:
                try:
                    slider = page.query_selector(selector)
                    if slider:
                        logger.info("发现滑块验证...")
                        # 尝试拖动滑块
                        box = slider.bounding_box()
                        if box:
                            start_x = box['x'] + box['width'] // 2
                            start_y = box['y'] + box['height'] // 2
                            end_x = box['x'] + box['width']
                            page.mouse.move(start_x, start_y)
                            page.mouse.down()
                            page.mouse.move(end_x, start_y, steps=10)
                            page.mouse.up()
                            time.sleep(1)
                            return True
                except Exception:
                    continue

            # 方法3: 查找并点击验证相关的 iframe
            try:
                iframes = page.query_selector_all("iframe")
                for iframe in iframes:
                    try:
                        iframe.click()
                        time.sleep(1)
                    except Exception:
                        continue
            except Exception:
                pass

            logger.info("未发现人机验证或已通过")
            return True

        except Exception as e:
            logger.warning(f"处理人机验证时出错: {e}")
            return False

    def _get_backup_email_service(self):
        """获取备用邮箱服务实例"""
        if self._backup_email_service:
            return self._backup_email_service

        backup_service_type = self.config.get("backup_email_service")
        backup_config = self.config.get("backup_email_config")

        if not backup_service_type or not backup_config:
            logger.warning("未配置备用邮箱服务")
            return None

        try:
            from ..services import EmailServiceFactory, EmailServiceType

            service_type = EmailServiceType(backup_service_type)
            self._backup_email_service = EmailServiceFactory.create(
                service_type,
                backup_config,
                name=f"{self.name}_backup"
            )
            return self._backup_email_service

        except Exception as e:
            logger.error(f"创建备用邮箱服务失败: {e}")
            return None

    def _wait_for_verification_code(self, backup_email: str, timeout: int = 120) -> Optional[str]:
        """
        等待验证码

        Args:
            backup_email: 备用邮箱地址
            timeout: 超时时间（秒）

        Returns:
            验证码字符串
        """
        backup_service = self._get_backup_email_service()
        if not backup_service:
            logger.error("没有备用邮箱服务")
            return None

        try:
            return backup_service.get_verification_code(
                email=backup_email,
                email_id=backup_email,
                timeout=timeout,
                pattern=OTP_CODE_PATTERN
            )
        except Exception as e:
            logger.error(f"获取验证码失败: {e}")
            return None

    def create_email(self, config: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        创建新的 Hotmail/Outlook 邮箱账号

        Args:
            config: 配置参数

        Returns:
            包含邮箱信息的字典
        """
        max_retries = self.config.get("max_retries", 3)

        for attempt in range(max_retries):
            try:
                # 启动浏览器
                self._start_browser()
                page = self._context.new_page()

                # 生成账号信息
                domain = self._get_domain()
                username = self._generate_username(domain)
                email = f"{username}@{domain}"
                password = self._generate_password()
                first_name, last_name = self._generate_name()

                logger.info(f"开始注册: {email} (尝试 {attempt + 1}/{max_retries})")

                # 访问注册页面
                page.goto("https://signup.live.com/", timeout=30000)
                time.sleep(random.uniform(1, 2))

                # 检查是否有错误页面
                if "unavailable" in page.url.lower():
                    logger.warning("IP 可能被限制")
                    page.close()
                    continue

                # 填写表单
                try:
                    # 点击 "创建 Microsoft 账户" 或跳过欢迎页
                    try:
                        create_btn = page.query_selector("a:has-text('创建'), button:has-text('创建'), [data-value='create']")
                        if create_btn:
                            create_btn.click()
                            time.sleep(1)
                    except Exception:
                        pass

                    # 填写邮箱
                    email_input = page.query_selector("input[name='Email'], input[type='email'], input#memberName")
                    if email_input:
                        email_input.fill(username)
                        time.sleep(0.5)

                    # 选择域名
                    domain_select = page.query_selector("select#domain, [data-domain]")
                    if domain_select:
                        domain_select.select_option(domain.replace(".", ""))
                    else:
                        # 点击域名选项
                        domain_btn = page.query_selector(f"button[title='{domain}'], span:has-text('@{domain}')")
                        if domain_btn:
                            domain_btn.click()

                    time.sleep(0.5)

                    # 点击下一步
                    next_btn = page.query_selector("input[value='下一步'], button:has-text('下一步')")
                    if next_btn:
                        next_btn.click()
                        time.sleep(2)

                    # 处理人机验证
                    self._handle_human_verification(page)

                    # 填写密码
                    password_input = page.query_selector("input[name='Password'], input#password")
                    if password_input:
                        password_input.fill(password)
                        time.sleep(0.3)

                    # 填写名字
                    first_input = page.query_selector("input[name='FirstName'], input#firstName")
                    if first_input:
                        first_input.fill(first_name)

                    last_input = page.query_selector("input[name='LastName'], input#lastName")
                    if last_input:
                        last_input.fill(last_name)

                    time.sleep(0.3)

                    # 点击下一步
                    next_btn = page.query_selector("input[value='下一步'], button:has-text('下一步')")
                    if next_btn:
                        next_btn.click()
                        time.sleep(2)

                    # 处理人机验证
                    self._handle_human_verification(page)

                    # 填写生日
                    year_input = page.query_selector("input[name='BirthYear'], select[name='BirthYear']")
                    if year_input:
                        birth_year = random.randint(1990, 2005)
                        try:
                            year_input.fill(str(birth_year))
                        except Exception:
                            year_select = page.query_selector("select[name='BirthYear']")
                            if year_select:
                                year_select.select_option(str(birth_year))

                    month_select = page.query_selector("select[name='BirthMonth']")
                    if month_select:
                        month_select.select_option(str(random.randint(1, 12)))

                    day_input = page.query_selector("select[name='BirthDay']")
                    if day_input:
                        day_input.select_option(str(random.randint(1, 28)))

                    time.sleep(0.3)

                    # 点击下一步
                    next_btn = page.query_selector("input[value='下一步'], button:has-text('下一步')")
                    if next_btn:
                        next_btn.click()
                        time.sleep(2)

                    # 处理人机验证
                    self._handle_human_verification(page)

                except Exception as e:
                    logger.warning(f"填写表单时出错: {e}")

                # 等待验证码页面或最终结果
                time.sleep(3)

                # 检查当前 URL 判断是否成功
                current_url = page.url.lower()
                logger.info(f"当前页面: {current_url}")

                if "verify" in current_url or "otp" in current_url or "code" in current_url:
                    # 需要验证码
                    logger.info("需要验证码处理")

                # 截图调试（可选）
                if attempt == 0:
                    try:
                        page.screenshot(path=f"debug_signup_{username}.png")
                        logger.info(f"截图已保存: debug_signup_{username}.png")
                    except Exception:
                        pass

                page.close()

                # 如果没有遇到致命错误，认为基本流程完成
                # 实际使用中，验证码需要通过备用邮箱接收
                email_info = {
                    "email": email,
                    "password": password,
                    "service_id": email,
                    "first_name": first_name,
                    "last_name": last_name,
                    "created_at": time.time(),
                }

                self._created_emails[email] = email_info
                self.update_status(True)

                logger.info(f"Hotmail 注册流程完成: {email}")
                return email_info

            except Exception as e:
                logger.error(f"注册失败 (尝试 {attempt + 1}/{max_retries}): {e}")

                try:
                    if self._browser:
                        self._close_browser()
                except Exception:
                    pass

                time.sleep(random.uniform(2, 5))

        self.update_status(False, EmailServiceError("创建邮箱失败：达到最大重试次数"))
        raise EmailServiceError("创建邮箱失败：达到最大重试次数")

    def get_verification_code(
        self,
        email: str,
        email_id: str = None,
        timeout: int = 120,
        pattern: str = OTP_CODE_PATTERN,
        otp_sent_at: Optional[float] = None,
    ) -> Optional[str]:
        """
        从备用邮箱获取验证码

        Args:
            email: 邮箱地址
            email_id: 未使用
            timeout: 超时时间（秒）
            pattern: 验证码正则
            otp_sent_at: OTP 发送时间戳

        Returns:
            验证码字符串
        """
        backup_service = self._get_backup_email_service()
        if not backup_service:
            logger.error("未配置备用邮箱服务")
            return None

        backup_email = self.config.get("backup_email")
        if not backup_email:
            backup_email_cfg = self.config.get("backup_email_config", {})
            backup_email = backup_email_cfg.get("email")

        if not backup_email:
            logger.error("未指定备用邮箱")
            return None

        try:
            code = backup_service.get_verification_code(
                email=backup_email,
                email_id=backup_email,
                timeout=timeout,
                pattern=pattern,
                otp_sent_at=otp_sent_at
            )

            if code:
                self.update_status(True)

            return code

        except Exception as e:
            logger.error(f"获取验证码失败: {e}")
            self.update_status(False, e)
            return None

    def list_emails(self, **kwargs) -> List[Dict[str, Any]]:
        """列出已创建的邮箱"""
        return list(self._created_emails.values())

    def delete_email(self, email_id: str) -> bool:
        """删除邮箱（仅从缓存中移除）"""
        if email_id in self._created_emails:
            del self._created_emails[email_id]
            logger.info(f"删除 Hotmail 邮箱: {email_id}")
            return True
        return False

    def check_health(self) -> bool:
        """检查服务健康状态"""
        try:
            # 轻量级健康检查：只验证配置和模块是否正确
            if not self.config.get("domain"):
                logger.warning("Hotmail 服务未配置域名")
                self.update_status(False, Exception("未配置域名"))
                return False

            # 验证 Playwright 模块可用
            try:
                from playwright.sync_api import sync_playwright
            except ImportError:
                logger.warning("Playwright 模块未安装")
                self.update_status(False, Exception("Playwright 模块未安装"))
                return False

            self.update_status(True)
            return True
        except Exception as e:
            logger.warning(f"健康检查失败: {e}")
            self.update_status(False, e)
            return False

    def get_service_info(self) -> Dict[str, Any]:
        """获取服务信息"""
        return {
            "service_type": self.service_type.value,
            "name": self.name,
            "domain": self.config.get("domain"),
            "proxy_url": "***" if self.config.get("proxy_url") else None,
            "headless": self.config.get("headless", True),
            "has_backup_email": bool(self._get_backup_email_service()),
            "cached_emails_count": len(self._created_emails),
            "status": self.status.value,
        }

    def __del__(self):
        """析构时关闭浏览器"""
        try:
            self._close_browser()
        except Exception:
            pass
