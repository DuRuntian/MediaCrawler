#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
携程游记爬虫
参考MediaCrawler项目架构，使用CDP模式连接本地浏览器
支持登录后爬取游记内容
"""

import asyncio
import os
import re
import time
import json
from typing import Dict, List, Optional
from urllib.parse import quote

from playwright.async_api import (
    BrowserContext,
    Page,
    Playwright,
    async_playwright,
)

# ==================== 配置参数 ====================
# 关键词
KEYWORDS = "兰州"
# 爬取页数
START_PAGE = 1
END_PAGE = 5
# 每页最大游记数
MAX_TRAVELS_PER_PAGE = 20
# 保存路径
SAVE_PATH = "./data/ctrip"
# 保存格式: jsonl / excel / txt
SAVE_FORMAT = "jsonl"
# 登录方式: qrcode / cookie
LOGIN_TYPE = "qrcode"
# Cookie字符串（如果使用cookie登录方式）
COOKIES = ""
# 是否使用CDP模式（连接本地浏览器）
ENABLE_CDP_MODE = True
# CDP调试端口
CDP_DEBUG_PORT = 9222
# 是否连接已打开的浏览器
CDP_CONNECT_EXISTING = True
# 自定义浏览器路径（留空自动检测）
CUSTOM_BROWSER_PATH = ""
# 是否无头模式
HEADLESS = False
# 是否保存登录状态
SAVE_LOGIN_STATE = True
# 爬取间隔（秒）
CRAWLER_SLEEP_SEC = 2
# 浏览器超时（秒）
BROWSER_LAUNCH_TIMEOUT = 60
# ================================================


class CtripTravelCrawler:
    """携程游记爬虫"""

    def __init__(self):
        self.index_url = "https://you.ctrip.com"
        self.search_url_template = "https://you.ctrip.com/searchsite/travels/?query={keyword}&isAnswered=&isRecommended=&publishDate=&PageNo={page}"
        self.browser_context: Optional[BrowserContext] = None
        self.context_page: Optional[Page] = None
        self.user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
        self.cdp_manager = None
        self.all_travels_data: List[Dict] = []

    async def start(self):
        """启动爬虫"""
        print("=" * 50)
        print("携程游记爬虫启动")
        print(f"关键词: {KEYWORDS}")
        print(f"页数范围: {START_PAGE} - {END_PAGE}")
        print(f"CDP模式: {ENABLE_CDP_MODE}")
        print("=" * 50)

        # 创建保存目录
        os.makedirs(SAVE_PATH, exist_ok=True)

        async with async_playwright() as playwright:
            # 启动浏览器
            if ENABLE_CDP_MODE:
                print("[浏览器] 使用CDP模式启动...")
                self.browser_context = await self.launch_browser_with_cdp(playwright)
            else:
                print("[浏览器] 使用标准模式启动...")
                chromium = playwright.chromium
                self.browser_context = await self.launch_browser_standard(chromium)

            # 创建页面
            self.context_page = await self.browser_context.new_page()

            # 登录
            await self.login()

            # 搜索并爬取
            await self.search_and_crawl()

            # 保存数据
            await self.save_data()

            # 关闭浏览器
            await self.close_browser()

        print(f"\n爬取完成！共获取 {len(self.all_travels_data)} 篇游记")

    async def launch_browser_with_cdp(self, playwright: Playwright) -> BrowserContext:
        """使用CDP模式启动浏览器"""
        from tools.cdp_browser import CDPBrowserManager
        from tools.browser_launcher import BrowserLauncher

        self.cdp_manager = CDPBrowserManager()

        # 连接已存在的浏览器或启动新浏览器
        if CDP_CONNECT_EXISTING:
            print(f"[CDP] 连接已存在的浏览器，端口 {CDP_DEBUG_PORT}")
            print("[CDP] 请确保浏览器已开启远程调试: chrome://inspect/#remote-debugging")
            browser_context = await self.cdp_manager.launch_and_connect(
                playwright,
                playwright_proxy=None,
                user_agent=self.user_agent,
                headless=HEADLESS,
            )
        else:
            # 自动检测并启动浏览器
            launcher = BrowserLauncher()
            browser_paths = launcher.detect_browser_paths()

            if CUSTOM_BROWSER_PATH and os.path.isfile(CUSTOM_BROWSER_PATH):
                browser_path = CUSTOM_BROWSER_PATH
            elif browser_paths:
                browser_path = browser_paths[0]
                browser_name, browser_version = launcher.get_browser_info(browser_path)
                print(f"[CDP] 检测到浏览器: {browser_name} ({browser_version})")
            else:
                raise RuntimeError("未找到本地浏览器，请安装Chrome或Edge")

            browser_context = await self.cdp_manager.launch_and_connect(
                playwright,
                playwright_proxy=None,
                user_agent=self.user_agent,
                headless=HEADLESS,
            )

        return browser_context

    async def launch_browser_standard(self, chromium, headless: bool = True) -> BrowserContext:
        """标准模式启动浏览器"""
        browser_context = await chromium.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=self.user_agent,
        )

        # 添加反检测脚本
        if os.path.exists("libs/stealth.min.js"):
            await browser_context.add_init_script(path="libs/stealth.min.js")

        return browser_context

    async def login(self):
        """登录携程"""
        print("\n[登录] 开始登录流程...")

        # 访问携程首页
        await self.context_page.goto(self.index_url, timeout=30000)
        await asyncio.sleep(2)

        # 检查登录状态
        is_logged_in = await self.check_login_state()

        if is_logged_in:
            print("[登录] 已登录，跳过登录流程")
            return

        if LOGIN_TYPE == "qrcode":
            await self.login_by_qrcode()
        elif LOGIN_TYPE == "cookie":
            await self.login_by_cookie()

    async def check_login_state(self) -> bool:
        """检查登录状态"""
        try:
            # 检查是否有用户头像或用户名元素
            user_selector = "xpath=//div[contains(@class, 'user')]//a[contains(@href, '/user/')]"
            is_visible = await self.context_page.is_visible(user_selector, timeout=3000)

            if is_visible:
                print("[登录] 检测到用户元素，已登录")
                return True

            # 检查Cookie
            cookies = await self.browser_context.cookies()
            cookie_dict = {c['name']: c['value'] for c in cookies}

            # 检查关键Cookie
            if cookie_dict.get('_u') or cookie_dict.get('ticket'):
                print("[登录] 检测到登录Cookie")
                return True

        except Exception as e:
            print(f"[登录] 检查登录状态异常: {e}")

        return False

    async def login_by_qrcode(self):
        """扫码登录"""
        print("[登录] 请使用携程App扫码登录...")

        # 点击登录按钮
        try:
            login_btn = await self.context_page.wait_for_selector(
                "xpath=//a[contains(text(), '登录') or contains(@class, 'login')]",
                timeout=5000
            )
            if login_btn:
                await login_btn.click()
                await asyncio.sleep(2)
        except Exception:
            print("[登录] 未找到登录按钮，可能已弹出登录框")

        # 等待用户扫码
        max_wait_time = 120  # 最大等待时间
        start_time = time.time()

        while time.time() - start_time < max_wait_time:
            is_logged_in = await self.check_login_state()
            if is_logged_in:
                print("[登录] 扫码登录成功！")

                # 保存登录状态
                if SAVE_LOGIN_STATE:
                    await self.save_login_state()
                return

            await asyncio.sleep(2)
            remaining = int(max_wait_time - (time.time() - start_time))
            if remaining % 10 == 0:
                print(f"[登录] 等待扫码... 剩余 {remaining} 秒")

        raise RuntimeError("扫码登录超时")

    async def login_by_cookie(self):
        """Cookie登录"""
        print("[登录] 使用Cookie登录...")

        if not COOKIES:
            raise ValueError("未配置COOKIES")

        # 解析Cookie字符串
        cookie_list = []
        for item in COOKIES.split(';'):
            item = item.strip()
            if '=' in item:
                name, value = item.split('=', 1)
                cookie_list.append({
                    'name': name.strip(),
                    'value': value.strip(),
                    'domain': '.ctrip.com',
                    'path': '/'
                })

        await self.browser_context.add_cookies(cookie_list)
        print(f"[登录] 已添加 {len(cookie_list)} 个Cookie")

        # 刷新页面验证
        await self.context_page.reload()
        await asyncio.sleep(2)

        is_logged_in = await self.check_login_state()
        if not is_logged_in:
            raise RuntimeError("Cookie登录失败，请检查Cookie是否有效")

        print("[登录] Cookie登录成功")

    async def save_login_state(self):
        """保存登录状态"""
        user_data_dir = os.path.join(os.getcwd(), "browser_data", "ctrip_user_data_dir")
        os.makedirs(user_data_dir, exist_ok=True)

        cookies = await self.browser_context.cookies()
        cookie_file = os.path.join(user_data_dir, "cookies.json")

        with open(cookie_file, 'w', encoding='utf-8') as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)

        print(f"[登录] 已保存登录状态到: {cookie_file}")

    async def search_and_crawl(self):
        """搜索并爬取游记"""
        print("\n[爬取] 开始搜索游记...")

        for page_num in range(START_PAGE, END_PAGE + 1):
            search_url = self.search_url_template.format(
                keyword=quote(KEYWORDS),
                page=page_num
            )

            print(f"\n[爬取] 第 {page_num} 页: {search_url}")

            try:
                await self.context_page.goto(search_url, timeout=30000)
                await asyncio.sleep(CRAWLER_SLEEP_SEC)

                # 解析列表页获取游记链接
                travel_links = await self.parse_travel_list()

                print(f"[爬取] 第 {page_num} 页找到 {len(travel_links)} 篇游记")

                # 爬取每篇游记详情
                for idx, travel_info in enumerate(travel_links[:MAX_TRAVELS_PER_PAGE]):
                    print(f"[爬取] 正在爬取第 {idx + 1} 篇: {travel_info['title'][:30]}...")
                    detail_data = await self.crawl_travel_detail(travel_info)
                    self.all_travels_data.append(detail_data)
                    await asyncio.sleep(CRAWLER_SLEEP_SEC)

            except Exception as e:
                print(f"[爬取] 第 {page_num} 页爬取失败: {e}")
                continue

    async def parse_travel_list(self) -> List[Dict]:
        """解析游记列表页"""
        travel_links = []

        try:
            # 等待页面加载
            await self.context_page.wait_for_load_state("networkidle", timeout=10000)

            # 使用XPath获取游记链接（参考教程）
            # 携程搜索页游记链接结构: /html/body/div[2]/div[2]/div[2]/div/div[1]/ul/li
            travel_items = await self.context_page.locator(
                "xpath=//ul[contains(@class, 'search_result_list')]//li//a[contains(@href, '/travels/')]"
            ).all()

            if not travel_items:
                # 备选方案
                travel_items = await self.context_page.locator(
                    "xpath=//a[contains(@href, '/travels/') and contains(@href, '.html')]"
                ).all()

            print(f"[解析] 找到 {len(travel_items)} 个游记链接元素")

            for item in travel_items:
                try:
                    href = await item.get_attribute('href')
                    if not href or '/travels/' not in href:
                        continue

                    # 获取标题
                    title = await item.inner_text()
                    title = title.strip() if title else ''

                    if not title:
                        # 从链接中提取ID作为标题
                        match = re.search(r'/travels/[^/]+/(\d+)\.html', href)
                        if match:
                            title = f"游记-{match.group(1)}"

                    # 补全URL
                    if href.startswith('/'):
                        url = f"https://you.ctrip.com{href}"
                    else:
                        url = href

                    travel_links.append({
                        'title': title,
                        'url': url,
                    })

                except Exception as e:
                    print(f"[解析] 解析单个游记链接失败: {e}")
                    continue

        except Exception as e:
            print(f"[解析] 解析列表页失败: {e}")

        # 去重
        seen_urls = set()
        unique_links = []
        for link in travel_links:
            if link['url'] not in seen_urls:
                seen_urls.add(link['url'])
                unique_links.append(link)

        return unique_links

    async def crawl_travel_detail(self, travel_info: Dict) -> Dict:
        """爬取游记详情页"""
        detail_data = travel_info.copy()

        try:
            await self.context_page.goto(travel_info['url'], timeout=30000)
            await asyncio.sleep(1)

            # 等待页面加载完成
            await self.context_page.wait_for_load_state("networkidle", timeout=10000)

            # 滚动页面加载完整内容
            await self.scroll_page()

            # 解析详情页
            await self.parse_travel_detail(detail_data)

        except Exception as e:
            print(f"[详情] 爬取详情页失败: {e}")
            detail_data['content'] = f"获取失败: {e}"
            detail_data['error'] = str(e)

        return detail_data

    async def scroll_page(self):
        """滚动页面加载完整内容"""
        try:
            for i in range(5):
                await self.context_page.evaluate("window.scrollBy(0, 500)")
                await asyncio.sleep(0.5)
        except Exception:
            pass

    async def parse_travel_detail(self, detail_data: Dict):
        """解析游记详情页"""
        try:
            # 提取标题
            try:
                title_elem = await self.context_page.locator(
                    "xpath=//h1[contains(@class, 'title')] | //h2[contains(@class, 'title')] | //div[@class='ctd_head_con']/h1"
                ).first.inner_text()
                if title_elem:
                    detail_data['title'] = title_elem.strip()
            except Exception:
                pass

            # 提取作者
            try:
                author_elem = await self.context_page.locator(
                    "xpath=//div[contains(@class, 'author')]//a | //span[contains(@class, 'author-name')] | //a[contains(@class, 'user')]"
                ).first.inner_text()
                detail_data['author'] = author_elem.strip() if author_elem else '匿名'
            except Exception:
                detail_data['author'] = '匿名'

            # 提取正文内容（关键XPath: class="ctd_content"）
            try:
                content_elem = await self.context_page.locator(
                    "xpath=//div[@class='ctd_content'] | //div[contains(@class, 'ctd_content')] | //div[contains(@class, 'article-content')] | //div[contains(@class, 'travel-content')]"
                ).first.inner_text()
                detail_data['content'] = content_elem.strip() if content_elem else ''
            except Exception:
                # 备选方案：获取所有段落文本
                try:
                    paragraphs = await self.context_page.locator(
                        "xpath=//div[contains(@class, 'content')]//p | //article//p"
                    ).all_inner_texts()
                    detail_data['content'] = '\n'.join(paragraphs)
                except Exception:
                    detail_data['content'] = ''

            # 提取游记信息（天数、人均、时间等）
            try:
                info_items = await self.context_page.locator(
                    "xpath=//ul[contains(@class, 'ctd_des_list')]//li | //div[contains(@class, 'travel-info')]//span"
                ).all_inner_texts()

                info_str = ' '.join(info_items)

                # 提取天数
                days_match = re.search(r'(\d+)\s*天', info_str)
                detail_data['travel_days'] = days_match.group(1) if days_match else ''

                # 提取人均花费
                cost_match = re.search(r'人均[^\d]*(\d+)', info_str)
                detail_data['per_capita_cost'] = cost_match.group(1) if cost_match else ''

                # 提取时间
                time_match = re.search(r'时间[^\d]*(\d+[^\d]*\d*[^\d]*\d*)', info_str)
                detail_data['travel_time'] = time_match.group(1) if time_match else ''

                # 提取和谁出行
                who_match = re.search(r'和谁[^\d\w]*(\w+)', info_str)
                detail_data['travel_partner'] = who_match.group(1) if who_match else ''

            except Exception:
                pass

            # 提取发布时间
            try:
                time_elem = await self.context_page.locator(
                    "xpath=//div[@class='time'] | //span[contains(@class, 'time')] | //div[contains(@class, 'date')]"
                ).first.inner_text()
                time_match = re.search(r'(\d{4}-\d{1,2}-\d{1,2})', time_elem)
                detail_data['publish_date'] = time_match.group(1) if time_match else time_elem.strip()
            except Exception:
                detail_data['publish_date'] = ''

            # 提取图片链接
            try:
                img_elements = await self.context_page.locator(
                    "xpath=//div[contains(@class, 'ctd_content')]//img | //div[contains(@class, 'content')]//img | //article//img"
                ).all()

                img_urls = []
                for img in img_elements[:20]:
                    src = await img.get_attribute('src')
                    if src and not src.startswith('data:'):
                        if src.startswith('//'):
                            src = f'https:{src}'
                        img_urls.append(src)

                detail_data['image_urls'] = img_urls
                detail_data['image_count'] = len(img_urls)
            except Exception:
                detail_data['image_urls'] = []
                detail_data['image_count'] = 0

            # 提取浏览数、评论数、喜欢数
            try:
                view_elem = await self.context_page.locator(
                    "xpath=//span[contains(@class, 'view')] | //span[contains(text(), '浏览')]"
                ).first.inner_text()
                view_match = re.search(r'(\d+)', view_elem)
                detail_data['view_count'] = view_match.group(1) if view_match else '0'
            except Exception:
                detail_data['view_count'] = '0'

            print(f"[详情] 内容长度: {len(detail_data.get('content', ''))} 字符, 图片: {detail_data.get('image_count', 0)} 张")

        except Exception as e:
            print(f"[详情] 解析详情页失败: {e}")

    async def save_data(self):
        """保存数据"""
        print(f"\n[保存] 保存数据到 {SAVE_PATH}...")

        filename = os.path.join(SAVE_PATH, f"{KEYWORDS}_游记.{SAVE_FORMAT}")

        if SAVE_FORMAT == "jsonl":
            with open(filename, 'w', encoding='utf-8') as f:
                for item in self.all_travels_data:
                    f.write(json.dumps(item, ensure_ascii=False) + '\n')

        elif SAVE_FORMAT == "json":
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.all_travels_data, f, ensure_ascii=False, indent=2)

        elif SAVE_FORMAT == "txt":
            with open(filename, 'w', encoding='utf-8') as f:
                for i, item in enumerate(self.all_travels_data, 1):
                    f.write(f"{'='*50}\n")
                    f.write(f"【第{i}篇】 {item.get('title', '')}\n")
                    f.write(f"作者: {item.get('author', '')}\n")
                    f.write(f"链接: {item.get('url', '')}\n")
                    f.write(f"天数: {item.get('travel_days', '')}\n")
                    f.write(f"人均: {item.get('per_capita_cost', '')}\n")
                    f.write(f"发布时间: {item.get('publish_date', '')}\n")
                    f.write(f"浏览: {item.get('view_count', '')}\n")
                    f.write(f"图片数: {item.get('image_count', 0)}\n")
                    f.write(f"\n正文:\n{item.get('content', '')}\n")
                    if item.get('image_urls'):
                        f.write(f"\n图片链接:\n{'; '.join(item.get('image_urls', []))}\n")
                    f.write(f"{'='*50}\n\n")

        elif SAVE_FORMAT == "excel":
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = KEYWORDS

            headers = ['标题', '作者', '链接', '正文', '天数', '人均', '发布时间', '浏览数', '图片数', '图片链接']
            ws.append(headers)

            for item in self.all_travels_data:
                row = [
                    item.get('title', ''),
                    item.get('author', ''),
                    item.get('url', ''),
                    item.get('content', ''),
                    item.get('travel_days', ''),
                    item.get('per_capita_cost', ''),
                    item.get('publish_date', ''),
                    item.get('view_count', ''),
                    item.get('image_count', 0),
                    '; '.join(item.get('image_urls', []))
                ]
                ws.append(row)

            wb.save(filename)

        print(f"[保存] 已保存 {len(self.all_travels_data)} 条数据到: {filename}")

    async def close_browser(self):
        """关闭浏览器"""
        if self.cdp_manager:
            await self.cdp_manager.cleanup()
        elif self.browser_context:
            await self.browser_context.close()


async def main():
    """主函数"""
    crawler = CtripTravelCrawler()
    await crawler.start()


if __name__ == '__main__':
    asyncio.run(main())