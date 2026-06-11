#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
携程游记爬虫
功能：爬取携程you.ctrip.com上的游记列表及详情页内容
输出：Excel文件（标题、作者、出游天数、人均花费、正文、图片链接等）
支持自动获取Cookie和关键词搜索
"""

import os
import platform
import subprocess
import time
import re
import requests
from lxml import etree
from multiprocessing.dummy import Pool
import openpyxl
from fake_useragent import UserAgent
from requests.exceptions import RequestException
from urllib.parse import quote

# ==================== 可配置参数（用户可在此修改） ====================
# 关键词（可修改为任意目的地或主题，如"上海"、"厦门"、"海边"等）
KEYWORD = "徐州"
# 爬取页码范围（1到END_PAGE）
START_PAGE = 1
END_PAGE = 5
# 并发线程数（建议3~5，过大易被封）
THREAD_POOL_SIZE = 3
# 重试次数
RETRY_TIMES = 3
# 请求超时（秒）
TIMEOUT = 20
# 输出文件名
OUTPUT_FILE = f'{KEYWORD}游记.xlsx'

# 浏览器配置
# 是否使用本地浏览器（True=使用本地Chrome/Edge，False=使用chromedriver-autoinstaller自动下载）
USE_LOCAL_BROWSER = True
# 本地浏览器调试端口
DEBUG_PORT = 9222
# 自定义浏览器路径（留空则自动检测）
CUSTOM_BROWSER_PATH = ""
# ================================================================

# 构建搜索页URL模板（支持关键词搜索）
BASE_URL = f'`https://you.ctrip.com/search/travels/{{}}?keyword={quote(KEYWORD)}`'

# 请求头配置
ua = UserAgent()
HEADERS = {
    'User-Agent': ua.random,
    'Referer': '`https://you.ctrip.com/`',
    'Cookie': '',
}


# ==================== 浏览器管理类（参考MediaCrawler项目） ====================

class BrowserLauncher:
    """浏览器启动器，检测和启动本地Chrome/Edge浏览器"""

    def __init__(self):
        self.system = platform.system()
        self.browser_process = None
        self.debug_port = None

    def detect_browser_paths(self):
        """检测系统中可用的浏览器路径"""
        paths = []

        if self.system == "Windows":
            possible_paths = [
                os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe"),
                os.path.expandvars(r"%PROGRAMFILES(X86)%\Microsoft\Edge\Application\msedge.exe"),
            ]
        elif self.system == "Darwin":  # macOS
            possible_paths = [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            ]
        else:  # Linux
            possible_paths = [
                "/usr/bin/google-chrome",
                "/usr/bin/google-chrome-stable",
                "/usr/bin/chromium-browser",
                "/usr/bin/chromium",
                "/usr/bin/microsoft-edge",
            ]

        for path in possible_paths:
            if os.path.isfile(path) and os.access(path, os.X_OK):
                paths.append(path)

        return paths

    def find_available_port(self, start_port=9222):
        """查找可用端口"""
        import socket
        port = start_port
        while port < start_port + 100:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(('localhost', port))
                    return port
            except OSError:
                port += 1
        raise RuntimeError(f"无法找到可用端口，从 {start_port} 到 {port-1}")

    def launch_browser(self, browser_path, debug_port, headless=False, user_data_dir=None):
        """启动浏览器进程"""
        args = [
            browser_path,
            f"--remote-debugging-port={debug_port}",
            "--remote-debugging-address=0.0.0.0",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-timer-throttling",
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
            "--disable-features=TranslateUI",
            "--disable-ipc-flooding-protection",
            "--disable-hang-monitor",
            "--disable-prompt-on-repost",
            "--disable-sync",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--exclude-switches=enable-automation",
            "--disable-infobars",
        ]

        if headless:
            args.extend(["--headless=new", "--disable-gpu"])
        else:
            args.extend(["--start-maximized"])

        if user_data_dir:
            args.append(f"--user-data-dir={user_data_dir}")

        print(f"[浏览器] 启动中: {browser_path}")
        print(f"[浏览器] 调试端口: {debug_port}")

        if self.system == "Windows":
            process = subprocess.Popen(
                args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
            )
        else:
            process = subprocess.Popen(
                args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                preexec_fn=os.setsid
            )

        self.browser_process = process
        return process

    def wait_for_browser_ready(self, debug_port, timeout=30):
        """等待浏览器就绪"""
        import socket
        print(f"[浏览器] 等待浏览器启动... (端口 {debug_port})")

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(1)
                    if s.connect_ex(('localhost', debug_port)) == 0:
                        print("[浏览器] 浏览器已就绪")
                        return True
            except Exception:
                pass
            time.sleep(0.5)

        print("[浏览器] 浏览器启动超时")
        return False

    def get_browser_info(self, browser_path):
        """获取浏览器信息"""
        try:
            name = "Unknown"
            if "chrome" in browser_path.lower():
                name = "Google Chrome"
            elif "edge" in browser_path.lower() or "msedge" in browser_path.lower():
                name = "Microsoft Edge"
            elif "chromium" in browser_path.lower():
                name = "Chromium"

            result = subprocess.run([browser_path, "--version"],
                                  capture_output=True, text=True, encoding='utf-8', errors='ignore', timeout=5)
            version = result.stdout.strip() if result.stdout else "Unknown"
            return name, version
        except Exception:
            return "Unknown", "Unknown"

    def cleanup(self):
        """清理浏览器进程"""
        if not self.browser_process:
            return

        print("[浏览器] 关闭中...")
        try:
            if self.system == "Windows":
                self.browser_process.terminate()
                self.browser_process.wait(timeout=5)
            else:
                import signal
                os.killpg(os.getpgid(self.browser_process.pid), signal.SIGTERM)
                self.browser_process.wait(timeout=5)
        except Exception as e:
            print(f"[浏览器] 关闭异常: {e}")
        finally:
            self.browser_process = None


def get_browser_cookies_via_playwright():
    """
    使用Playwright通过CDP连接本地浏览器获取Cookie
    """
    try:
        from playwright.sync_api import sync_playwright
        import httpx

        launcher = BrowserLauncher()

        # 获取浏览器路径
        browser_path = None
        if CUSTOM_BROWSER_PATH and os.path.isfile(CUSTOM_BROWSER_PATH):
            browser_path = CUSTOM_BROWSER_PATH
        else:
            browser_paths = launcher.detect_browser_paths()
            if browser_paths:
                browser_path = browser_paths[0]

        if not browser_path:
            print("[错误] 未找到本地浏览器，请安装Chrome或Edge")
            return None

        browser_name, browser_version = launcher.get_browser_info(browser_path)
        print(f"[浏览器] 检测到: {browser_name} ({browser_version})")

        # 启动浏览器
        debug_port = launcher.find_available_port(DEBUG_PORT)
        user_data_dir = os.path.join(os.getcwd(), "browser_data", "ctrip_cookies")
        os.makedirs(user_data_dir, exist_ok=True)

        launcher.launch_browser(browser_path, debug_port, headless=True, user_data_dir=user_data_dir)

        if not launcher.wait_for_browser_ready(debug_port, timeout=30):
            print("[错误] 浏览器启动失败")
            launcher.cleanup()
            return None

        # 通过CDP获取Cookie
        cookies = None
        try:
            # 获取WebSocket URL
            async with httpx.AsyncClient() as client:
                response = await client.get(f"http://localhost:{debug_port}/json/version", timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    ws_url = data.get("webSocketDebuggerUrl")

                    # 使用Playwright连接
                    with sync_playwright() as p:
                        browser = p.chromium.connect_over_cdp(ws_url, timeout=10000)
                        if browser.contexts:
                            context = browser.contexts[0]
                            # 访问携程获取Cookie
                            page = context.new_page()
                            search_url = f'`https://you.ctrip.com/search/travels/?keyword={quote(KEYWORD)}`'
                            page.goto(search_url, timeout=30000)
                            time.sleep(3)
                            cookies = context.cookies()
                            browser.close()

        except Exception as e:
            print(f"[CDP] 连接异常: {e}")

        launcher.cleanup()
        return cookies

    except ImportError as e:
        print(f"[错误] Playwright未安装: {e}")
        print("提示：pip install playwright && playwright install")
        return None
    except Exception as e:
        print(f"[错误] 获取Cookie失败: {e}")
        return None


def get_cookies_flexible():
    """
    灵活的Cookie获取策略
    """
    cookies = None

    # 方式1：使用Playwright + 本地浏览器（通过CDP连接）
    if USE_LOCAL_BROWSER:
        print("[Cookie] 正在使用本地浏览器获取Cookie...")
        cookies = get_browser_cookies_via_playwright()
        if cookies:
            return cookies

    # 方式2：使用chromedriver-autoinstaller自动下载
    print("[Cookie] 正在使用自动下载浏览器获取Cookie...")
    try:
        import chromedriver_autoinstaller
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options

        chromedriver_autoinstaller.install()
        chrome_options = Options()
        chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument(f'--user-agent={ua.random}')

        driver = webdriver.Chrome(options=chrome_options)
        try:
            search_url = f'`https://you.ctrip.com/search/travels/?keyword={quote(KEYWORD)}`'
            driver.get(search_url)
            time.sleep(3)
            cookies = driver.get_cookies()
            print(f"[Selenium] 获取到 {len(cookies)} 个Cookie")
            return {c['name']: c['value'] for c in cookies}
        finally:
            driver.quit()
    except Exception as e:
        print(f"[Selenium] 失败: {e}")

    print("[警告] 无法自动获取Cookie，将使用受限访问")
    return {}


def update_headers_with_cookies(cookies):
    """更新请求头中的Cookie"""
    if cookies:
        if isinstance(cookies, list):
            cookie_str = '; '.join([f"{c['name']}={c['value']}" for c in cookies])
        else:
            cookie_str = '; '.join([f'{k}={v}' for k, v in cookies.items()])
        HEADERS['Cookie'] = cookie_str
        print(f"[Cookie] 已更新 ({len(cookies)} 项)")


def get_page(url, retry=RETRY_TIMES):
    """获取页面HTML，带重试机制"""
    for attempt in range(retry):
        try:
            response = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if response.status_code == 200:
                return response.text
            else:
                print(f"请求失败，状态码: {response.status_code}")
        except RequestException as e:
            print(f"请求异常: {e}")
        time.sleep(3)
    return None


def parse_travel_list(html):
    """解析游记列表页，提取每篇游记的基本信息"""
    selector = etree.HTML(html)
    results = []

    travel_items = selector.xpath('//a[contains(@class, "cursor-pointer")]')
    if not travel_items:
        travel_items = selector.xpath('//div[contains(@class, "travel-item")]//a')
    if not travel_items:
        travel_items = selector.xpath('//a[contains(@href, "/travels/")]')

    for item in travel_items:
        try:
            href = item.xpath('./@href')
            href = href[0] if href else ''

            if not href or '/travels/' not in href:
                continue

            title_elem = item.xpath('.//h2//text()') or item.xpath('.//h3//text()') or item.xpath('.//span//text()')
            title = ''.join([t.strip() for t in title_elem if t.strip()]) if title_elem else ''

            if not title:
                title_match = re.search(r'/travels/[^/]+/(\d+)\.html', href)
                if title_match:
                    title = f"游记-{title_match.group(1)}"

            author_elem = item.xpath('.//span[contains(@class, "author")]//text()') or \
                         item.xpath('.//div[contains(@class, "author")]//text()')
            author = ''.join([a.strip() for a in author_elem if a.strip()]) if author_elem else '匿名'

            summary_elem = item.xpath('.//p[contains(@class, "desc")]//text()') or \
                          item.xpath('.//p[contains(@class, "summary")]//text()') or \
                          item.xpath('.//div[contains(@class, "summary")]//text()')
            summary = ''.join([s.strip() for s in summary_elem if s.strip()]) if summary_elem else ''

            url = href
            if url and not url.startswith('http'):
                url = f'`https://you.ctrip.com{url}`'

            if url and title:
                results.append({
                    'title': title,
                    'url': url,
                    'summary': summary,
                    'author': author
                })
        except Exception as e:
            print(f"解析条目时出错: {e}")
            continue

    seen = set()
    unique_results = []
    for r in results:
        if r['url'] not in seen:
            seen.add(r['url'])
            unique_results.append(r)

    return unique_results


def parse_travel_detail(html, base_data):
    """解析游记详情页"""
    selector = etree.HTML(html)

    content_blocks = selector.xpath('//div[contains(@class, "article-content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//div[contains(@class, "rich_media_content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//div[contains(@class, "travel-content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//article//text()')

    full_content = ''.join(content_blocks).strip() if content_blocks else ''
    full_content = re.sub(r'\s+', ' ', full_content)

    img_urls = []
    img_elements = selector.xpath('//div[contains(@class, "article-content")]//img/@src')
    if not img_elements:
        img_elements = selector.xpath('//div[contains(@class, "rich_media_content")]//img/@src')
    if not img_elements:
        img_elements = selector.xpath('//div[contains(@class, "travel-content")]//img/@src')
    if not img_elements:
        img_elements = selector.xpath('//img[contains(@class, "content-img")]/@src')
    if not img_elements:
        img_elements = selector.xpath('//article//img/@src')

    for img in img_elements:
        if img and not img.startswith('data:'):
            if img.startswith('//'):
                img = f'https:{img}'
            img_urls.append(img)

    meta_text = selector.xpath('//div[contains(@class, "travel-info")]//text()')
    if not meta_text:
        meta_text = selector.xpath('//div[contains(@class, "info-bar")]//text()')
    if not meta_text:
        meta_text = selector.xpath('//div[contains(@class, "meta")]//text()')
    meta_str = ''.join(meta_text) if meta_text else ''

    days_match = re.search(r'(\d+)\s*天', meta_str)
    travel_days = days_match.group(1) if days_match else ''

    cost_match = re.search(r'人均[花费]*[\:：]?\s*(\d+)', meta_str)
    per_capita_cost = cost_match.group(1) if cost_match else ''

    time_match = re.search(r'(\d{4}-\d{1,2}-\d{1,2})', meta_str)
    publish_date = time_match.group(1) if time_match else ''

    base_data.update({
        'content': full_content,
        'image_urls': '; '.join(img_urls[:20]),
        'image_count': len(img_urls),
        'travel_days': travel_days,
        'per_capita_cost': per_capita_cost,
        'publish_date': publish_date
    })

    return base_data


def fetch_and_parse_detail(travel_item):
    """获取详情页并解析"""
    url = travel_item.get('url')
    if not url:
        return travel_item

    print(f"正在爬取: {travel_item.get('title', '未知标题')}")
    html = get_page(url)
    if html:
        travel_item = parse_travel_detail(html, travel_item)
    else:
        travel_item.update({'content': '获取失败', 'image_urls': '', 'image_count': 0})

    time.sleep(0.5)
    return travel_item


def save_to_excel(data_list, filename):
    """保存数据到Excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'{KEYWORD}游记'

    headers = ['标题', '作者', '链接', '摘要', '正文', '出游天数', '人均花费',
               '发布日期', '图片链接', '图片数量']
    ws.append(headers)

    for item in data_list:
        row = [
            item.get('title', ''),
            item.get('author', ''),
            item.get('url', ''),
            item.get('summary', ''),
            item.get('content', ''),
            item.get('travel_days', ''),
            item.get('per_capita_cost', ''),
            item.get('publish_date', ''),
            item.get('image_urls', ''),
            item.get('image_count', 0)
        ]
        ws.append(row)

    wb.save(filename)
    print(f"数据已保存至: {filename}")


def main():
    """主函数"""
    print("=" * 50)
    print("携程游记爬虫启动")
    print(f"关键词: {KEYWORD}")
    print(f"使用本地浏览器: {USE_LOCAL_BROWSER}")
    print("=" * 50)

    # 获取Cookie
    print("\n正在获取Cookie...")
    cookies = get_cookies_flexible()
    update_headers_with_cookies(cookies)

    all_travels = []

    # 遍历列表页
    for page in range(START_PAGE, END_PAGE + 1):
        url = BASE_URL.format(page)
        print(f"\n正在处理第{page}页: {url}")
        html = get_page(url)

        if not html:
            print(f"第{page}页获取失败，跳过")
            continue

        travels = parse_travel_list(html)
        print(f"第{page}页共找到 {len(travels)} 篇游记")
        all_travels.extend(travels)
        time.sleep(1)

    print(f"\n共找到 {len(all_travels)} 篇游记，开始获取详情页...")

    # 多线程获取详情
    pool = Pool(THREAD_POOL_SIZE)
    detailed_travels = pool.map(fetch_and_parse_detail, all_travels)
    pool.close()
    pool.join()

    # 保存到Excel
    save_to_excel(detailed_travels, OUTPUT_FILE)
    print(f"\n爬取完成！共成功获取 {len(detailed_travels)} 篇游记的详细内容")


if __name__ == '__main__':
    main()