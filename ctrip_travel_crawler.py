#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
携程游记爬虫
功能：爬取携程you.ctrip.com上的游记列表及详情页内容
输出：Excel文件（标题、作者、出游天数、人均花费、正文、图片链接等）
支持自动获取Cookie和关键词搜索
"""

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
# ================================================================

# 构建搜索页URL模板（支持关键词搜索）
BASE_URL = f'`https://you.ctrip.com/search/travels/{{}}?keyword={quote(KEYWORD)}`'

# 请求头配置
ua = UserAgent()
HEADERS = {
    'User-Agent': ua.random,
    'Referer': '`https://you.ctrip.com/`',
    'Cookie': '',  # 自动获取，无需手动填写
}


def get_cookie_from_login():
    """
    通过登录接口获取Cookie
    注意：这种方式获取的Cookie可能权限有限，建议使用Selenium方式
    """
    try:
        # 尝试访问携程登录页面获取初始Cookie
        session = requests.Session()
        session.headers.update(HEADERS)
        # 访问首页获取基础Cookie
        session.get('`https://www.ctrip.com/`', timeout=10)
        return session.cookies.get_dict()
    except Exception as e:
        print(f"获取Cookie失败: {e}")
        return {}


def get_cookies_with_selenium():
    """
    使用Selenium自动登录获取Cookie（需要安装selenium和chromedriver）
    这种方式可以绕过登录限制，获取完整Cookie
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        import chromedriver_autoinstaller

        # 自动安装chromedriver
        chromedriver_autoinstaller.install()

        # 配置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 无头模式
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        # 设置User-Agent
        chrome_options.add_argument(f'user-agent={ua.random}')

        driver = webdriver.Chrome(options=chrome_options)

        try:
            # 访问携程游记搜索页
            search_url = f'`https://you.ctrip.com/search/travels/?keyword={quote(KEYWORD)}`'
            driver.get(search_url)

            # 等待页面加载
            time.sleep(3)

            # 获取Cookie
            cookies = driver.get_cookies()
            cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

            print(f"Selenium成功获取 {len(cookies)} 个Cookie")
            return cookie_dict

        finally:
            driver.quit()

    except ImportError as e:
        print(f"Selenium未安装或配置不正确: {e}")
        print("提示：安装命令 - pip install selenium chromedriver-autoinstaller")
        return None
    except Exception as e:
        print(f"Selenium获取Cookie失败: {e}")
        return None


def get_cookies_with_playwright():
    """
    使用Playwright自动登录获取Cookie（需要安装playwright和浏览器）
    另一种无头浏览器方案，比Selenium更现代
    """
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent=ua.random,
                viewport={'width': 1920, 'height': 1080}
            )
            page = context.new_page()

            try:
                # 访问搜索页
                search_url = f'`https://you.ctrip.com/search/travels/?keyword={quote(KEYWORD)}`'
                page.goto(search_url, timeout=30000)
                time.sleep(3)

                # 获取Cookie
                cookies = context.cookies()
                cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

                print(f"Playwright成功获取 {len(cookies)} 个Cookie")
                return cookie_dict

            finally:
                browser.close()

    except ImportError as e:
        print(f"Playwright未安装或配置不正确: {e}")
        print("提示：安装命令 - pip install playwright && playwright install")
        return None
    except Exception as e:
        print(f"Playwright获取Cookie失败: {e}")
        return None


def get_cookies_flexible():
    """
    灵活的Cookie获取策略，依次尝试多种方式
    """
    cookies = None

    # 方式1：尝试Playwright（推荐，成功率最高）
    print("正在尝试使用Playwright获取Cookie...")
    cookies = get_cookies_with_playwright()
    if cookies:
        return cookies

    # 方式2：尝试Selenium
    print("正在尝试使用Selenium获取Cookie...")
    cookies = get_cookies_with_selenium()
    if cookies:
        return cookies

    # 方式3：使用基础Session
    print("正在尝试基础方式获取Cookie...")
    cookies = get_cookie_from_login()
    if cookies:
        return cookies

    print("警告：无法自动获取Cookie，将使用受限的匿名访问")
    return {}


def update_headers_with_cookies(cookies):
    """更新请求头中的Cookie"""
    if cookies:
        cookie_str = '; '.join([f'{k}={v}' for k, v in cookies.items()])
        HEADERS['Cookie'] = cookie_str
        print("Cookie已更新到请求头")


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
        time.sleep(3)  # 重试前等待
    return None


def parse_travel_list(html):
    """
    解析游记列表页，提取每篇游记的基本信息
    """
    selector = etree.HTML(html)
    results = []

    # 通过XPath提取游记卡片
    travel_items = selector.xpath('//a[contains(@class, "cursor-pointer")]')

    # 备选方案
    if not travel_items:
        travel_items = selector.xpath('//div[contains(@class, "travel-item")]//a')
    if not travel_items:
        travel_items = selector.xpath('//a[contains(@href, "/travels/")]')

    for item in travel_items:
        try:
            href = item.xpath('./@href')
            href = href[0] if href else ''

            # 过滤非游记详情页的链接
            if not href or '/travels/' not in href:
                continue

            # 提取标题
            title_elem = item.xpath('.//h2//text()') or item.xpath('.//h3//text()') or item.xpath('.//span//text()')
            title = ''.join([t.strip() for t in title_elem if t.strip()]) if title_elem else ''

            if not title:
                # 尝试从链接本身获取标题
                title_match = re.search(r'/travels/[^/]+/(\d+)\.html', href)
                if title_match:
                    title = f"游记-{title_match.group(1)}"

            # 提取作者
            author_elem = item.xpath('.//span[contains(@class, "author")]//text()') or \
                         item.xpath('.//div[contains(@class, "author")]//text()')
            author = ''.join([a.strip() for a in author_elem if a.strip()]) if author_elem else '匿名'

            # 提取摘要
            summary_elem = item.xpath('.//p[contains(@class, "desc")]//text()') or \
                          item.xpath('.//p[contains(@class, "summary")]//text()') or \
                          item.xpath('.//div[contains(@class, "summary")]//text()')
            summary = ''.join([s.strip() for s in summary_elem if s.strip()]) if summary_elem else ''

            # 处理URL
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

    # 去重
    seen = set()
    unique_results = []
    for r in results:
        if r['url'] not in seen:
            seen.add(r['url'])
            unique_results.append(r)

    return unique_results


def parse_travel_detail(html, base_data):
    """
    解析游记详情页，提取正文、图片、游记信息等
    """
    selector = etree.HTML(html)

    # 提取正文（多种匹配方式）
    content_blocks = selector.xpath('//div[contains(@class, "article-content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//div[contains(@class, "rich_media_content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//div[contains(@class, "travel-content")]//text()')
    if not content_blocks:
        content_blocks = selector.xpath('//article//text()')

    full_content = ''.join(content_blocks).strip() if content_blocks else ''
    # 清理多余空白
    full_content = re.sub(r'\s+', ' ', full_content)

    # 提取图片链接
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

    # 提取游记信息字段
    meta_text = selector.xpath('//div[contains(@class, "travel-info")]//text()')
    if not meta_text:
        meta_text = selector.xpath('//div[contains(@class, "info-bar")]//text()')
    if not meta_text:
        meta_text = selector.xpath('//div[contains(@class, "meta")]//text()')
    meta_str = ''.join(meta_text) if meta_text else ''

    # 正则提取出游天数
    days_match = re.search(r'(\d+)\s*天', meta_str)
    travel_days = days_match.group(1) if days_match else ''

    # 正则提取人均花费
    cost_match = re.search(r'人均[花费]*[\:：]?\s*(\d+)', meta_str)
    per_capita_cost = cost_match.group(1) if cost_match else ''

    # 提取发布时间
    time_match = re.search(r'(\d{4}-\d{1,2}-\d{1,2})', meta_str)
    publish_date = time_match.group(1) if time_match else ''

    # 更新数据
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

    # 表头
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
    # 第一步：自动获取Cookie
    print("=" * 50)
    print("携程游记爬虫启动")
    print(f"关键词: {KEYWORD}")
    print("=" * 50)

    print("\n正在获取Cookie...")
    cookies = get_cookies_flexible()
    update_headers_with_cookies(cookies)

    all_travels = []

    # 第二步：遍历列表页，获取所有游记概要
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

    # 第三步：多线程获取详情页内容
    pool = Pool(THREAD_POOL_SIZE)
    detailed_travels = pool.map(fetch_and_parse_detail, all_travels)
    pool.close()
    pool.join()

    # 第四步：保存到Excel
    save_to_excel(detailed_travels, OUTPUT_FILE)
    print(f"\n爬取完成！共成功获取 {len(detailed_travels)} 篇游记的详细内容")


if __name__ == '__main__':
    main()