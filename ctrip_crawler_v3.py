#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
携程游记爬虫 v3.0 - 浏览器版本
功能：使用浏览器爬取携程游记，绕过反爬机制
依赖：pip install playwright httpx beautifulsoup4 lxml openpyxl

================================================================================
                          【基础配置区 - 请先修改这里】
================================================================================
"""

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  城市配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CITY_NAME = "lanzhou"           # 城市拼音（beijing/shanghai/lanzhou/xiamen等）
CITY_ID = "231"                # 城市ID（可不填，自动从映射表获取）

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  爬取配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DEFAULT_PAGES = 3               # 默认爬取页数
TRAVELS_PER_PAGE = 20          # 每页游记数量

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  输出配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT_FORMAT = "txt"           # 输出格式：txt 或 excel
OUTPUT_DIR = "./output"        # 输出目录

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  浏览器配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BROWSER_PORT = 9222            # 浏览器调试端口
HEADLESS = True                # 是否隐藏浏览器窗口
REQUEST_DELAY = 3              # 操作间隔时间（秒）

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  城市ID映射表
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CITY_ID_MAP = {
    "beijing": "1", "shanghai": "2", "guangzhou": "3", "shenzhen": "4",
    "chengdu": "5", "hangzhou": "6", "wuhan": "7", "xian": "8",
    "nanjing": "9", "chongqing": "10", "tianjin": "11", "suzhou": "12",
    "xiamen": "21", "changsha": "14", "qingdao": "15", "lanzhou": "231",
    "lhasa": "232", "urumqi": "233",
}

"""
================================================================================
                              【代码区 - 无需修改】
================================================================================
"""

import asyncio
import os
import re
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional

import httpx
from bs4 import BeautifulSoup


# ════════════════════════════════════════════════════════════════════════════════
#  URL工具类
# ════════════════════════════════════════════════════════════════════════════════
class URLBuilder:
    """携程URL构造器"""
    
    BASE_URL = "https://you.ctrip.com"
    
    @classmethod
    def get_city_identifier(cls, city_name: str, city_id: str = "") -> str:
        """获取城市标识"""
        if city_id:
            return f"{city_name}{city_id}"
        
        mapped_id = CITY_ID_MAP.get(city_name.lower(), "")
        if mapped_id:
            return f"{city_name}{mapped_id}"
        
        return city_name
    
    @classmethod
    def get_list_url(cls, city_name: str, city_id: str = "", page: int = 1) -> str:
        """获取游记列表URL"""
        identifier = cls.get_city_identifier(city_name, city_id)
        if page == 1:
            return f"{cls.BASE_URL}/travels/{identifier}.html"
        else:
            return f"{cls.BASE_URL}/travels/{identifier}-p{page}.html"


# ════════════════════════════════════════════════════════════════════════════════
#  浏览器管理器
# ════════════════════════════════════════════════════════════════════════════════
class BrowserManager:
    """浏览器管理器 - 使用 Playwright CDP 模式"""
    
    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
    
    async def init(self, port: int = BROWSER_PORT, headless: bool = HEADLESS):
        """初始化浏览器"""
        from playwright.async_api import async_playwright
        
        self.playwright = await async_playwright().start()
        
        # 尝试连接已有浏览器或启动新浏览器
        try:
            # 尝试连接
            async with httpx.AsyncClient() as client:
                resp = await client.get(f"http://localhost:{port}/json/version", timeout=5)
                ws_url = resp.json().get("webSocketDebuggerUrl")
            
            print(f"  [浏览器] 连接到已有浏览器 (端口 {port})")
            self.browser = await self.playwright.chromium.connect_over_cdp(ws_url)
            
        except Exception:
            # 启动新浏览器
            print(f"  [浏览器] 启动新浏览器 (端口 {port})")
            import platform
            system = platform.system()
            
            if system == "Windows":
                browser_path = os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe")
                if not os.path.exists(browser_path):
                    browser_path = os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe")
            elif system == "Darwin":
                browser_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            else:
                browser_path = "/usr/bin/google-chrome"
            
            if not os.path.exists(browser_path):
                print(f"  [错误] 未找到浏览器: {browser_path}")
                print(f"  [提示] 请安装 Chrome 或使用 --connect-existing 连接已有浏览器")
                sys.exit(1)
            
            self.browser = await self.playwright.chromium.launch(
                headless=headless,
                executable_path=browser_path,
                args=[
                    f"--remote-debugging-port={port}",
                    "--no-first-run",
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                ]
            )
        
        # 创建上下文和页面
        self.context = await self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        self.page = await self.context.new_page()
        print("  [浏览器] 初始化完成")
    
    async def goto(self, url: str, wait_time: float = 3):
        """访问页面"""
        await self.page.goto(url, timeout=60000)
        await asyncio.sleep(wait_time)
    
    async def get_cookies(self):
        """获取Cookie"""
        return await self.context.cookies()
    
    async def close(self):
        """关闭浏览器"""
        if self.page:
            await self.page.close()
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()


# ════════════════════════════════════════════════════════════════════════════════
#  解析类
# ════════════════════════════════════════════════════════════════════════════════
class Parser:
    """数据解析器"""
    
    @staticmethod
    async def parse_list_page(page, city_name: str, city_id: str) -> List[Dict]:
        """解析游记列表页（从浏览器页面）"""
        travels = []
        
        # 使用 JavaScript 提取数据
        items = await page.evaluate("""
            () => {
                // 查找所有游记卡片
                const selectors = [
                    'div.travel_item',
                    'div.journal-item', 
                    'article.journal-item',
                    '.travel-item',
                    '.journal-item',
                    'li[data-id]'
                ];
                
                let cards = [];
                for (const sel of selectors) {
                    cards = document.querySelectorAll(sel);
                    if (cards.length > 0) break;
                }
                
                // 如果没找到，查找包含 /travels/城市ID/游记ID.html 链接的元素
                if (cards.length === 0) {
                    const links = document.querySelectorAll('a[href*="/travels/"]');
                    const seen = new Set();
                    cards = [];
                    links.forEach(link => {
                        const href = link.getAttribute('href');
                        if (href && /\\/travels\\/\\w+\\d+\\/\\d+\\.html/.test(href)) {
                            const parent = link.closest('div, li, article');
                            if (parent && !seen.has(parent)) {
                                seen.add(parent);
                                cards.push(parent);
                            }
                        }
                    });
                }
                
                return Array.from(cards).slice(0, 30).map(card => {
                    // 提取链接
                    const link = card.querySelector('a[href*="/travels/"]') || card.querySelector('a');
                    let href = link ? link.getAttribute('href') : '';
                    let travelId = '';
                    
                    const match = href.match(/\\/travels\\/\\w+\\/(\\d+)\\.html/);
                    if (match) travelId = match[1];
                    
                    return {
                        title: link ? link.textContent.trim() : '',
                        href: href,
                        travelId: travelId,
                        author: card.querySelector('.author, .user-name, [class*="author"], .name')?.textContent?.trim() || '',
                        date: card.querySelector('.date, .time, [class*="time"]')?.textContent?.trim() || '',
                        views: card.querySelector('.view, .views, [class*="view"]')?.textContent?.trim() || ''
                    };
                }).filter(item => item.title && item.travelId);
            }
        """)
        
        # 转换格式
        for item in items:
            travels.append({
                "title": item["title"],
                "travel_id": item["travelId"],
                "detail_url": f"https://you.ctrip.com{item['href']}" if item['href'].startswith('/') else item['href'],
                "author": item["author"] or "未知",
                "date": item["date"],
                "view_count": item["views"],
            })
        
        return travels
    
    @staticmethod
    async def parse_detail_page(page) -> Dict:
        """解析游记详情页"""
        detail = await page.evaluate("""
            () => {
                // 提取标题
                const title = document.querySelector('h1.title, h1, .article-title, [class*="title"]')?.textContent?.trim() || '';
                
                // 提取作者
                const author = document.querySelector('.author-name, .author, [class*="author"], a[href*="/user/"]')?.textContent?.trim() || '未知';
                
                // 提取日期
                const date = document.querySelector('.publish-time, .publishTime, time, [class*="publish"]')?.textContent?.trim() || '';
                
                // 提取正文内容
                const contentSelectors = [
                    'div.article-content',
                    'div.article_content', 
                    'div.content',
                    '#article-content',
                    'div.detail-content',
                    '[class*="content"]'
                ];
                
                let content = '';
                for (const sel of contentSelectors) {
                    const elem = document.querySelector(sel);
                    if (elem) {
                        // 获取所有文字节点
                        const textNodes = [];
                        const walker = document.createTreeWalker(elem, NodeFilter.SHOW_TEXT, null, false);
                        let node;
                        while (node = walker.nextNode()) {
                            const text = node.textContent.trim();
                            if (text) textNodes.push(text);
                        }
                        content = textNodes.join('\\n');
                        break;
                    }
                }
                
                // 如果没找到，使用 body
                if (!content) {
                    const body = document.body.cloneNode(true);
                    ['script', 'style', 'nav', 'header', 'footer', 'aside'].forEach(tag => {
                        body.querySelectorAll(tag).forEach(el => el.remove());
                    });
                    const textNodes = [];
                    const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT, null, false);
                    let node;
                    while (node = walker.nextNode()) {
                        const text = node.textContent.trim();
                        if (text && text.length > 10) textNodes.push(text);
                    }
                    content = textNodes.slice(0, 200).join('\\n');
                }
                
                // 提取图片
                const images = Array.from(document.querySelectorAll('img'))
                    .map(img => img.src || img.getAttribute('data-src'))
                    .filter(src => src && src.startsWith('http'))
                    .slice(0, 20);
                
                // 提取标签
                const tags = Array.from(document.querySelectorAll('.tag, .tags a, [class*="tag"] a'))
                    .map(tag => tag.textContent.trim())
                    .filter(tag => tag);
                
                return { title, author, date, content, images, tags };
            }
        """)
        
        return detail


# ════════════════════════════════════════════════════════════════════════════════
#  数据存储类
# ════════════════════════════════════════════════════════════════════════════════
class DataStorage:
    """数据存储器"""
    
    def __init__(self):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    def save(self, travels: List[Dict], output_format: str = ""):
        if not output_format:
            output_format = OUTPUT_FORMAT
        
        if output_format.lower() == "txt":
            self._save_to_txt(travels)
        elif output_format.lower() == "excel":
            self._save_to_excel(travels)
    
    def _get_filename(self, extension: str) -> str:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ctrip_travels_{CITY_NAME}_{timestamp}.{extension}"
        return os.path.join(OUTPUT_DIR, filename)
    
    def _save_to_txt(self, travels: List[Dict]):
        filename = self._get_filename("txt")
        
        with open(filename, "w", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write(f"携程游记爬取结果\n")
            f.write(f"城市: {CITY_NAME}\n")
            f.write(f"数量: {len(travels)} 篇\n")
            f.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            for idx, travel in enumerate(travels, 1):
                f.write(f"\n{'─' * 80}\n")
                f.write(f"【第 {idx} 篇】\n")
                f.write(f"{'─' * 80}\n")
                f.write(f"标题: {travel.get('title', '')}\n")
                f.write(f"作者: {travel.get('author', '未知')}\n")
                f.write(f"日期: {travel.get('date', '')}\n")
                f.write(f"阅读: {travel.get('view_count', '')}\n")
                f.write(f"链接: {travel.get('detail_url', '')}\n")
                
                tags = travel.get("tags", [])
                if tags:
                    f.write(f"标签: {', '.join(tags)}\n")
                
                f.write(f"\n【正文内容】\n")
                content = travel.get("content", "")
                if content:
                    f.write(content)
                else:
                    f.write("（未获取到详细内容）")
                f.write("\n")
        
        print(f"\n✅ 数据已保存到: {filename}")
    
    def _save_to_excel(self, travels: List[Dict]):
        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, PatternFill
        except ImportError:
            print("[错误] 请安装: pip install openpyxl")
            self._save_to_txt(travels)
            return
        
        filename = self._get_filename("xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "携程游记"
        
        headers = ["序号", "标题", "作者", "日期", "阅读数", "标签", "正文内容", "详情链接"]
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for row, travel in enumerate(travels, 2):
            ws.cell(row=row, column=1, value=row - 1)
            ws.cell(row=row, column=2, value=travel.get("title", ""))
            ws.cell(row=row, column=3, value=travel.get("author", "未知"))
            ws.cell(row=row, column=4, value=travel.get("date", ""))
            ws.cell(row=row, column=5, value=travel.get("view_count", ""))
            ws.cell(row=row, column=6, value=", ".join(travel.get("tags", [])))
            
            content = travel.get("content", "")[:30000]
            ws.cell(row=row, column=7, value=content)
            ws.cell(row=row, column=8, value=travel.get("detail_url", ""))
        
        wb.save(filename)
        print(f"\n✅ 数据已保存到: {filename}")


# ════════════════════════════════════════════════════════════════════════════════
#  爬虫主类
# ════════════════════════════════════════════════════════════════════════════════
class CtripTravelCrawler:
    """携程游记爬虫主类"""
    
    def __init__(self):
        self.browser = BrowserManager()
        self.storage = DataStorage()
    
    async def crawl_list_page(self, page_num: int) -> List[Dict]:
        """爬取单页游记列表"""
        url = URLBuilder.get_list_url(CITY_NAME, CITY_ID, page_num)
        print(f"\n[第 {page_num} 页] 访问: {url}")
        
        await self.browser.goto(url, wait_time=3)
        
        travels = await Parser.parse_list_page(self.browser.page, CITY_NAME, CITY_ID)
        print(f"    [完成] 解析到 {len(travels)} 篇游记")
        
        return travels
    
    async def crawl_detail(self, travel: Dict, idx: int, total: int) -> Dict:
        """爬取单篇游记详情"""
        title = travel.get("title", "")[:30]
        print(f"    [{idx}/{total}] {title}...")
        
        await self.browser.goto(travel["detail_url"], wait_time=2)
        
        detail = await Parser.parse_detail_page(self.browser.page)
        travel.update(detail)
        
        return travel
    
    async def crawl(self) -> List[Dict]:
        """执行爬取任务"""
        print("=" * 70)
        print("  携程游记爬虫 v3.0  -  浏览器版本")
        print("=" * 70)
        print(f"  城市: {CITY_NAME} ({CITY_ID})")
        print(f"  页数: {DEFAULT_PAGES}")
        print(f"  输出: {OUTPUT_FORMAT}")
        print("=" * 70)
        
        # 初始化浏览器
        print("\n[启动] 初始化浏览器...")
        await self.browser.init(port=BROWSER_PORT, headless=HEADLESS)
        
        all_travels = []
        
        # 阶段1：爬取列表页
        print("\n【阶段1】正在爬取游记列表...")
        for page in range(1, DEFAULT_PAGES + 1):
            travels = await self.crawl_list_page(page)
            all_travels.extend(travels)
            
            if page < DEFAULT_PAGES:
                print(f"  [等待] {REQUEST_DELAY}秒后继续...")
                await asyncio.sleep(REQUEST_DELAY)
        
        print(f"\n[汇总] 共获取 {len(all_travels)} 篇游记")
        
        if not all_travels:
            print("[错误] 未获取到游记数据")
            return []
        
        # 阶段2：爬取详情
        print("\n【阶段2】正在爬取游记详情...")
        for idx, travel in enumerate(all_travels, 1):
            await self.crawl_detail(travel, idx, len(all_travels))
            await asyncio.sleep(REQUEST_DELAY)
        
        return all_travels
    
    async def run(self):
        """运行爬虫"""
        try:
            travels = await self.crawl()
            
            if travels:
                self.storage.save(travels)
                print(f"\n🎉 爬取完成！共获取 {len(travels)} 篇游记")
            else:
                print("\n❌ 未获取到游记数据")
                
        except KeyboardInterrupt:
            print("\n\n[中断] 用户取消")
        except Exception as e:
            print(f"\n[错误] {e}")
            import traceback
            traceback.print_exc()
        finally:
            await self.browser.close()
            print("[完成] 浏览器已关闭")


# ════════════════════════════════════════════════════════════════════════════════
#  程序入口
# ════════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    crawler = CtripTravelCrawler()
    asyncio.run(crawler.run())
