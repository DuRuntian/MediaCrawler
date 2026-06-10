#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
携程游记爬虫 - 无需登录版本 v2.0
功能：爬取携程游记列表及详情，采集文字内容
依赖：pip install httpx beautifulsoup4 lxml openpyxl

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
TRAVELS_PER_PAGE = 20          # 每页游记数量（携程通常每页20篇）

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  输出配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT_FORMAT = "txt"           # 输出格式：txt 或 excel
OUTPUT_DIR = "./output"        # 输出目录

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  反爬配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REQUEST_DELAY = 2               # 请求间隔时间（秒）
REQUEST_TIMEOUT = 30           # 请求超时时间（秒）
MAX_RETRIES = 3                # 最大重试次数

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
        """获取城市标识（如 lanzhou231）"""
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
    
    @classmethod
    def get_detail_url(cls, travel_id: str, city_name: str, city_id: str = "") -> str:
        """获取游记详情URL"""
        identifier = cls.get_city_identifier(city_name, city_id)
        return f"{cls.BASE_URL}/travels/{identifier}/{travel_id}.html"


# ════════════════════════════════════════════════════════════════════════════════
#  HTTP请求类
# ════════════════════════════════════════════════════════════════════════════════
class HTTPClient:
    """HTTP客户端"""
    
    def __init__(self):
        self.headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Referer": "https://you.ctrip.com/",
        }
    
    async def get(self, url: str, retries: int = 0) -> Optional[str]:
        """发送GET请求"""
        try:
            async with httpx.AsyncClient(
                timeout=httpx.Timeout(REQUEST_TIMEOUT),
                follow_redirects=True,
                headers=self.headers
            ) as client:
                response = await client.get(url)
                
                if response.status_code == 200:
                    return response.text
                elif response.status_code in [403, 432, 503]:
                    if retries < MAX_RETRIES:
                        print(f"    [警告] 状态码 {response.status_code}，{REQUEST_DELAY * (retries + 1)}秒后重试...")
                        await asyncio.sleep(REQUEST_DELAY * (retries + 1))
                        return await self.get(url, retries + 1)
                    else:
                        print(f"    [错误] 达到最大重试次数")
                        return None
                else:
                    print(f"    [错误] 状态码 {response.status_code}")
                    return None
                    
        except Exception as e:
            print(f"    [错误] 请求异常: {e}")
            if retries < MAX_RETRIES:
                await asyncio.sleep(REQUEST_DELAY)
                return await self.get(url, retries + 1)
            return None


# ════════════════════════════════════════════════════════════════════════════════
#  解析类
# ════════════════════════════════════════════════════════════════════════════════
class Parser:
    """数据解析器"""
    
    @staticmethod
    def parse_list_page(html: str, city_name: str, city_id: str) -> List[Dict]:
        """解析游记列表页"""
        travels = []
        soup = BeautifulSoup(html, "lxml")
        
        # 尝试多种选择器
        selectors = [
            "div.travel_item", "div.journal-item", "article.journal-item",
            "div[class*='travel']", "div[class*='journal']",
            "li.travel-item", "div.list-item",
        ]
        
        items = []
        matched_selector = None
        for selector in selectors:
            items = soup.select(selector)
            if items:
                matched_selector = selector
                print(f"    [解析] 选择器: {selector}，匹配 {len(items)} 个元素")
                break
        
        # 兜底：查找包含travels链接的元素
        if not items:
            links = soup.find_all("a", href=lambda h: h and "/travels/" in h)
            seen = set()
            for link in links:
                parent = link.parent
                if parent and id(parent) not in seen:
                    seen.add(id(parent))
                    items.append(parent)
            if items:
                print(f"    [解析] 兜底匹配，找到 {len(items)} 个元素")
        
        for item in items:
            try:
                travel = Parser._extract_travel_info(item, city_name, city_id)
                if travel and travel.get("title") and travel.get("travel_id"):
                    travels.append(travel)
            except Exception:
                continue
        
        return travels
    
    @staticmethod
    def _extract_travel_info(element, city_name: str, city_id: str) -> Optional[Dict]:
        """从元素中提取游记信息"""
        travel = {}
        
        # 提取标题和链接
        link_elem = element.select_one("a[href*='/travels/']")
        if not link_elem:
            link_elem = element.find("a")
        
        if link_elem:
            href = link_elem.get("href", "")
            travel["title"] = link_elem.get_text(strip=True)
            
            # 从href中提取游记ID
            match = re.search(r'/travels/\w+/\d+\.html', href)
            if match:
                parts = href.split("/")
                if len(parts) >= 5:
                    travel["travel_id"] = parts[-1].replace(".html", "")
                    travel["detail_url"] = URLBuilder.BASE_URL + href
            else:
                return None
        else:
            return None
        
        # 提取作者
        author_elem = element.select_one(
            ".author, .user-name, .username, [class*='author'], span.name"
        )
        travel["author"] = author_elem.get_text(strip=True) if author_elem else "未知"
        
        # 提取日期
        date_elem = element.select_one(".date, .time, [class*='time'], span.date")
        travel["date"] = date_elem.get_text(strip=True) if date_elem else ""
        
        # 提取阅读数
        view_elem = element.select_one(".view, .views, [class*='view'], span.view")
        travel["view_count"] = view_elem.get_text(strip=True) if view_elem else ""
        
        return travel
    
    @staticmethod
    def parse_detail_page(html: str, travel_id: str) -> Dict:
        """解析游记详情页"""
        soup = BeautifulSoup(html, "lxml")
        detail = {
            "travel_id": travel_id,
            "content": "",
            "images": [],
            "tags": [],
        }
        
        # 提取标题
        title_elem = soup.select_one("h1.title, h1, .article-title, [class*='title']")
        if title_elem:
            detail["title"] = title_elem.get_text(strip=True)
        else:
            title_elem = soup.select_one("meta[property='og:title']")
            detail["title"] = title_elem.get("content", "") if title_elem else ""
        
        # 提取作者
        author_elem = soup.select_one(
            ".author-name, .author, [class*='author'], a[href*='/user/']"
        )
        detail["author"] = author_elem.get_text(strip=True) if author_elem else "未知"
        
        # 提取发布日期
        date_elem = soup.select_one(
            ".publish-time, .publishTime, [class*='publish'], span.date, time"
        )
        detail["publish_date"] = date_elem.get_text(strip=True) if date_elem else ""
        
        # 提取正文内容（核心！）
        content_selectors = [
            "div.article-content", "div.article_content", "div.content",
            "div[class*='content']", "#article-content", "div.detail-content",
        ]
        
        content_elem = None
        for selector in content_selectors:
            content_elem = soup.select_one(selector)
            if content_elem:
                break
        
        if content_elem:
            # 提取纯文本内容
            content_text = content_elem.get_text(separator="\n", strip=True)
            lines = [line.strip() for line in content_text.split("\n") if line.strip()]
            detail["content"] = "\n".join(lines)
            
            # 提取图片链接
            for img in content_elem.find_all("img"):
                src = img.get("src") or img.get("data-src") or ""
                if src and src.startswith("http"):
                    detail["images"].append(src)
        else:
            # 兜底：获取body内容
            body = soup.find("body")
            if body:
                for tag in body.find_all(["script", "style", "nav", "header", "footer"]):
                    tag.decompose()
                content_text = body.get_text(separator="\n", strip=True)
                lines = [line.strip() for line in content_text.split("\n") if line.strip()]
                detail["content"] = "\n".join(lines[:500])
        
        # 提取标签
        tag_elems = soup.select(".tag, .tags a, [class*='tag'] a, .keywords a")
        detail["tags"] = [tag.get_text(strip=True) for tag in tag_elems if tag.get_text(strip=True)]
        
        return detail


# ════════════════════════════════════════════════════════════════════════════════
#  数据存储类
# ════════════════════════════════════════════════════════════════════════════════
class DataStorage:
    """数据存储器"""
    
    def __init__(self):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    def save(self, travels: List[Dict], output_format: str = ""):
        """保存数据"""
        if not output_format:
            output_format = OUTPUT_FORMAT
        
        if output_format.lower() == "txt":
            self._save_to_txt(travels)
        elif output_format.lower() == "excel":
            self._save_to_excel(travels)
        else:
            print(f"[错误] 不支持的输出格式: {output_format}")
    
    def _get_filename(self, extension: str) -> str:
        """生成文件名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ctrip_travels_{CITY_NAME}_{timestamp}.{extension}"
        return os.path.join(OUTPUT_DIR, filename)
    
    def _save_to_txt(self, travels: List[Dict]):
        """保存为TXT格式"""
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
                f.write(f"日期: {travel.get('date', '') or travel.get('publish_date', '')}\n")
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
        return filename
    
    def _save_to_excel(self, travels: List[Dict]):
        """保存为Excel格式"""
        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, PatternFill
        except ImportError:
            print("[错误] 请先安装 openpyxl: pip install openpyxl")
            print("[提示] 降级为TXT格式保存...")
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
        
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 20
        ws.column_dimensions['G'].width = 50
        ws.column_dimensions['H'].width = 40
        
        for row, travel in enumerate(travels, 2):
            ws.cell(row=row, column=1, value=row - 1)
            ws.cell(row=row, column=2, value=travel.get("title", ""))
            ws.cell(row=row, column=3, value=travel.get("author", "未知"))
            ws.cell(row=row, column=4, value=travel.get("date", "") or travel.get("publish_date", ""))
            ws.cell(row=row, column=5, value=travel.get("view_count", ""))
            ws.cell(row=row, column=6, value=", ".join(travel.get("tags", [])))
            
            content = travel.get("content", "")
            if len(content) > 30000:
                content = content[:30000] + "..."
            ws.cell(row=row, column=7, value=content)
            ws.cell(row=row, column=8, value=travel.get("detail_url", ""))
        
        wb.save(filename)
        print(f"\n✅ 数据已保存到: {filename}")
        return filename


# ════════════════════════════════════════════════════════════════════════════════
#  爬虫主类
# ════════════════════════════════════════════════════════════════════════════════
class CtripTravelCrawler:
    """携程游记爬虫主类"""
    
    def __init__(self):
        self.http_client = HTTPClient()
        self.storage = DataStorage()
    
    async def crawl_list_page(self, page: int) -> List[Dict]:
        """爬取单页游记列表"""
        url = URLBuilder.get_list_url(CITY_NAME, CITY_ID, page)
        print(f"\n[第 {page} 页] 请求: {url}")
        
        html = await self.http_client.get(url)
        if not html:
            print(f"    [错误] 获取页面失败")
            return []
        
        travels = Parser.parse_list_page(html, CITY_NAME, CITY_ID)
        print(f"    [完成] 解析到 {len(travels)} 篇游记")
        
        return travels
    
    async def crawl_detail(self, travel: Dict) -> Dict:
        """爬取单篇游记详情"""
        url = travel.get("detail_url", "")
        if not url:
            return travel
        
        travel_id = travel.get("travel_id", "")
        title = travel.get("title", "")[:30]
        print(f"    -> {title}...")
        
        html = await self.http_client.get(url)
        if not html:
            print(f"      [警告] 获取详情失败")
            return travel
        
        detail = Parser.parse_detail_page(html, travel_id)
        travel.update(detail)
        
        return travel
    
    async def crawl(self) -> List[Dict]:
        """执行爬取任务"""
        print("=" * 70)
        print(f"  携程游记爬虫 v2.0  -  无需登录版本")
        print("=" * 70)
        print(f"  城市: {CITY_NAME} ({CITY_ID})")
        print(f"  页数: {DEFAULT_PAGES}")
        print(f"  输出: {OUTPUT_FORMAT}")
        print("=" * 70)
        
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
            print(f"  [{idx}/{len(all_travels)}] ", end="")
            await self.crawl_detail(travel)
            await asyncio.sleep(REQUEST_DELAY)
        
        return all_travels
    
    def run(self):
        """运行爬虫"""
        try:
            travels = asyncio.run(self.crawl())
            
            if travels:
                self.storage.save(travels)
                print(f"\n🎉 爬取完成！共获取 {len(travels)} 篇游记")
            else:
                print("\n❌ 未获取到游记数据")
                
        except KeyboardInterrupt:
            print("\n\n[中断] 用户取消")
            sys.exit(0)
        except Exception as e:
            print(f"\n[错误] {e}")
            import traceback
            traceback.print_exc()


# ════════════════════════════════════════════════════════════════════════════════
#  程序入口
# ════════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    crawler = CtripTravelCrawler()
    crawler.run()
