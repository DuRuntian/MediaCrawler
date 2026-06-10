#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
携程游记爬虫 v5.0 - 调用系统 Chrome 浏览器版
==============================================

功能：
  1. 自动检测并启动系统 Chrome 浏览器（带远程调试端口）
  2. 通过 CDP 协议连接浏览器
  3. 爬取携程游记列表及详情
  4. 采集游记正文文字内容
  5. 保存为 TXT 或 Excel 文件

依赖：
  pip install playwright httpx beautifulsoup4 lxml openpyxl
  (注意：本版本使用系统 Chrome，不需要 playwright install chromium)

使用：
  python ctrip_travel_crawler.py
  (在代码开头的配置区修改城市和页数)
"""

# ============================================================
#  配置区（请在这里修改参数）
# ============================================================
CITY_NAME = "lanzhou"           # 城市拼音，如 beijing/shanghai/lanzhou/xiamen
CITY_ID = "231"                 # 城市数字ID，可留空从映射表自动获取
DEFAULT_PAGES = 3               # 爬取页数
OUTPUT_FORMAT = "txt"           # 输出格式：txt 或 excel
OUTPUT_DIR = "./output"         # 输出目录
CHROME_DEBUG_PORT = 9222        # Chrome 调试端口（默认即可）
REQUEST_DELAY = 2               # 操作间隔秒数（反爬友好）


# ============================================================
#  城市 ID 映射表（补充自己需要的城市）
# ============================================================
CITY_ID_MAP = {
    "beijing": "1", "shanghai": "2", "guangzhou": "3", "shenzhen": "4",
    "chengdu": "5", "hangzhou": "6", "wuhan": "7", "xian": "8",
    "nanjing": "9", "chongqing": "10", "tianjin": "11", "suzhou": "12",
    "xiamen": "21", "changsha": "14", "qingdao": "15", "lanzhou": "231",
    "lhasa": "232", "urumqi": "233", "dali": "35", "sanya": "88",
}


# ============================================================
#  下面是代码主体，一般不需要修改
# ============================================================
import asyncio
import os
import platform
import socket
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional

import httpx


# ------------------------------------------------------------
#  1. Chrome 浏览器启动与连接
# ------------------------------------------------------------
def find_chrome_path() -> Optional[str]:
    """
    在当前系统中自动查找 Chrome 可执行文件路径
    返回：Chrome 完整路径；找不到返回 None
    """
    system = platform.system()

    if system == "Windows":
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
    elif system == "Darwin":  # macOS
        candidates = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/usr/local/bin/google-chrome",
        ]
    else:  # Linux
        candidates = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser",
            "/usr/bin/chromium",
            "/snap/bin/chromium",
        ]

    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def find_free_port(start_port: int = 9222) -> int:
    """
    查找可用端口
    """
    port = start_port
    while port < start_port + 100:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("localhost", port))
                return port
        except OSError:
            port += 1
    raise RuntimeError(f"在 {start_port}-{start_port + 99} 范围内找不到可用端口")


def start_chrome_with_debug_port(port: int, user_data_dir: str) -> Optional[subprocess.Popen]:
    """
    以调试端口启动系统 Chrome 浏览器
    返回：进程对象；启动失败返回 None
    """
    chrome_path = find_chrome_path()
    if not chrome_path:
        print("[错误] 未找到系统 Chrome 浏览器。请确保已安装 Chrome。")
        return None

    os.makedirs(user_data_dir, exist_ok=True)

    cmd = [
        chrome_path,
        f"--remote-debugging-port={port}",
        "--remote-debugging-address=0.0.0.0",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-blink-features=AutomationControlled",
        "--disable-infobars",
        "--disable-dev-shm-usage",
        "--no-sandbox",
        "--start-maximized",
        "about:blank",
    ]

    try:
        if platform.system() == "Windows":
            process = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                preexec_fn=os.setsid,
            )

        # 等待 Chrome 启动并监听端口
        deadline = time.time() + 20
        while time.time() < deadline:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(1)
                    if s.connect_ex(("localhost", port)) == 0:
                        print(f"[启动] Chrome 已启动，调试端口: {port}")
                        print(f"[启动] Chrome 路径: {chrome_path}")
                        print(f"[启动] 用户数据: {user_data_dir}")
                        return process
            except Exception:
                pass
            time.sleep(0.5)

        print("[错误] Chrome 启动超时，未能连接到调试端口")
        return None

    except Exception as e:
        print(f"[错误] 启动 Chrome 失败: {e}")
        return None


async def connect_browser_via_cdp(port: int):
    """
    通过 CDP 协议连接已启动的 Chrome（Playwright 客户端模式）
    返回：(playwright, browser, context, page)
    """
    from playwright.async_api import async_playwright

    # 1. 启动 Playwright
    pw = await async_playwright().start()

    # 2. 从 CDP 端点获取 WebSocket URL
    async with httpx.AsyncClient(timeout=10) as client:
        resp = await client.get(f"http://localhost:{port}/json/version")
        ws_url = resp.json().get("webSocketDebuggerUrl")

    if not ws_url:
        raise RuntimeError("无法从 CDP 端点获取 WebSocket URL")

    print(f"[连接] 通过 CDP 连接浏览器: {ws_url}")

    # 3. 连接
    browser = await pw.chromium.connect_over_cdp(ws_url)

    # 4. 获取或创建上下文
    if browser.contexts:
        context = browser.contexts[0]
        print("[连接] 复用现有浏览器上下文")
    else:
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0.0.0 Safari/537.36"
            ),
        )
        print("[连接] 创建新的浏览器上下文")

    # 5. 获取或创建页面
    if context.pages:
        page = context.pages[0]
    else:
        page = await context.new_page()

    print("[连接] 浏览器已就绪")
    return pw, browser, context, page


# ------------------------------------------------------------
#  2. URL 构造
# ------------------------------------------------------------
def build_list_url(city_name: str, city_id: str = "", page: int = 1) -> str:
    """
    构造携程游记列表页 URL
    第1页: https://you.ctrip.com/travels/lanzhou231.html
    第N页: https://you.ctrip.com/travels/lanzhou231-pN.html
    """
    if not city_id:
        city_id = CITY_ID_MAP.get(city_name.lower(), "")
    identifier = f"{city_name}{city_id}" if city_id else city_name
    if page == 1:
        return f"https://you.ctrip.com/travels/{identifier}.html"
    return f"https://you.ctrip.com/travels/{identifier}-p{page}.html"


# ------------------------------------------------------------
#  3. 数据提取（通过浏览器页面直接执行 JS）
# ------------------------------------------------------------
async def extract_travel_list(page) -> List[Dict]:
    """
    从当前列表页提取游记信息
    每个游记返回: {title, travel_id, detail_url, author, date, view_count}
    """
    items = await page.evaluate("""
        () => {
            const results = [];
            const hrefPattern = /\\/travels\\/\\w+\\d+\\/(\\d+)\\.html/;

            // 遍历所有包含游记链接的元素
            document.querySelectorAll('a[href*="/travels/"]').forEach(link => {
                const href = link.getAttribute('href') || '';
                const match = href.match(hrefPattern);
                if (!match) return;

                const travelId = match[1];
                const title = (link.textContent || '').trim();
                if (!title || title.length < 2) return;

                // 找到父容器，提取作者、日期、阅读数
                const parent = link.closest('div, li, article, section');
                if (!parent) return;

                // 作者
                let author = '';
                const authorSel = parent.querySelector(
                    '.author, .user-name, [class*="author"], .name, span.name'
                );
                if (authorSel) author = authorSel.textContent.trim();

                // 日期
                let date = '';
                const dateSel = parent.querySelector('.date, .time, [class*="time"]');
                if (dateSel) date = dateSel.textContent.trim();

                // 阅读
                let view = '';
                const viewSel = parent.querySelector('.view, .views, [class*="view"]');
                if (viewSel) view = viewSel.textContent.trim();

                results.push({
                    title,
                    travelId,
                    href,
                    author,
                    date,
                    view_count: view
                });
            });

            // 去重（按 travelId）
            const seen = new Set();
            return results.filter(item => {
                if (seen.has(item.travelId)) return false;
                seen.add(item.travelId);
                return true;
            });
        }
    """)

    # 整理字段，补全详情 URL
    travels = []
    for item in items:
        href = item["href"]
        detail_url = f"https://you.ctrip.com{href}" if href.startswith("/") else href
        travels.append({
            "title": item["title"],
            "travel_id": item["travelId"],
            "detail_url": detail_url,
            "author": item.get("author") or "未知",
            "date": item.get("date") or "",
            "view_count": item.get("view_count") or "",
        })

    return travels


async def extract_travel_detail(page) -> Dict:
    """
    从详情页提取正文内容
    返回：{title, author, publish_date, content, images, tags}
    """
    data = await page.evaluate("""
        () => {
            // 标题
            let title = '';
            const titleSel = document.querySelector('h1.title, h1, .article-title');
            if (titleSel) title = titleSel.textContent.trim();

            // 作者
            let author = '';
            const authorSel = document.querySelector(
                '.author-name, .author, [class*="author"], a[href*="/user/"]'
            );
            if (authorSel) author = authorSel.textContent.trim();

            // 发布日期
            let pubDate = '';
            const dateSel = document.querySelector(
                '.publish-time, .publishTime, time, [class*="publish"]'
            );
            if (dateSel) pubDate = dateSel.textContent.trim();

            // 正文内容（核心！优先从文章容器提取，失败则回退到 body）
            let content = '';
            const contentSel = (
                document.querySelector('div.article-content') ||
                document.querySelector('div.article_content') ||
                document.querySelector('div.content') ||
                document.querySelector('#article-content') ||
                document.querySelector('div.detail-content') ||
                document.querySelector('article')
            );

            if (contentSel) {
                const nodes = contentSel.querySelectorAll(
                    'p, h1, h2, h3, h4, li, div[class*="para"], span[class*="text"]'
                );
                if (nodes.length > 0) {
                    const lines = [];
                    nodes.forEach(n => {
                        const t = n.textContent.trim();
                        if (t && t.length > 2) lines.push(t);
                    });
                    content = lines.join('\\n');
                } else {
                    content = contentSel.innerText || contentSel.textContent || '';
                }
            }

            // 如果从文章容器拿不到足够内容，从 body 拿
            if (content.length < 200) {
                const body = document.body;
                if (body) {
                    const clone = body.cloneNode(true);
                    ['script', 'style', 'nav', 'header', 'footer', 'aside'].forEach(sel => {
                        clone.querySelectorAll(sel).forEach(el => el.remove());
                    });
                    content = clone.innerText || clone.textContent || '';
                }
            }

            // 图片
            const images = Array.from(document.querySelectorAll('img'))
                .map(img => img.src || img.getAttribute('data-src') || '')
                .filter(src => src && (src.startsWith('http') || src.startsWith('//')))
                .map(src => src.startsWith('//') ? 'https:' + src : src)
                .slice(0, 20);

            // 标签
            const tags = Array.from(document.querySelectorAll(
                '.tag, .tags a, [class*="tag"] a, .keywords a'
            ))
                .map(t => t.textContent.trim())
                .filter(t => t && t.length < 20);

            return {
                title,
                author: author || '未知',
                publish_date: pubDate,
                content: content.trim(),
                images,
                tags
            };
        }
    """)

    return data


# ------------------------------------------------------------
#  4. 保存数据
# ------------------------------------------------------------
def save_results(travels: List[Dict], fmt: str, city: str, outdir: str):
    """
    保存结果
    """
    os.makedirs(outdir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if fmt.lower() == "excel":
        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, PatternFill
        except ImportError:
            print("[警告] 缺少 openpyxl，降级为 TXT 保存")
            fmt = "txt"

    if fmt.lower() == "txt":
        path = os.path.join(outdir, f"ctrip_{city}_{timestamp}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("=" * 78 + "\n")
            f.write(f"  携程游记爬取结果 - 城市: {city}\n")
            f.write(f"  共 {len(travels)} 篇 | 采集时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 78 + "\n")

            for idx, t in enumerate(travels, 1):
                f.write(f"\n{'─' * 78}\n")
                f.write(f"  【第 {idx} 篇】\n")
                f.write(f"{'─' * 78}\n")
                f.write(f"  标题: {t.get('title', '')}\n")
                f.write(f"  作者: {t.get('author', '')}\n")
                f.write(f"  日期: {t.get('date', '') or t.get('publish_date', '')}\n")
                f.write(f"  阅读: {t.get('view_count', '')}\n")
                f.write(f"  链接: {t.get('detail_url', '')}\n")
                if t.get("tags"):
                    f.write(f"  标签: {', '.join(t['tags'])}\n")
                f.write(f"\n  --- 正文 ---\n\n")
                content = t.get("content", "")
                if content:
                    f.write("  " + content.replace("\n", "\n  ") + "\n")
                else:
                    f.write("  （未获取到内容）\n")

        print(f"\n[保存] 已写入 TXT: {path}")

    elif fmt.lower() == "excel":
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill

        path = os.path.join(outdir, f"ctrip_{city}_{timestamp}.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "携程游记"

        headers = ["序号", "标题", "作者", "日期", "阅读", "标签", "正文", "链接"]
        fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
        font = Font(bold=True, color="FFFFFF")
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center")

        widths = [6, 40, 12, 15, 10, 20, 50, 45]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

        for row, t in enumerate(travels, 2):
            ws.cell(row=row, column=1, value=row - 1)
            ws.cell(row=row, column=2, value=t.get("title", ""))
            ws.cell(row=row, column=3, value=t.get("author", ""))
            ws.cell(row=row, column=4, value=t.get("date", "") or t.get("publish_date", ""))
            ws.cell(row=row, column=5, value=t.get("view_count", ""))
            ws.cell(row=row, column=6, value=", ".join(t.get("tags", [])))
            ws.cell(row=row, column=7, value=(t.get("content", "") or "")[:30000])
            ws.cell(row=row, column=8, value=t.get("detail_url", ""))

        wb.save(path)
        print(f"\n[保存] 已写入 Excel: {path}")


# ------------------------------------------------------------
#  5. 主流程
# ------------------------------------------------------------
async def main_async():
    """
    异步主流程：启动浏览器 → 爬取列表 → 爬取详情 → 保存
    """
    # 参数准备
    city_id = CITY_ID or CITY_ID_MAP.get(CITY_NAME.lower(), "")
    if not city_id:
        print(f"[警告] 未在 CITY_ID_MAP 中找到城市 {CITY_NAME}，请手动设置 CITY_ID")
    print("=" * 70)
    print(f"  携程游记爬虫 v5.0  |  城市: {CITY_NAME}({city_id})  |  页数: {DEFAULT_PAGES}")
    print("=" * 70)

    # --- 1. 启动 Chrome ---
    print("\n[步骤1] 启动系统 Chrome 浏览器...")
    user_data_dir = os.path.join(os.getcwd(), "chrome_data", f"ctrip_{CITY_NAME}")
    port = find_free_port(CHROME_DEBUG_PORT)
    chrome_process = start_chrome_with_debug_port(port, user_data_dir)
    if not chrome_process:
        print("[错误] 无法启动 Chrome，程序退出")
        return

    # 给浏览器一点时间
    await asyncio.sleep(2)

    # --- 2. 连接浏览器 ---
    pw, browser, context, page = None, None, None, None
    try:
        print("\n[步骤2] 连接到浏览器...")
        pw, browser, context, page = await connect_browser_via_cdp(port)

        # --- 3. 爬取列表 ---
        print("\n[步骤3] 爬取游记列表...")
        all_travels: List[Dict] = []

        for p in range(1, DEFAULT_PAGES + 1):
            url = build_list_url(CITY_NAME, city_id, p)
            print(f"  [{p}/{DEFAULT_PAGES}] 访问 {url}")

            # 访问页面，等待网络稳定
            try:
                await page.goto(url, timeout=60000, wait_until="networkidle")
            except Exception:
                # 某些页面 networkidle 触发慢，退化为 domcontentloaded
                await page.goto(url, timeout=60000, wait_until="domcontentloaded")
            await asyncio.sleep(REQUEST_DELAY)

            # 提取
            travels = await extract_travel_list(page)
            print(f"          找到 {len(travels)} 篇")
            all_travels.extend(travels)

            # 翻页间隔
            if p < DEFAULT_PAGES:
                await asyncio.sleep(REQUEST_DELAY)

        print(f"\n  [列表] 共获取 {len(all_travels)} 篇游记")

        if not all_travels:
            print("[错误] 未获取到任何游记，请检查城市ID是否正确")
            return

        # --- 4. 爬取详情 ---
        print(f"\n[步骤4] 爬取 {len(all_travels)} 篇游记详情...")
        for idx, t in enumerate(all_travels, 1):
            short_title = t.get("title", "")[:25]
            print(f"  [{idx}/{len(all_travels)}] {short_title}...", end=" ", flush=True)

            try:
                await page.goto(t["detail_url"], timeout=60000, wait_until="domcontentloaded")
                await asyncio.sleep(REQUEST_DELAY)
                detail = await extract_travel_detail(page)
                t.update(detail)
                print("✅")
            except Exception as e:
                print(f"❌ ({e})")
                continue

        # --- 5. 保存 ---
        print("\n[步骤5] 保存数据...")
        save_results(all_travels, OUTPUT_FORMAT, CITY_NAME, OUTPUT_DIR)
        print(f"\n🎉 全部完成！共采集 {len(all_travels)} 篇游记")

    finally:
        # --- 6. 清理（仅关闭本次创建的 page 和上下文，保留浏览器进程给用户查看）---
        print("\n[清理] 释放资源...")
        try:
            if page:
                await page.close()
        except Exception:
            pass
        try:
            if browser and browser.is_connected():
                await browser.close()
        except Exception:
            pass
        try:
            if pw:
                await pw.stop()
        except Exception:
            pass
        # 不主动关闭 Chrome 进程，让用户可以继续查看页面
        print("[清理] 已断开与浏览器的连接（Chrome 窗口保持打开）")


def main():
    """
    同步入口
    """
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        print("\n\n[中断] 用户取消")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
