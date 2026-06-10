# -*- coding: utf-8 -*-
# =============================================================================
#  携程游记爬虫 - 徐州（Xuzhou）
#  参考项目 tools/browser_launcher.py + tools/cdp_browser.py 的架构：
#    1. 自动启动系统 Chrome（带远程调试端口 + user-data-dir 持久化登录态）
#    2. 打开携程登录页，用户扫码/手机号登录
#    3. 按回车后自动通过 CDP 抓浏览器 Cookie
#    4. 用 requests + lxml 爬取列表页与详情页
#    5. 保存为 TXT 或 Excel
# =============================================================================

import os
import re
import sys
import time
import socket
import signal
import atexit
import platform
import subprocess
import asyncio
from datetime import datetime
from typing import Optional, List, Dict, Tuple
from pathlib import Path

import requests
from lxml import etree
from fake_useragent import UserAgent
from requests.exceptions import RequestException

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  【基础配置区】—— 所有需要手动调整的参数都在这里
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# —— 爬取目标 ——
CITY_CODE = 230              # 城市代码：徐州=230（URL里是 Xuzhou230）
CITY_NAME = "Xuzhou"         # 城市英文/拼音，用于URL拼接和文件名
START_PAGE = 1               # 起始页
END_PAGE = 5                 # 结束页

# —— 请求配置 ——
THREAD_POOL_SIZE = 3         # 并发线程数（建议 3-5，过大易被拦截）
RETRY_TIMES = 3              # 每页/详情的重试次数
REQUEST_TIMEOUT = 20         # 单次请求超时（秒）
LIST_PAGE_INTERVAL = 1.0     # 列表页之间的等待时间
DETAIL_PAGE_INTERVAL = 0.8   # 详情页之间的等待时间

# —— 浏览器 & Cookie ——
CHROME_DEBUG_PORT = 9222     # Chrome 远程调试端口
BROWSER_LAUNCH_TIMEOUT = 30  # 浏览器启动超时
HEADLESS = False             # 是否无头（首次登录建议 False，方便看到二维码）
LOGIN_URL = "https://you.ctrip.com"   # 携程游记首页，让用户在这里登录
LOGIN_WAIT_PROMPT = True              # 是否等用户按回车再抓 Cookie
SAVE_LOGIN_STATE = True               # 是否持久化用户数据（下次可能免登录）
COOKIE_FILE = f".cookies_ctrip_{CITY_NAME.lower()}.json"   # Cookie 缓存文件

# —— 输出配置 ——
OUTPUT_FORMAT = "excel"      # 输出格式：excel 或 txt
OUTPUT_DIR = "./output"      # 输出目录（相对或绝对路径都行）
OUTPUT_FILENAME = ""         # 自定义文件名（不带后缀），留空则自动生成

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  【URL 模板】—— 一般不用改
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BASE_URL = f"https://you.ctrip.com/travels/{CITY_NAME}{CITY_CODE}/t3-p{{page}}.html"
TRAVEL_BASE_URL = "https://you.ctrip.com"

ua = UserAgent()


# =============================================================================
#  1. 浏览器启动器（对应 tools/browser_launcher.py 的最小实现版）
# =============================================================================

class BrowserLauncher:
    """自动探测系统 Chrome/Edge 并以 CDP 模式启动，保持登录态。"""

    def __init__(self):
        self.system = platform.system()
        self.browser_process: Optional[subprocess.Popen] = None

    # ---------------- 浏览器路径探测 ----------------
    def detect_browser_paths(self) -> List[str]:
        paths: List[str] = []
        if self.system == "Windows":
            candidates = [
                os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe"),
            ]
        elif self.system == "Darwin":
            candidates = [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            ]
        else:
            candidates = [
                "/usr/bin/google-chrome",
                "/usr/bin/google-chrome-stable",
                "/usr/bin/chromium-browser",
                "/usr/bin/chromium",
            ]
        for p in candidates:
            if os.path.isfile(p) and (self.system != "Windows" or os.access(p, os.X_OK)):
                paths.append(p)
            elif os.path.isfile(p):
                paths.append(p)
        return paths

    # ---------------- 端口探测 ----------------
    def find_available_port(self, start_port: int = 9222) -> int:
        port = start_port
        for _ in range(100):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(("localhost", port))
                    return port
            except OSError:
                port += 1
        raise RuntimeError(f"无可用端口，已尝试 {start_port}-{start_port + 99}")

    # ---------------- 启动浏览器 ----------------
    def launch(self, browser_path: str, debug_port: int, user_data_dir: Optional[str] = None) -> subprocess.Popen:
        args = [
            browser_path,
            f"--remote-debugging-port={debug_port}",
            "--remote-debugging-address=0.0.0.0",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-blink-features=AutomationControlled",
            "--exclude-switches=enable-automation",
            "--disable-infobars",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--start-maximized",
        ]
        if HEADLESS:
            args.extend(["--headless=new", "--disable-gpu"])
        if user_data_dir:
            os.makedirs(user_data_dir, exist_ok=True)
            args.append(f"--user-data-dir={user_data_dir}")

        # 打开携程首页，确保一启动就在目标域下
        args.append(LOGIN_URL)

        try:
            if self.system == "Windows":
                proc = subprocess.Popen(
                    args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
                )
            else:
                proc = subprocess.Popen(
                    args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, preexec_fn=os.setsid
                )
            self.browser_process = proc
            return proc
        except Exception as e:
            raise RuntimeError(f"启动浏览器失败: {e}")

    # ---------------- 等待端口就绪 ----------------
    def wait_for_ready(self, debug_port: int, timeout: int = BROWSER_LAUNCH_TIMEOUT) -> bool:
        start = time.time()
        while time.time() - start < timeout:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(1)
                    if s.connect_ex(("localhost", debug_port)) == 0:
                        return True
            except Exception:
                pass
            time.sleep(0.5)
        return False

    # ---------------- 清理 ----------------
    def cleanup(self):
        if not self.browser_process or self.browser_process.poll() is not None:
            return
        try:
            if self.system == "Windows":
                self.browser_process.terminate()
                try:
                    self.browser_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(self.browser_process.pid)],
                        capture_output=True, check=False
                    )
            else:
                pgid = os.getpgid(self.browser_process.pid)
                os.killpg(pgid, signal.SIGTERM)
                try:
                    self.browser_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    os.killpg(pgid, signal.SIGKILL)
            self.browser_process = None
        except Exception:
            pass


# =============================================================================
#  2. CDP 浏览器 Cookie 抓取器（对应 tools/cdp_browser.py 的最小实现）
# =============================================================================

class CDPCookieGrabber:
    """通过 CDP + Playwright 连接已启动的浏览器，抓取 Cookie。"""

    @staticmethod
    async def grab_cookies(debug_port: int, user_data_dir: Optional[str] = None) -> List[Dict]:
        """
        打开携程登录页，用户扫码/登录完成后按回车 → 返回 cookies 列表
        cookies 格式：[{"name": "...", "value": "...", "domain": "...", ...}, ...]
        """
        try:
            from playwright.async_api import async_playwright
        except ImportError:
            raise RuntimeError("缺少 playwright 依赖，请先执行: pip install playwright")

        async with async_playwright() as pw:
            # 获取 WebSocket URL
            try:
                resp = requests.get(f"http://localhost:{debug_port}/json/version", timeout=10)
                ws_url = resp.json().get("webSocketDebuggerUrl")
            except Exception as e:
                raise RuntimeError(f"无法连接 CDP 端口 {debug_port}: {e}")

            if not ws_url:
                raise RuntimeError("CDP 端点未返回 webSocketDebuggerUrl")

            print(f"\n>>> 已连接浏览器 CDP: {ws_url}")
            browser = await pw.chromium.connect_over_cdp(ws_url)

            # 优先使用现有上下文（因为我们刚通过 subprocess 打开了 LOGIN_URL）
            if browser.contexts:
                context = browser.contexts[0]
            else:
                context = await browser.new_context(viewport={"width": 1920, "height": 1080})

            # 打开/等待一个页面
            if context.pages:
                page = context.pages[0]
            else:
                page = await context.new_page()

            # 打开携程游记首页
            try:
                await page.goto(LOGIN_URL, timeout=60000, wait_until="domcontentloaded")
            except Exception:
                pass

            print("\n" + "=" * 60)
            print("  >>> 请在刚打开的 Chrome 窗口中完成携程登录（扫码/手机号均可）")
            print("  >>> 登录完成后，回到本终端按【回车键】继续...")
            print("  >>> (若想跳过登录直接爬取，也可直接回车)")
            print("=" * 60 + "\n")

            # 等待用户按回车
            if LOGIN_WAIT_PROMPT:
                try:
                    input("")
                except (EOFError, KeyboardInterrupt):
                    print("\n用户取消登录步骤。")

            # 抓取 cookies
            cookies = await context.cookies()
            print(f">>> 已抓取到 {len(cookies)} 个 Cookie")

            # 关闭 Playwright 连接（不杀 Chrome 进程）
            await page.close()
            await context.close()
            await browser.close()

            return cookies


# =============================================================================
#  3. Cookie 工具
# =============================================================================

def cookies_to_header(cookies: List[Dict]) -> str:
    """把 CDP 抓出的 cookies 列表拼成 HTTP Cookie 请求头字符串。"""
    parts = []
    for c in cookies:
        try:
            parts.append(f"{c['name']}={c['value']}")
        except KeyError:
            continue
    return "; ".join(parts)


def load_cached_cookies() -> Optional[List[Dict]]:
    """从缓存文件读取 cookies（避免每次都扫码）。"""
    if SAVE_LOGIN_STATE and os.path.isfile(COOKIE_FILE):
        try:
            import json
            with open(COOKIE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and data:
                print(f">>> 从缓存文件 {COOKIE_FILE} 读取到 {len(data)} 个 Cookie")
                use = input(">>> 是否直接使用缓存 Cookie？[Y/n] ").strip().lower()
                if use in ("", "y", "yes"):
                    return data
        except Exception:
            pass
    return None


def save_cookies_to_cache(cookies: List[Dict]):
    """保存 cookies 到文件，供下次复用。"""
    if not SAVE_LOGIN_STATE:
        return
    try:
        import json
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)
        print(f">>> Cookie 已缓存到 {COOKIE_FILE}")
    except Exception:
        pass


# =============================================================================
#  4. 请求 & 解析层
# =============================================================================

def build_headers(cookie_str: str) -> Dict[str, str]:
    """构造统一的请求头。"""
    return {
        "User-Agent": ua.random,
        "Cookie": cookie_str,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Referer": "https://you.ctrip.com/",
    }


def fetch_html(url: str, headers: Dict[str, str]) -> Optional[str]:
    """带重试的 HTML 请求。"""
    for attempt in range(1, RETRY_TIMES + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 200 and resp.text:
                return resp.text
            print(f"  [警告] 状态码 {resp.status_code}，第 {attempt}/{RETRY_TIMES} 次重试...")
        except RequestException as e:
            print(f"  [警告] 请求异常: {e}，第 {attempt}/{RETRY_TIMES} 次重试...")
        time.sleep(2)
    return None


def parse_travel_list(html: str) -> List[Dict]:
    """解析游记列表页，返回 {title, url, author, summary} 列表。"""
    results: List[Dict] = []
    selector = etree.HTML(html)
    if selector is None:
        return results

    # 尝试多种容器选择器
    candidates = (
        selector.xpath("//a[contains(@class, 'cursor-pointer')]")
        or selector.xpath("//div[contains(@class, 'travel-item')]//a")
        or selector.xpath("//div[contains(@class, 'journal-item')]//a")
        or selector.xpath("//a[contains(@href, '/travels/')]")
    )

    seen = set()
    for a in candidates:
        href = "".join(a.xpath("./@href")).strip()
        if not href:
            continue
        title = "".join(a.xpath(".//text()")).strip() or \
                "".join(a.xpath("./text()")).strip()
        if not title or len(title) < 2:
            continue
        if not href.startswith("http"):
            href = TRAVEL_BASE_URL + (href if href.startswith("/") else "/" + href)
        # 去重（按 URL）
        if href in seen:
            continue
        seen.add(href)

        results.append({
            "title": title,
            "url": href,
            "author": "",
            "summary": "",
        })

    return results


def parse_travel_detail(html: str) -> Dict[str, str]:
    """解析游记详情页：正文、图片、出游天数、人均、发布日期等。"""
    sel = etree.HTML(html)
    result = {
        "content": "",
        "image_urls": [],
        "travel_days": "",
        "per_capita_cost": "",
        "publish_date": "",
    }
    if sel is None:
        return result

    # 正文：尝试多种选择器
    text_parts: List[str] = []
    for xpath in [
        "//div[contains(@class, 'ctrip-article-content')]//text()",
        "//div[contains(@class, 'travel-content')]//text()",
        "//div[@class='rich_media_content']//text()",
        "//article//text()",
        "//div[contains(@class, 'content')]//text()",
    ]:
        nodes = sel.xpath(xpath)
        if nodes:
            text_parts = [str(n).strip() for n in nodes if str(n).strip()]
            if len(text_parts) > 3:
                break

    result["content"] = re.sub(r"\s+", " ", " ".join(text_parts))

    # 图片
    for xpath in [
        "//div[contains(@class, 'ctrip-article-content')]//img/@src",
        "//div[contains(@class, 'travel-content')]//img/@src",
        "//img[contains(@class, 'content-img')]/@src",
        "//article//img/@src",
    ]:
        imgs = sel.xpath(xpath)
        if imgs:
            cleaned: List[str] = []
            for src in imgs:
                src = str(src).strip()
                if not src or src.startswith("data:"):
                    continue
                if src.startswith("//"):
                    src = "https:" + src
                cleaned.append(src)
            if cleaned:
                result["image_urls"] = cleaned
                break

    # meta 区：天数/人均/发布日期
    meta_text = ""
    for xpath in [
        "//div[contains(@class, 'travel-info')]//text()",
        "//div[contains(@class, 'info-bar')]//text()",
        "//div[contains(@class, 'meta')]//text()",
    ]:
        nodes = sel.xpath(xpath)
        if nodes:
            meta_text = "".join(str(n) for n in nodes)
            break

    if meta_text:
        m = re.search(r"(\d+)\s*天", meta_text)
        if m:
            result["travel_days"] = m.group(1)
        m = re.search(r"人均[花费]*[\:：]?\s*(\d+)", meta_text)
        if m:
            result["per_capita_cost"] = m.group(1)
        m = re.search(r"(\d{4}-\d{1,2}-\d{1,2})", meta_text)
        if m:
            result["publish_date"] = m.group(1)

    return result


# =============================================================================
#  5. 输出层：TXT / Excel
# =============================================================================

def ensure_output_dir():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)


def output_filename() -> str:
    if OUTPUT_FILENAME:
        return OUTPUT_FILENAME
    return f"ctrip_{CITY_NAME.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def save_to_excel(data: List[Dict]):
    try:
        import openpyxl
    except ImportError:
        print(">>> 缺少 openpyxl，降级保存为 TXT")
        save_to_txt(data)
        return

    ensure_output_dir()
    path = os.path.join(OUTPUT_DIR, output_filename() + ".xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = CITY_NAME

    headers = ["序号", "标题", "作者", "出游天数", "人均花费",
               "发布日期", "摘要/正文前500字", "图片链接(分号分隔)", "原始链接"]
    ws.append(headers)

    for i, item in enumerate(data, 1):
        row = [
            i,
            item.get("title", ""),
            item.get("author", ""),
            item.get("travel_days", ""),
            item.get("per_capita_cost", ""),
            item.get("publish_date", ""),
            item.get("content", "")[:500],
            "; ".join(item.get("image_urls", [])[:10]),
            item.get("url", ""),
        ]
        ws.append(row)

    wb.save(path)
    print(f">>> Excel 已保存到: {path}")


def save_to_txt(data: List[Dict]):
    ensure_output_dir()
    path = os.path.join(OUTPUT_DIR, output_filename() + ".txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write(f"  携程游记爬取结果 - {CITY_NAME}\n")
        f.write(f"  共 {len(data)} 篇 | 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 70 + "\n\n")

        for i, item in enumerate(data, 1):
            f.write(f"━━━ 【{i:>3}/{len(data)}】 {item.get('title', '无标题')} ━━━━━━━━━━━━━━\n")
            f.write(f"作者   : {item.get('author', '')}\n")
            f.write(f"天数   : {item.get('travel_days', '')} 天\n")
            f.write(f"人均   : {item.get('per_capita_cost', '')} 元\n")
            f.write(f"发布   : {item.get('publish_date', '')}\n")
            f.write(f"链接   : {item.get('url', '')}\n")
            if item.get("image_urls"):
                f.write(f"图片   :\n")
                for img in item["image_urls"][:10]:
                    f.write(f"         - {img}\n")
            f.write("\n【正文】\n")
            f.write(item.get("content", "(未解析到正文)"))
            f.write("\n\n\n")

    print(f">>> TXT 已保存到: {path}")


# =============================================================================
#  6. 主流程
# =============================================================================

_launcher_for_cleanup: Optional[BrowserLauncher] = None


def _atexit_cleanup():
    if _launcher_for_cleanup:
        _launcher_for_cleanup.cleanup()


async def get_cookie_string() -> str:
    """获取可用 Cookie：优先使用缓存 → 否则启动浏览器让用户扫码。"""
    # 1) 尝试缓存
    cached = load_cached_cookies()
    if cached:
        return cookies_to_header(cached)

    # 2) 启动浏览器
    global _launcher_for_cleanup
    launcher = BrowserLauncher()
    _launcher_for_cleanup = launcher
    atexit.register(_atexit_cleanup)

    paths = launcher.detect_browser_paths()
    if not paths:
        raise RuntimeError(
            "未检测到 Chrome/Edge 浏览器。请先安装 Chrome，或手动把浏览器路径填到 "
            "BrowserLauncher.detect_browser_paths 的 candidates 列表里。"
        )

    port = launcher.find_available_port(CHROME_DEBUG_PORT)
    user_data_dir = os.path.join(os.getcwd(), "browser_data", f"ctrip_{CITY_NAME.lower()}")
    print(f">>> 使用浏览器: {paths[0]}")
    print(f">>> 调试端口: {port}")
    print(f">>> 用户数据目录: {user_data_dir}")
    print(">>> 即将打开 Chrome，请在浏览器中扫码登录携程...")

    launcher.launch(paths[0], port, user_data_dir=user_data_dir)
    if not launcher.wait_for_ready(port):
        raise RuntimeError(f"浏览器启动超时（端口 {port}）")

    # 3) 通过 CDP 抓 Cookie
    cookies = await CDPCookieGrabber.grab_cookies(port)
    save_cookies_to_cache(cookies)
    return cookies_to_header(cookies)


def main_sync():
    """同步主入口：列表 + 详情 爬取。"""
    print("=" * 70)
    print(f"  携程游记爬虫 | 目标城市：{CITY_NAME}({CITY_CODE}) | "
          f"页码：{START_PAGE}-{END_PAGE}")
    print("=" * 70)

    # ---- A. 获取 Cookie ----
    cookie_str = asyncio.run(get_cookie_string())
    if not cookie_str:
        print(">>> [警告] 未获取到有效 Cookie，将尝试无 Cookie 爬取（可能会被拦截）")

    headers = build_headers(cookie_str)

    # ---- B. 爬取列表 ----
    all_travels: List[Dict] = []
    print("\n>>> 阶段 1：爬取游记列表")
    for page in range(START_PAGE, END_PAGE + 1):
        url = BASE_URL.format(page=page)
        print(f"  [P{page}] {url}")
        html = fetch_html(url, headers)
        if not html:
            print(f"    ✗ 列表页获取失败，跳过")
            continue
        travels = parse_travel_list(html)
        print(f"    ✓ 解析到 {len(travels)} 篇")
        all_travels.extend(travels)
        time.sleep(LIST_PAGE_INTERVAL)

    if not all_travels:
        print(">>> 未找到任何游记。可能是 Cookie 失效或页面结构变化。")
        return

    print(f"\n>>> 阶段 2：爬取 {len(all_travels)} 篇游记详情（并发={THREAD_POOL_SIZE}）")

    # ---- C. 爬取详情 ----
    def fetch_detail(idx_url: Tuple[int, Dict]) -> Dict:
        idx, item = idx_url
        print(f"  [{idx + 1}/{len(all_travels)}] {item.get('title', '')[:40]}")
        html = fetch_html(item["url"], headers)
        if html:
            detail = parse_travel_detail(html)
            item.update(detail)
        time.sleep(DETAIL_PAGE_INTERVAL)
        return item

    if THREAD_POOL_SIZE > 1:
        from multiprocessing.dummy import Pool
        with Pool(THREAD_POOL_SIZE) as pool:
            all_travels = pool.map(fetch_detail, enumerate(all_travels))
    else:
        all_travels = [fetch_detail(p) for p in enumerate(all_travels)]

    # ---- D. 保存 ----
    print("\n>>> 阶段 3：保存结果")
    fmt = OUTPUT_FORMAT.lower()
    if fmt == "excel":
        save_to_excel(all_travels)
    else:
        save_to_txt(all_travels)

    # ---- E. 清理浏览器进程 ----
    if _launcher_for_cleanup:
        _launcher_for_cleanup.cleanup()
        atexit.unregister(_atexit_cleanup)
        print(">>> 浏览器进程已关闭")

    print(f"\n🎉 完成！共处理 {len(all_travels)} 篇游记")


if __name__ == "__main__":
    try:
        main_sync()
    except KeyboardInterrupt:
        print("\n>>> 用户中断。")
        if _launcher_for_cleanup:
            _launcher_for_cleanup.cleanup()
        sys.exit(1)
    except RuntimeError as e:
        print(f"\n❌ {e}")
        sys.exit(1)
