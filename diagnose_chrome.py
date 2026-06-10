#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速诊断脚本：测试系统 Chrome 能否被成功启动并通过 CDP 连接
    python diagnose_chrome.py
"""
import socket
import subprocess
import time
import asyncio
import httpx
import platform
import os


def find_chrome():
    system = platform.system()
    if system == "Windows":
        cands = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
    elif system == "Darwin":
        cands = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
    else:
        cands = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser",
            "/usr/bin/chromium",
        ]
    for p in cands:
        if os.path.isfile(p):
            return p
    return None


async def test_chrome():
    print("=" * 60)
    print("  Chrome 浏览器诊断工具")
    print("=" * 60)

    # 1. 检测 Chrome 路径
    chrome_path = find_chrome()
    if not chrome_path:
        print("\n❌ 未找到 Chrome 浏览器，请先安装 Chrome")
        return
    print(f"✅ 找到 Chrome: {chrome_path}")

    # 2. 查找可用端口
    port = 9222
    while True:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("localhost", port))
                break
        except OSError:
            port += 1
    print(f"✅ 使用调试端口: {port}")

    # 3. 启动 Chrome
    print(f"\n▶ 启动 Chrome (调试端口 {port})...")
    data_dir = os.path.join(os.getcwd(), "chrome_test_data")
    os.makedirs(data_dir, exist_ok=True)

    cmd = [
        chrome_path,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={data_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-blink-features=AutomationControlled",
        "about:blank",
    ]
    try:
        if platform.system() == "Windows":
            proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, preexec_fn=os.setsid
            )
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        return

    # 4. 等待并测试端口
    print("▶ 等待端口就绪...", end=" ", flush=True)
    ready = False
    for _ in range(30):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1)
                if s.connect_ex(("localhost", port)) == 0:
                    ready = True
                    break
        except Exception:
            pass
        time.sleep(0.5)

    if not ready:
        print("❌ 超时，端口未就绪")
        proc.terminate()
        return
    print("✅")

    # 5. 测试 CDP 端点
    print("▶ 查询 CDP 端点...", end=" ", flush=True)
    async with httpx.AsyncClient(timeout=10) as client:
        resp = await client.get(f"http://localhost:{port}/json/version")
        data = resp.json()
    ws_url = data.get("webSocketDebuggerUrl")
    if not ws_url:
        print("❌ 未获取到 WebSocket URL")
        return
    print(f"✅\n   WebSocket: {ws_url}")

    # 6. 使用 Playwright 连接
    print("\n▶ Playwright 连接测试...", end=" ", flush=True)
    try:
        from playwright.async_api import async_playwright
        pw = await async_playwright().start()
        browser = await pw.chromium.connect_over_cdp(ws_url)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = context.pages[0] if context.pages else await context.new_page()

        # 打开一个测试页面验证
        await page.goto("https://you.ctrip.com/travels/lanzhou231.html", timeout=30000, wait_until="domcontentloaded")
        title = await page.title()
        print(f"✅\n   页面标题: {title}")

        # 提取一些链接
        links = await page.evaluate("""
            () => {
                const out = [];
                document.querySelectorAll('a[href*="/travels/"]').forEach(a => {
                    const href = a.getAttribute('href');
                    const m = href && href.match(/\\/travels\\/\\w+\\d+\\/(\\d+)\\.html/);
                    if (m && out.length < 5) out.push({title: (a.textContent||'').slice(0,40), id: m[1]});
                });
                return out;
            }
        """)
        if links:
            print(f"\n✅ 成功解析到 {len(links)} 条测试游记:")
            for item in links:
                print(f"   - {item['title']} (ID: {item['id']})")
        else:
            print("\n⚠ 未解析到游记，可能需要登录或页面结构变化")

        # 清理
        await page.close()
        await browser.close()
        await pw.stop()
    except Exception as e:
        print(f"❌ 失败: {e}")
        return

    print("\n🎉 所有测试通过！Chrome 启动 + CDP 连接 + 页面抓取正常")

    # 关闭浏览器
    try:
        proc.terminate()
    except Exception:
        pass


if __name__ == "__main__":
    asyncio.run(test_chrome())
