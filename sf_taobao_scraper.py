#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 昆明市法拍房数据爬取工具 v0.13
=============================================
改进：
  - 自动安装依赖（selenium, requests, beautifulsoup4）
  - 全自动下载 Chrome/Edge WebDriver
  - 统一配置文件 crawl_config.json 管理所有状态与进度
  - 任何退出自动保存进度，启动时智能恢复
  - 无需手动输入 Cookie，浏览器登录后自动保存
  - 日志文件输出到输出目录
"""

import re
import os
import sys
import json
import time
import random
import logging
import csv
import zipfile
import shutil
import subprocess
import urllib.parse
from datetime import datetime
from pathlib import Path
import io

# 第三方库动态检查
try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    pass  # 将在依赖安装阶段处理

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        TimeoutException, NoSuchElementException,
        WebDriverException, StaleElementReferenceException
    )
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

# ============================================================
# 配置常量
# ============================================================

LIST_CITY = "昆明"
LIST_URL = "https://sf.taobao.com/list/50025969__2___{city_gbk}.htm"
ITEMS_PER_PAGE = 40
MAX_PAGES = 150
MIN_DELAY = 2.0
MAX_DELAY = 5.0

SUCCESS_KEYWORDS = ["本场已结束"]
FAIL_KEYWORDS = ["本场已流拍，无人出价"]

NON_RESIDENTIAL = ["商铺", "门面", "门市", "写字楼", "办公楼", "商业",
                   "土地", "厂房", "仓库", "在建工程", "机器设备", "车辆", "股权", "债权"]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
]

# ============================================================
# 日志配置（基础，稍后添加文件输出）
# ============================================================
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S")
logger = logging.getLogger(__name__)

# ============================================================
# 工具函数
# ============================================================

def get_script_dir():
    return str(Path(__file__).parent.resolve())

def random_delay(low=None, high=None):
    time.sleep(random.uniform(low or MIN_DELAY, high or MAX_DELAY))

def sanitize_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:200]

def safe_makedirs(path):
    os.makedirs(path, exist_ok=True)

def get_webdriver_path(browser_type):
    """返回 webdriver 的期望路径"""
    script_dir = get_script_dir()
    if browser_type == "edge":
        return os.path.join(script_dir, "msedgedriver.exe")
    return os.path.join(script_dir, "chromedriver.exe")

def download_chromedriver():
    """自动下载 ChromeDriver（需联网）"""
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "chromedriver.exe")
    if os.path.isfile(target):
        return target
    # 获取 Chrome 版本
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True, stderr=subprocess.STDOUT
        )
        version = re.search(r'REG_SZ\s+([\d.]+)', output.decode()).group(1)
    except:
        # 尝试用 powershell
        try:
            result = subprocess.run(
                ["powershell", "-Command",
                 "(Get-Item 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe').VersionInfo.ProductVersion"],
                capture_output=True, text=True, timeout=15
            )
            version = result.stdout.strip()
        except:
            logger.error("无法获取 Chrome 版本，自动下载失败")
            return None

    logger.info(f"Chrome 版本: {version}")
    major = version.split('.')[0]
    # 查询合适的 chromedriver 版本
    try:
        resp = requests.get(f"https://registry.npmmirror.com/-/binary/chromedriver/{version}/")
        if resp.status_code != 200:
            resp = requests.get(f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major}")
            latest_version = resp.text.strip()
            download_url = f"https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_win32.zip"
        else:
            download_url = f"https://registry.npmmirror.com/-/binary/chromedriver/{version}/chromedriver_win32.zip"
    except Exception as e:
        logger.error(f"查询 ChromeDriver 失败: {e}")
        return None

    try:
        logger.info(f"下载 ChromeDriver: {download_url[:70]}...")
        resp = requests.get(download_url, timeout=120, stream=True)
        if resp.status_code != 200:
            raise Exception("下载失败")
        zip_path = os.path.join(script_dir, "_chromedriver.zip")
        with open(zip_path, "wb") as f:
            for chunk in resp.iter_content(8192): f.write(chunk)
        with zipfile.ZipFile(zip_path, "r") as z:
            for name in z.namelist():
                if "chromedriver" in name and name.endswith(".exe"):
                    with z.open(name) as src, open(target, "wb") as dst:
                        dst.write(src.read())
                    break
        os.remove(zip_path)
        if os.path.isfile(target):
            logger.info(f"✓ chromedriver: {target}")
            return target
    except Exception as e:
        logger.warning(f"下载 ChromeDriver 失败: {e}")
    return None

def download_edgedriver():
    """自动下载 EdgeDriver"""
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "msedgedriver.exe")
    if os.path.isfile(target):
        return target
    version = None
    # 获取 Edge 版本
    for p in [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if os.path.isfile(p):
            try:
                result = subprocess.run(
                    ["powershell", "-Command", f"(Get-Item '{p}').VersionInfo.ProductVersion"],
                    capture_output=True, text=True, timeout=15
                )
                version = result.stdout.strip()
                break
            except: pass
    if not version:
        logger.error("无法获取 Edge 版本，自动下载失败")
        return None

    logger.info(f"Edge 版本: {version}")
    urls = [
        f"https://registry.npmmirror.com/-/binary/edgedriver/{version}/edgedriver_win64.zip",
        f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip",
    ]
    for url in urls:
        try:
            logger.info(f"下载 EdgeDriver: {url[:70]}...")
            resp = requests.get(url, timeout=120, stream=True)
            if resp.status_code != 200: continue
            zip_path = os.path.join(script_dir, "_edgedriver.zip")
            with open(zip_path, "wb") as f:
                for chunk in resp.iter_content(8192): f.write(chunk)
            with zipfile.ZipFile(zip_path, "r") as z:
                for name in z.namelist():
                    if "msedgedriver" in name and name.endswith(".exe"):
                        with z.open(name) as src, open(target, "wb") as dst:
                            dst.write(src.read())
                        break
            os.remove(zip_path)
            if os.path.isfile(target):
                logger.info(f"✓ msedgedriver: {target}")
                return target
        except Exception as e:
            logger.warning(f"下载失败: {e}")
    logger.error("自动下载 EdgeDriver 失败，请手动下载")
    return None

def detect_browser():
    """检测浏览器类型，返回 ('edge'/'chrome', 本地driver路径或None)"""
    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if not os.path.isfile(edge_exe):
        edge_exe = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    if os.path.isfile(edge_exe):
        return "edge", get_webdriver_path("edge") if os.path.isfile(get_webdriver_path("edge")) else None

    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if not os.path.isfile(chrome_exe):
        chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    if os.path.isfile(chrome_exe):
        return "chrome", get_webdriver_path("chrome") if os.path.isfile(get_webdriver_path("chrome")) else None

    # 检查 PATH
    if shutil.which("msedge"): return "edge", None
    if shutil.which("google-chrome") or shutil.which("chrome"): return "chrome", None
    return None, None

# ============================================================
# 配置管理
# ============================================================

DEFAULT_CONFIG = {
    "cookie": "",
    "date_start": "2024-01-01",
    "date_end": "2026-03-31",
    "mode": "full",
    "task_status": "not_started",   # not_started / in_progress / completed
    "list_progress": {
        "last_page": 0,
        "all_urls": []             # 收集到的所有 URL（去重前）
    },
    "detail_progress": {
        "visited_urls": [],        # 已处理的详情 URL
        "saved_count": 0
    },
    "last_update": ""
}

class ConfigManager:
    def __init__(self, config_path=None):
        if config_path is None:
            config_path = os.path.join(get_script_dir(), "crawl_config.json")
        self.path = config_path
        self.data = self.load()

    def load(self):
        if os.path.isfile(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                # 合并缺失字段
                for key, val in DEFAULT_CONFIG.items():
                    if key not in config:
                        config[key] = val
                return config
            except:
                pass
        return dict(DEFAULT_CONFIG)

    def save(self, data=None):
        if data:
            self.data = data
        self.data["last_update"] = datetime.now().isoformat()
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def update_progress(self, **kwargs):
        for k, v in kwargs.items():
            self.data[k] = v
        self.save()

    def is_task_completed(self):
        return self.data.get("task_status") == "completed"

    def is_task_in_progress(self):
        return self.data.get("task_status") == "in_progress"

    def mark_completed(self):
        self.data["task_status"] = "completed"
        # 清除详细进度以节省空间
        self.data["list_progress"]["all_urls"] = []
        self.data["list_progress"]["last_page"] = 0
        self.data["detail_progress"]["visited_urls"] = []
        self.save()

    def reset_for_new_task(self):
        """保留 cookie 和上次日期作为默认值，重置进度"""
        cookie = self.data.get("cookie", "")
        date_start = self.data.get("date_start", "2024-01-01")
        date_end = self.data.get("date_end", "2026-03-31")
        self.data = dict(DEFAULT_CONFIG)
        self.data["cookie"] = cookie
        self.data["date_start"] = date_start
        self.data["date_end"] = date_end
        self.save()

# ============================================================
# 列表页爬取
# ============================================================

class ListPageScraper:
    def __init__(self, driver):
        self.driver = driver

    def fetch_page(self, page, date_start=None, date_end=None):
        city_gbk = urllib.parse.quote(LIST_CITY.encode("gbk"))
        url = LIST_URL.format(city_gbk=city_gbk)
        params = {
            "auction_source": "0",
            "st_param": "-1",
            "auction_start_seg": "0",
            "page": str(page),
        }
        if date_start:
            params["auction_start_from"] = date_start
        if date_end:
            params["auction_start_to"] = date_end
        full_url = url + "?" + urllib.parse.urlencode(params)

        logger.info(f"列表页第 {page} 页")
        try:
            self.driver.get(full_url)
            WebDriverWait(self.driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(4)
        except TimeoutException:
            logger.warning("加载超时")
            return "timeout", []

        cur_url = self.driver.current_url
        if any(k in cur_url for k in ["deny", "punish", "login"]):
            logger.error("⚠️ 触发反爬或需要登录")
            return "blocked", []

        # JSON 提取
        items = []
        try:
            script_el = self.driver.find_element(By.ID, "sf-item-list-data")
            text = script_el.get_attribute("innerHTML").strip()
            if text:
                data = json.loads(text)
                items = data.get("data", [])
                if items:
                    return "ok", self._normalize(items)
        except: pass

        # 源码正则
        page_source = self.driver.page_source
        m = re.search(r'id="sf-item-list-data"[^>]*>(.*?)</script>', page_source, re.DOTALL)
        if m:
            try:
                data = json.loads(m.group(1).strip())
                items = data.get("data", [])
                if items:
                    return "ok", self._normalize(items)
            except: pass

        # HTML 解析
        try:
            links = self.driver.find_elements(By.CSS_SELECTOR,
                "a[href*='sf-item'], a[href*='susong-item'], a[href*='auction']")
            seen = set()
            for link in links:
                href = link.get_attribute("href") or ""
                if not href or "htm" not in href or href in seen: continue
                seen.add(href)
                title = link.text.strip()
                if not title or len(title) < 5: continue
                id_m = re.search(r"(\d{10,})\.htm", href)
                items.append({
                    "itemUrl": href,
                    "title": title,
                    "id": id_m.group(1) if id_m else "",
                })
        except Exception as e:
            logger.warning(f"HTML解析异常: {e}")

        return "ok", items

    def _normalize(self, raw):
        result = []
        for item in raw:
            url = item.get("itemUrl", "")
            if url and not url.startswith("http"):
                url = "https:" + url if url.startswith("//") else "https://sf.taobao.com" + url
            result.append({
                "itemUrl": url,
                "title": item.get("title", ""),
                "id": item.get("id", ""),
            })
        return result

# ============================================================
# 详情页爬取
# ============================================================

class DetailPageScraper:
    def __init__(self, driver):
        self.driver = driver

    def fetch_detail(self, url):
        result = {"url": url, "auction_type": "", "title": "", "status": "",
                  "is_success": False, "intro_table": {"table_rows": [], "raw": ""}, "attachments": []}
        try:
            self.driver.get(url)
            WebDriverWait(self.driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(3)
        except TimeoutException:
            return result

        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url:
            return "blocked", result
        if "login" in cur_url:
            return "need_login", result

        body_text = self.driver.find_element(By.TAG_NAME, "body").text
        result["title"] = self._extract_title()
        result["auction_type"] = self._extract_auction_type(body_text)
        status_info = self._extract_status(body_text)
        result["status"] = status_info["text"]
        result["is_success"] = status_info["success"]
        result["intro_table"] = self._extract_intro_table()
        result["attachments"] = self._extract_attachments()
        return "ok", result

    def _extract_title(self):
        for sel in ["h1", ".pm-main h1", ".item-title", ".title h1",
                    "[class*='title'] h1", "h1[class*='title']"]:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                t = el.text.strip()
                if t and len(t) > 5: return t
            except: pass
        t = self.driver.title
        if t:
            return t.split("-")[0].strip().split("—")[0].strip()
        return ""

    def _extract_auction_type(self, text):
        title = self._extract_title()
        for kw in ["一拍", "二拍", "三拍", "变卖"]:
            if kw in title: return kw
        m = re.search(r'(一拍|二拍|三拍|变卖)', text[:500])
        if m: return m.group(1)
        url = self.driver.current_url
        if "auction_type=" in url:
            m = re.search(r'auction_type=(\d)', url)
            if m:
                return {"1": "一拍", "2": "二拍", "3": "三拍", "5": "变卖"}.get(m.group(1), "")
        return ""

    def _extract_status(self, text):
        search = text[:2000]
        for kw in FAIL_KEYWORDS:
            if kw in search:
                for line in search.split('\n'):
                    if kw in line: return {"text": line.strip(), "success": False}
        for kw in SUCCESS_KEYWORDS:
            if kw in search:
                for line in search.split('\n'):
                    if kw in line: return {"text": line.strip(), "success": True}
        # CSS 兜底
        for sel in ["[class*='status']", "[class*='auction-status']", "[class*='state']", ".pm-status"]:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                t = el.text.strip()
                if t:
                    if any(k in t for k in FAIL_KEYWORDS): return {"text": t, "success": False}
                    if any(k in t for k in SUCCESS_KEYWORDS): return {"text": t, "success": True}
            except: pass
        return {"text": "未知", "success": False}

    def _extract_intro_table(self):
        table_rows = []
        raw_text = ""
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#J_desc > table"))
            )
            table = self.driver.find_element(By.CSS_SELECTOR, "#J_desc > table")
            table_rows = self._parse_table_to_grid(table)
            try:
                raw_text = self.driver.find_element(By.ID, "J_desc").text
            except:
                raw_text = self.driver.find_element(By.TAG_NAME, "body").text
            if table_rows:
                return {"table_rows": table_rows, "raw": raw_text}
        except TimeoutException:
            pass

        # 备选：所有表格
        try:
            tables = self.driver.find_elements(By.TAG_NAME, "table")
            for t in tables:
                if len(t.find_elements(By.TAG_NAME, "tr")) < 2: continue
                grid = self._parse_table_to_grid(t)
                if grid and len(grid) > 1:
                    table_rows = grid
                    break
            if not table_rows:
                table_rows = self._fallback_text_parse()
            if not raw_text:
                raw_text = self.driver.find_element(By.TAG_NAME, "body").text
        except Exception as e:
            logger.warning(f"备选提取异常: {e}")
        return {"table_rows": table_rows, "raw": raw_text}

    def _parse_table_to_grid(self, table):
        rows = table.find_elements(By.TAG_NAME, "tr")
        if not rows: return []
        max_cols, grid = 0, []
        for r in rows:
            cells = r.find_elements(By.XPATH, "./th | ./td")
            row_cells = []
            for c in cells:
                colspan = int(c.get_attribute("colspan") or 1)
                rowspan = int(c.get_attribute("rowspan") or 1)
                row_cells.append((c.text.strip(), colspan, rowspan))
            grid.append(row_cells)
            max_cols = max(max_cols, sum(c[1] for c in row_cells))
        nrows = len(grid)
        matrix = [[None]*max_cols for _ in range(nrows)]
        for r_idx, row_cells in enumerate(grid):
            c_idx = 0
            for txt, cs, rs in row_cells:
                while c_idx < max_cols and matrix[r_idx][c_idx] is not None:
                    c_idx += 1
                if c_idx >= max_cols: break
                for dr in range(rs):
                    for dc in range(cs):
                        nr, nc = r_idx+dr, c_idx+dc
                        if nr < nrows and nc < max_cols:
                            matrix[nr][nc] = txt if dr==0 and dc==0 else ""
                c_idx += cs
        return [row for row in matrix if any(c and c.strip() for c in row)]

    def _fallback_text_parse(self):
        body = self.driver.find_element(By.TAG_NAME, "body")
        page_text = body.text
        intro_start = -1
        for mk in ["标的物介绍", "拍卖标的基本情况", "标的物详情"]:
            idx = page_text.find(mk)
            if idx != -1:
                intro_start = idx + len(mk)
                break
        if intro_start == -1:
            return []
        intro = page_text[intro_start:]
        for em in ["竞买公告", "竞买须知", "拍卖规则", "附件", "相关附件",
                   "重要提示", "竞买人条件", "咨询时间", "拍卖余款"]:
            idx = intro.find(em)
            if idx != -1: intro = intro[:idx]
        table_data = {}
        lines = [l.strip() for l in intro.split('\n') if l.strip()]
        i = 0
        while i < len(lines):
            line = lines[i]
            if len(line) < 2: i+=1; continue
            m = re.match(r'^([^：:]+)[：:]\s*(.+)$', line)
            if m:
                table_data[m.group(1).strip()] = m.group(2).strip()
            elif i+1 < len(lines):
                nxt = lines[i+1]
                if len(line) < 20 and not re.match(r'^[\d.]+', line) and nxt and line != nxt:
                    table_data[line] = nxt
                    i += 1
            i += 1
        if table_data:
            return [["项目", "内容"]] + [[k, v] for k, v in table_data.items()]
        return []

    def _extract_attachments(self):
        atts = []
        try:
            selectors = ["a[href*='.pdf']", "a[href*='.doc']", "a[href*='.xls']",
                         "a[href*='.zip']", "a[href*='.rar']", "a[href*='.7z']",
                         "a[href*='attachment']", "a[href*='download']"]
            seen = set()
            for sel in selectors:
                for el in self.driver.find_elements(By.CSS_SELECTOR, sel):
                    href = el.get_attribute("href") or ""
                    name = el.text.strip()
                    if href and href not in seen:
                        seen.add(href)
                        if not name:
                            name = href.split("/")[-1].split("?")[0]
                        atts.append({"name": name, "url": href})
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            for name in re.findall(r'相关附件[：:]\s*(\S+\.(?:pdf|docx?|xlsx?|zip|rar))', body_text, re.I):
                if not any(name in a["name"] for a in atts):
                    atts.append({"name": name, "url": ""})
        except Exception as e:
            logger.warning(f"附件提取异常: {e}")
        return atts

# ============================================================
# 附件下载
# ============================================================

def download_attachment(url, save_path, driver=None):
    if not url: return False
    try:
        headers = {"User-Agent": random.choice(USER_AGENTS)}
        resp = requests.get(url, headers=headers, timeout=60, stream=True)
        if resp.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(8192): f.write(chunk)
            return True
    except: pass
    if driver:
        try:
            driver.get(url)
            time.sleep(5)
            return True
        except: pass
    return False

# ============================================================
# 主爬虫
# ============================================================

class SfTaobaoScraper:
    def __init__(self, config: ConfigManager, browser_type, driver_path):
        self.config = config
        self.browser_type = browser_type
        self.driver_path = driver_path
        self.driver = None
        self.list_scraper = None
        self.detail_scraper = None
        self.stats = {"total_urls": 0, "total_saved": 0, "total_skipped": 0, "total_failed": 0, "start_time": None}
        self.output_root = os.path.join(get_script_dir(), "output")
        self.log_handler = None
        self._all_urls_dedup = []

    def start(self):
        self.stats["start_time"] = time.time()
        if self.browser_type == "edge":
            opts = EdgeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            svc = EdgeService(self.driver_path)
            self.driver = webdriver.Edge(service=svc, options=opts)
        else:
            opts = ChromeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            svc = ChromeService(self.driver_path)
            self.driver = webdriver.Chrome(service=svc, options=opts)

        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        self.driver.get("https://sf.taobao.com/")

        # 尝试注入保存的 cookie
        saved_cookie = self.config.data.get("cookie", "")
        if saved_cookie:
            self._inject_cookie(saved_cookie)
            self.driver.refresh()
            time.sleep(2)

        print("\n请在浏览器中手动登录淘宝账号（若已登录可忽略）")
        while True:
            choice = input("登录成功？(Y/N): ").strip().upper()
            if choice == "Y":
                break
        # 保存当前 cookie
        cookies = self.driver.get_cookies()
        cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
        self.config.data["cookie"] = cookie_str
        self.config.save()

        self.list_scraper = ListPageScraper(self.driver)
        self.detail_scraper = DetailPageScraper(self.driver)
        logger.info("浏览器准备就绪")

    def _inject_cookie(self, cookie_str):
        for item in cookie_str.split(";"):
            item = item.strip()
            if "=" in item:
                n, v = item.split("=", 1)
                try:
                    self.driver.add_cookie({"name": n.strip(), "value": v.strip(), "domain": ".taobao.com", "path": "/"})
                except: pass

    def _setup_log_file(self, date_range_str):
        log_dir = os.path.join(self.output_root, date_range_str)
        safe_makedirs(log_dir)
        log_path = os.path.join(log_dir, f"crawl_{date_range_str}.log")
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setLevel(logging.INFO)
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(fh)
        self.log_handler = fh
        logger.info(f"日志文件: {log_path}")

    def _handle_blocked(self):
        print("\n⚠️ 检测到反爬验证，请在浏览器中手动完成验证，完成后输入 Y 继续")
        while True:
            c = input("是否已完成验证并继续？(Y/N): ").strip().upper()
            if c == "Y": return True
            if c == "N": return False

    def _deduplicate_urls(self, urls):
        seen = set()
        dedup = []
        for url in urls:
            m = re.search(r"(\d{10,})\.htm", url)
            uid = m.group(1) if m else url
            if uid not in seen:
                seen.add(uid)
                dedup.append(url)
        return dedup

    def _save_item(self, detail, date_range_str):
        title = detail.get("title", "未知")
        safe_title = sanitize_filename(title) or f"item_{int(time.time())}"
        item_dir = os.path.join(self.output_root, date_range_str, safe_title)
        safe_makedirs(item_dir)

        with open(os.path.join(item_dir, "url.txt"), "w", encoding="utf-8") as f:
            f.write(detail.get("url", ""))
        info = [
            f"拍卖类型: {detail.get('auction_type')}",
            f"拍卖物名称: {detail.get('title')}",
            f"拍卖状态: {detail.get('status')}",
            f"是否成交: {'是' if detail.get('is_success') else '否'}",
            f"详情页链接: {detail.get('url')}",
            f"爬取时间: {datetime.now():%Y-%m-%d %H:%M:%S}",
        ]
        with open(os.path.join(item_dir, "拍卖信息.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(info))

        intro = detail.get("intro_table", {})
        rows = intro.get("table_rows", [])
        if rows:
            with open(os.path.join(item_dir, "标的物介绍.csv"), "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
        raw = intro.get("raw", "")
        if raw:
            with open(os.path.join(item_dir, "标的物介绍.txt"), "w", encoding="utf-8") as f:
                f.write(raw)

        atts = detail.get("attachments", [])
        if atts:
            att_dir = os.path.join(item_dir, "附件")
            safe_makedirs(att_dir)
            for att in atts:
                name = sanitize_filename(att.get("name", "attachment"))
                if name and not os.path.exists(os.path.join(att_dir, name)):
                    download_attachment(att.get("url"), os.path.join(att_dir, name), self.driver)

        logger.info(f"✓ 已保存: {safe_title[:50]}")
        return item_dir

    def run(self, date_start, date_end, mode):
        date_range_str = f"{date_start}_to_{date_end}"
        self._setup_log_file(date_range_str)
        if mode == "full":
            self._run_full(date_start, date_end, date_range_str)
        else:
            self._run_detail_only(date_range_str)
        self._print_stats()

    def _run_full(self, date_start, date_end, date_range_str):
        config = self.config.data
        list_prog = config.get("list_progress", {})
        last_page = list_prog.get("last_page", 0)
        all_urls = list_prog.get("all_urls", [])
        start_page = last_page + 1 if last_page > 0 else 1

        logger.info("阶段1：收集拍品 URL")
        for page in range(start_page, MAX_PAGES + 1):
            status, items = self.list_scraper.fetch_page(page, date_start, date_end)
            if status == "blocked":
                if self._handle_blocked():
                    status, items = self.list_scraper.fetch_page(page, date_start, date_end)
                else:
                    break
            if status == "timeout":
                time.sleep(10)
                continue
            if not items:
                break

            for it in items:
                url = it.get("itemUrl")
                if url:
                    all_urls.append(url)
            # 保存进度
            list_prog["last_page"] = page
            list_prog["all_urls"] = all_urls
            self.config.save()
            self.stats["total_urls"] = len(all_urls)
            if len(items) < ITEMS_PER_PAGE:
                break
            random_delay()

        all_urls = self._deduplicate_urls(all_urls)
        logger.info(f"去重后 URL 数量: {len(all_urls)}")
        self._all_urls_dedup = all_urls
        # 进入详情处理前清空列表 URL 以免配置文件过大
        self.config.data["list_progress"]["all_urls"] = []
        self.config.save()
        self._process_details(all_urls, date_range_str)

    def _run_detail_only(self, date_range_str):
        url_file = os.path.join(get_script_dir(), "urls.txt")
        if not os.path.isfile(url_file):
            logger.error(f"urls.txt 不存在")
            return
        with open(url_file, "r", encoding="utf-8") as f:
            urls = [l.strip() for l in f if l.strip().startswith("http")]
        urls = self._deduplicate_urls(urls)
        logger.info(f"从 urls.txt 加载 {len(urls)} 个 URL")
        self._process_details(urls, date_range_str)

    def _process_details(self, urls, date_range_str):
        config = self.config.data
        detail_prog = config.get("detail_progress", {})
        visited = set(detail_prog.get("visited_urls", []))
        saved_count = detail_prog.get("saved_count", 0)
        self.stats["total_saved"] = saved_count

        logger.info(f"阶段2：处理详情页 ({len(urls)} 个)")
        consecutive_fail = 0

        for idx, url in enumerate(urls):
            if url in visited:
                continue
            if "sf_item" not in url and "auction" not in url and "susong" not in url:
                visited.add(url)
                continue

            logger.info(f"[{idx+1}/{len(urls)}] {url[:80]}")
            random_delay()

            result = self.detail_scraper.fetch_detail(url)
            if isinstance(result, tuple):
                status, detail = result
            else:
                status = "error"
                detail = result

            if status == "blocked":
                consecutive_fail += 1
                if consecutive_fail >= 3:
                    if self._handle_blocked():
                        consecutive_fail = 0
                        continue
                    else:
                        break
                visited.add(url)
                continue
            elif status == "need_login":
                if self._handle_blocked():
                    continue
                else:
                    break
            elif status == "error":
                visited.add(url)
                self.stats["total_failed"] += 1
                consecutive_fail += 1
                if consecutive_fail >= 10:
                    logger.error("连续10次失败，暂停")
                    break
                continue
            else:
                consecutive_fail = 0
                visited.add(url)

                title = detail.get("title", "")
                if any(kw in title for kw in NON_RESIDENTIAL):
                    logger.info(f"⏭ 跳过非住宅: {title[:40]}")
                    self.stats["total_skipped"] += 1
                elif not detail.get("is_success"):
                    logger.info(f"⏭ 跳过未成交: {detail.get('status')}")
                    self.stats["total_skipped"] += 1
                else:
                    self._save_item(detail, date_range_str)
                    self.stats["total_saved"] += 1

                # 保存详情进度
                detail_prog["visited_urls"] = list(visited)
                detail_prog["saved_count"] = self.stats["total_saved"]
                self.config.save()

        # 正常结束
        self.config.mark_completed()

    def _print_stats(self):
        elapsed = time.time() - self.stats["start_time"] if self.stats["start_time"] else 0
        print("\n" + "="*60)
        print(f"📊 统计: 总URL {self.stats['total_urls']}, 已保存 {self.stats['total_saved']}, "
              f"跳过 {self.stats['total_skipped']}, 失败 {self.stats['total_failed']}")
        print(f"⏱ 耗时 {elapsed/60:.1f} 分钟")
        print(f"📁 输出: {self.output_root}")

    def save_summary(self, date_range_str):
        summary_dir = os.path.join(self.output_root, date_range_str)
        if not os.path.isdir(summary_dir): return
        summary_file = os.path.join(summary_dir, "_汇总.csv")
        rows = []
        for name in sorted(os.listdir(summary_dir)):
            d = os.path.join(summary_dir, name)
            if not os.path.isdir(d) or name.startswith("_"): continue
            row = {"目录名": name}
            info_file = os.path.join(d, "拍卖信息.txt")
            if os.path.isfile(info_file):
                with open(info_file, "r", encoding="utf-8") as f:
                    for line in f:
                        if ":" in line:
                            k, v = line.split(":", 1)
                            row[k.strip()] = v.strip()
            intro_file = os.path.join(d, "标的物介绍.txt")
            if os.path.isfile(intro_file):
                try:
                    with open(intro_file, "r", encoding="utf-8") as f:
                        text = f.read(500).replace('\n', ' ').replace('\r', ' ')
                    row["标的物介绍摘要"] = text[:200]
                except: pass
            att_dir = os.path.join(d, "附件")
            if os.path.isdir(att_dir):
                row["附件数量"] = str(len(os.listdir(att_dir)))
            rows.append(row)
        if rows:
            keys = []
            for r in rows:
                for k in r:
                    if k not in keys: keys.append(k)
            with open(summary_file, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(rows)
            logger.info(f"汇总已保存: {summary_file} ({len(rows)} 条)")

    def stop(self):
        if self.log_handler:
            logger.removeHandler(self.log_handler)
            self.log_handler.close()
        if self.driver:
            try: self.driver.quit()
            except: pass

# ============================================================
# 入口
# ============================================================

def install_dependencies():
    print("正在安装依赖：selenium requests beautifulsoup4 ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "requests", "beautifulsoup4"])
        return True
    except Exception as e:
        print(f"安装失败: {e}")
        return False

def main():
    print("""
╔══════════════════════════════════════════════════════════════╗
║       阿里司法拍卖 - 法拍房数据爬取工具 v0.13                ║
╚══════════════════════════════════════════════════════════════╝
    """)
    # 1. 检查 selenium
    if not HAS_SELENIUM:
        print("❌ 未检测到 selenium")
        ans = input("是否自动安装依赖？(Y/N): ").strip().upper()
        if ans == "Y":
            if not install_dependencies():
                print("安装失败，请手动执行: pip install selenium requests beautifulsoup4")
                sys.exit(1)
            else:
                print("安装成功，请重新运行程序")
                sys.exit(0)
        else:
            sys.exit(1)

    # 2. 检测浏览器与 driver
    browser_type, driver_path = detect_browser()
    if not browser_type:
        print("❌ 未检测到 Edge 或 Chrome 浏览器")
        sys.exit(1)
    print(f"✓ 浏览器: {browser_type.upper()}")

    if not driver_path or not os.path.isfile(driver_path):
        print("WebDriver 缺失，正在自动下载...")
        if browser_type == "edge":
            driver_path = download_edgedriver()
        else:
            driver_path = download_chromedriver()
        if not driver_path:
            print("❌ 自动下载失败，请手动放置 WebDriver 到脚本目录")
            sys.exit(1)
        print(f"✓ WebDriver 已下载: {driver_path}")

    # 3. 配置管理
    config = ConfigManager()
    if config.is_task_in_progress():
        print("\n⚠️ 检测到未完成的任务")
        ans = input("是否恢复上次进度？(Y/N): ").strip().upper()
        if ans == "Y":
            # 沿用配置中的参数
            date_start = config.data["date_start"]
            date_end = config.data["date_end"]
            mode = config.data["mode"]
            print(f"恢复任务：{date_start} ~ {date_end}，模式：{mode}")
        else:
            config.reset_for_new_task()
            date_start = input("开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
            date_end = input("结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"
            mode = input("模式: 1=完整 2=仅详情页 (默认 1): ").strip()
            mode = "detail_only" if mode == "2" else "full"
            config.data["date_start"] = date_start
            config.data["date_end"] = date_end
            config.data["mode"] = mode
            config.data["task_status"] = "in_progress"
            config.save()
    elif config.is_task_completed():
        print("\n上次任务已完成")
        ans = input("是否开始新任务？(Y/N): ").strip().upper()
        if ans != "Y":
            sys.exit(0)
        config.reset_for_new_task()
        date_start = input("开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
        date_end = input("结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"
        mode = input("模式: 1=完整 2=仅详情页 (默认 1): ").strip()
        mode = "detail_only" if mode == "2" else "full"
        config.data["date_start"] = date_start
        config.data["date_end"] = date_end
        config.data["mode"] = mode
        config.data["task_status"] = "in_progress"
        config.save()
    else:
        # 首次运行
        date_start = input("开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
        date_end = input("结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"
        mode = input("模式: 1=完整 2=仅详情页 (默认 1): ").strip()
        mode = "detail_only" if mode == "2" else "full"
        config.data["date_start"] = date_start
        config.data["date_end"] = date_end
        config.data["mode"] = mode
        config.data["task_status"] = "in_progress"
        config.save()

    scraper = SfTaobaoScraper(config, browser_type, driver_path)
    date_range_str = f"{config.data['date_start']}_to_{config.data['date_end']}"
    try:
        scraper.start()
        scraper.run(config.data["date_start"], config.data["date_end"], config.data["mode"])
        scraper.save_summary(date_range_str)
    except KeyboardInterrupt:
        logger.info("用户中断，进度已保存")
        scraper.save_summary(date_range_str)
    except Exception as e:
        logger.error(f"异常退出: {e}")
        import traceback; traceback.print_exc()
        scraper.save_summary(date_range_str)
    finally:
        scraper.stop()
    print("\n✅ 程序结束")

if __name__ == "__main__":
    main()