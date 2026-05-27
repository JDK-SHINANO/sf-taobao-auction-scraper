#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 昆明市法拍房数据爬取工具
========================================
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
import webbrowser
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

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
# 配置
# ============================================================

KUNMING_DISTRICTS = {
    "五华区": "530102", "盘龙区": "530103", "官渡区": "530111",
    "西山区": "530112", "东川区": "530113", "呈贡区": "530114",
    "晋宁区": "530115", "富民县": "530124", "宜良县": "530125",
    "石林彝族自治县": "530126", "嵩明县": "530127",
    "禄劝彝族苗族自治县": "530128", "寻甸回族彝族自治县": "530129",
    "安宁市": "530181",
}

LIST_CITY = "昆明"
LIST_URL = "https://sf.taobao.com/list/50025969__2___{city_gbk}.htm"
ITEMS_PER_PAGE = 40
MAX_PAGES = 150
MIN_DELAY = 2.0
MAX_DELAY = 5.0

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
]

# 拍卖成功/失败关键词
SUCCESS_KEYWORDS = ["本场已结束"]
FAIL_KEYWORDS = ["本场已流拍，无人出价"]


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

def detect_browser():
    script_dir = get_script_dir()
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in edge_paths:
        if os.path.isfile(p):
            local = os.path.join(script_dir, "msedgedriver.exe")
            return "edge", local if os.path.isfile(local) else None
    for p in chrome_paths:
        if os.path.isfile(p):
            local = os.path.join(script_dir, "chromedriver.exe")
            return "chrome", local if os.path.isfile(local) else None
    if shutil.which("msedge"): return "edge", None
    if shutil.which("google-chrome") or shutil.which("chrome"): return "chrome", None
    return None, None

def get_edge_version():
    edge_exe = None
    for p in [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if os.path.isfile(p):
            edge_exe = p; break
    if not edge_exe: return None
    try:
        result = subprocess.run(
            ["powershell", "-Command", f"(Get-Item '{edge_exe}').VersionInfo.ProductVersion"],
            capture_output=True, text=True, timeout=15
        )
        ver = result.stdout.strip()
        if ver and "." in ver: return ver
    except: pass
    return None

def download_edgedriver():
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "msedgedriver.exe")
    if os.path.isfile(target):
        return target
    version = get_edge_version()
    if not version:
        logger.error("无法获取 Edge 版本，请手动下载 msedgedriver.exe")
        return None
    logger.info(f"Edge 版本: {version}")
    urls = [
        f"https://registry.npmmirror.com/-/binary/edgedriver/{version}/edgedriver_win64.zip",
        f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip",
    ]
    for url in urls:
        try:
            logger.info(f"下载 msedgedriver: {url[:70]}...")
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
    logger.error("自动下载失败，请手动下载 msedgedriver.exe")
    return None


# ============================================================
# Checkpoint
# ============================================================

class Checkpoint:
    def __init__(self, name="crawl"):
        self.file = os.path.join(get_script_dir(), f".checkpoint_{name}.json")
        self.data = self._load()

    def _load(self):
        if os.path.isfile(self.file):
            try:
                with open(self.file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except: pass
        return {"visited_urls": [], "last_page": 0, "saved_count": 0, "timestamp": ""}

    def save(self, visited_urls, last_page, saved_count):
        self.data = {
            "visited_urls": list(visited_urls),
            "last_page": last_page,
            "saved_count": saved_count,
            "timestamp": datetime.now().isoformat(),
        }
        with open(self.file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def clear(self):
        if os.path.isfile(self.file): os.remove(self.file)
        self.data = {"visited_urls": [], "last_page": 0, "saved_count": 0, "timestamp": ""}

    @property
    def visited_urls(self): return set(self.data.get("visited_urls", []))
    @property
    def last_page(self): return self.data.get("last_page", 0)
    @property
    def saved_count(self): return self.data.get("saved_count", 0)


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

        logger.info(f"列表页第 {page} 页: {full_url[:90]}...")

        try:
            self.driver.get(full_url)
            WebDriverWait(self.driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(4)
        except TimeoutException:
            logger.warning(f"列表页第 {page} 页加载超时")
            return "timeout", []

        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url or "login" in cur_url:
            logger.error("⚠️ 触发反爬验证或 Cookie 失效！")
            return "blocked", []

        # 方法1: JSON
        try:
            script_el = self.driver.find_element(By.ID, "sf-item-list-data")
            script_text = script_el.get_attribute("innerHTML").strip()
            if script_text:
                data = json.loads(script_text)
                items = data.get("data", [])
                if items:
                    logger.info(f"  → {len(items)} 条 (JSON)")
                    return "ok", self._normalize_items(items)
        except (NoSuchElementException, json.JSONDecodeError):
            pass

        # 方法2: 源码
        page_source = self.driver.page_source
        m = re.search(r'id="sf-item-list-data"[^>]*>(.*?)</script>', page_source, re.DOTALL)
        if m:
            try:
                data = json.loads(m.group(1).strip())
                items = data.get("data", [])
                if items:
                    logger.info(f"  → {len(items)} 条 (源码JSON)")
                    return "ok", self._normalize_items(items)
            except json.JSONDecodeError:
                pass

        # 方法3: 链接遍历
        items = []
        try:
            links = self.driver.find_elements(
                By.CSS_SELECTOR,
                "a[href*='sf-item'], a[href*='susong-item'], a[href*='auction']"
            )
            seen = set()
            for link in links:
                href = link.get_attribute("href") or ""
                if not href or "htm" not in href: continue
                if href in seen: continue
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

        if items:
            logger.info(f"  → {len(items)} 条 (HTML)")
        else:
            logger.warning(f"  → 未解析到数据")
        return "ok", items

    def _normalize_items(self, raw_items):
        result = []
        for item in raw_items:
            item_url = item.get("itemUrl", "")
            if item_url and not item_url.startswith("http"):
                item_url = "https:" + item_url if item_url.startswith("//") else "https://sf.taobao.com" + item_url
            result.append({
                "itemUrl": item_url,
                "title": item.get("title", ""),
                "id": item.get("id", ""),
                "status": item.get("status", "") or item.get("statusText", ""),
            })
        return result


# ============================================================
# 详情页爬取
# ============================================================

class DetailPageScraper:
    def __init__(self, driver):
        self.driver = driver

    def fetch_detail(self, url):
        result = {
            "url": url,
            "auction_type": "",
            "title": "",
            "status": "",
            "is_success": False,
            "intro_table": {"table_rows": [], "raw": ""},
            "attachments": [],
        }

        try:
            self.driver.get(url)
            WebDriverWait(self.driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(3)
        except TimeoutException:
            logger.warning(f"  超时: {url}")
            return result

        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url:
            logger.error("  ⚠️ 触发反爬验证")
            return "blocked", result
        if "login" in cur_url:
            logger.error("  ⚠️ Cookie 失效，需要重新登录")
            return "need_login", result

        page_text = self.driver.find_element(By.TAG_NAME, "body").text

        result["title"] = self._extract_title()
        result["auction_type"] = self._extract_auction_type(page_text)

        status_info = self._extract_status(page_text)
        result["status"] = status_info["text"]
        result["is_success"] = status_info["success"]

        intro = self._extract_intro_table()
        result["intro_table"] = intro

        result["attachments"] = self._extract_attachments()

        return "ok", result

    def _extract_title(self):
        selectors = [
            "h1", ".pm-main h1", ".item-title", ".title h1",
            "[class*='title'] h1", "h1[class*='title']",
        ]
        for sel in selectors:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                text = el.text.strip()
                if text and len(text) > 5:
                    return text
            except NoSuchElementException:
                continue
        title_tag = self.driver.title
        if title_tag:
            return title_tag.split("-")[0].strip().split("—")[0].strip()
        return ""

    def _extract_auction_type(self, page_text):
        title = self._extract_title()
        for kw in ["一拍", "二拍", "三拍", "变卖"]:
            if kw in title:
                return kw
        m = re.search(r'(一拍|二拍|三拍|变卖)', page_text[:500])
        if m:
            return m.group(1)
        cur_url = self.driver.current_url
        if "auction_type" in cur_url:
            m = re.search(r'auction_type=(\d)', cur_url)
            if m:
                type_map = {"1": "一拍", "2": "二拍", "3": "三拍", "5": "变卖"}
                return type_map.get(m.group(1), "")
        return ""

    def _extract_status(self, page_text):
        search_text = page_text[:2000]
        status_text = ""

        for kw in FAIL_KEYWORDS:
            if kw in search_text:
                for line in search_text.split('\n'):
                    if kw in line:
                        status_text = line.strip()
                        break
                return {"text": status_text, "success": False}

        for kw in SUCCESS_KEYWORDS:
            if kw in search_text:
                for line in search_text.split('\n'):
                    if kw in line:
                        status_text = line.strip()
                        break
                return {"text": status_text, "success": True}

        status_selectors = [
            "[class*='status']", "[class*='auction-status']",
            "[class*='state']", ".pm-status",
        ]
        for sel in status_selectors:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                text = el.text.strip()
                if text:
                    status_text = text
                    if any(kw in text for kw in FAIL_KEYWORDS):
                        return {"text": text, "success": False}
                    if any(kw in text for kw in SUCCESS_KEYWORDS):
                        return {"text": text, "success": True}
                    break
            except NoSuchElementException:
                continue

        if not status_text:
            logger.warning("  ⚠ 未找到拍卖状态，默认跳过")
            return {"text": "未知", "success": False}
        return {"text": status_text, "success": False}

    def _extract_intro_table(self):
        """
        优先通过 #J_desc > table 提取完整表格，支持多列与合并单元格。
        若未找到则回退到通用表格查找或文本解析。
        """
        table_rows = []
        raw_text = ""

        # 第一步：优先使用指定选择器
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#J_desc > table"))
            )
            table_elem = self.driver.find_element(By.CSS_SELECTOR, "#J_desc > table")
            table_rows = self._parse_table_to_grid(table_elem)
            # 获取 raw text
            try:
                desc_elem = self.driver.find_element(By.ID, "J_desc")
                raw_text = desc_elem.text
            except:
                raw_text = self.driver.find_element(By.TAG_NAME, "body").text
            if table_rows:
                return {"table_rows": table_rows, "raw": raw_text}
        except TimeoutException:
            logger.debug("未找到 #J_desc > table，尝试备选方案")

        # 第二步：备选方案：遍历页面所有表格或文本解析
        try:
            tables = self.driver.find_elements(By.TAG_NAME, "table")
            for table in tables:
                rows = table.find_elements(By.TAG_NAME, "tr")
                if len(rows) < 2:
                    continue
                grid = self._parse_table_to_grid(table)
                if grid and len(grid) > 1:
                    table_rows = grid
                    break
            # 如果仍未找到，尝试文本解析
            if not table_rows:
                logger.debug("HTML表格未找到有效数据，启用文本解析")
                table_rows = self._fallback_text_parse_to_rows()
            # 补充 raw text
            if not raw_text:
                raw_text = self.driver.find_element(By.TAG_NAME, "body").text
                # 尝试截取标的物介绍部分
                intro_start = -1
                for marker in ["标的物介绍", "拍卖标的基本情况", "标的物详情"]:
                    idx = raw_text.find(marker)
                    if idx != -1:
                        intro_start = idx
                        break
                if intro_start != -1:
                    intro_text = raw_text[intro_start:]
                    for end_marker in ["竞买公告", "竞买须知", "拍卖规则", "附件", "相关附件",
                                       "重要提示", "竞买人条件", "咨询时间", "拍卖余款"]:
                        end_idx = intro_text.find(end_marker)
                        if end_idx != -1:
                            intro_text = intro_text[:end_idx]
                            break
                    raw_text = intro_text.strip()
        except Exception as e:
            logger.warning(f"备选提取异常: {e}")

        return {"table_rows": table_rows, "raw": raw_text}

    def _parse_table_to_grid(self, table_element):
        """
        将 Selenium table 元素解析为二维网格，处理 colspan/rowspan。
        返回 list of lists，第一行为表头（如果存在 th）。
        """
        rows_el = table_element.find_elements(By.TAG_NAME, "tr")
        if not rows_el:
            return []

        max_cols = 0
        grid = []

        for row_el in rows_el:
            cells = row_el.find_elements(By.XPATH, "./th | ./td")
            row_cells = []
            for cell in cells:
                colspan = int(cell.get_attribute("colspan") or 1)
                rowspan = int(cell.get_attribute("rowspan") or 1)
                row_cells.append((cell.text.strip(), colspan, rowspan))
            grid.append(row_cells)
            cols_in_row = sum(c[1] for c in row_cells)
            if cols_in_row > max_cols:
                max_cols = cols_in_row

        rows_count = len(grid)
        total_rows = rows_count
        matrix = [[None] * max_cols for _ in range(total_rows)]

        for r_idx, row_cells in enumerate(grid):
            c_idx = 0
            for cell_text, colspan, rowspan in row_cells:
                while c_idx < max_cols and matrix[r_idx][c_idx] is not None:
                    c_idx += 1
                if c_idx >= max_cols:
                    break

                for dr in range(rowspan):
                    for dc in range(colspan):
                        nr = r_idx + dr
                        nc = c_idx + dc
                        if nr < total_rows and nc < max_cols:
                            matrix[nr][nc] = cell_text if dr == 0 and dc == 0 else ""

                c_idx += colspan

        cleaned = []
        for row in matrix:
            if any(cell and cell.strip() for cell in row):
                cleaned.append(row)
        return cleaned

    def _fallback_text_parse_to_rows(self):
        """备选文本解析，转换为两列格式"""
        table_data = {}
        try:
            body = self.driver.find_element(By.TAG_NAME, "body")
            page_text = body.text
            intro_start = -1
            for marker in ["标的物介绍", "拍卖标的基本情况", "标的物详情"]:
                idx = page_text.find(marker)
                if idx != -1:
                    intro_start = idx + len(marker)
                    break
            if intro_start != -1:
                intro_text = page_text[intro_start:]
                for end_marker in ["竞买公告", "竞买须知", "拍卖规则", "附件", "相关附件",
                                   "重要提示", "竞买人条件", "咨询时间", "拍卖余款"]:
                    end_idx = intro_text.find(end_marker)
                    if end_idx != -1:
                        intro_text = intro_text[:end_idx]
                        break

                lines = [l.strip() for l in intro_text.split('\n') if l.strip()]
                i = 0
                while i < len(lines):
                    line = lines[i]
                    if len(line) < 2:
                        i += 1
                        continue
                    m = re.match(r'^([^：:]+)[：:]\s*(.+)$', line)
                    if m:
                        key = m.group(1).strip()
                        val = m.group(2).strip()
                        if key and val:
                            table_data[key] = val
                    elif i + 1 < len(lines):
                        next_line = lines[i + 1]
                        if (len(line) < 20 and not re.match(r'^[\d.]+', line)
                            and len(next_line) > 0 and line != next_line):
                            table_data[line] = next_line
                            i += 1
                    i += 1
        except Exception as e:
            logger.warning(f"文本解析出错: {e}")

        if table_data:
            rows = [["项目", "内容"]]
            for k, v in table_data.items():
                rows.append([k, v])
            return rows
        return []

    def _extract_attachments(self):
        attachments = []
        try:
            selectors = [
                "a[href*='.pdf']", "a[href*='.doc']", "a[href*='.xls']",
                "a[href*='.zip']", "a[href*='.rar']", "a[href*='.7z']",
                "a[href*='attachment']", "a[href*='download']",
            ]
            seen_urls = set()
            for sel in selectors:
                for el in self.driver.find_elements(By.CSS_SELECTOR, sel):
                    href = el.get_attribute("href") or ""
                    name = el.text.strip()
                    if href and href not in seen_urls:
                        seen_urls.add(href)
                        if not name:
                            name = href.split("/")[-1].split("?")[0]
                        attachments.append({"name": name, "url": href})

            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            att_m = re.findall(r'相关附件[：:]\s*(\S+\.(?:pdf|docx?|xlsx?|zip|rar))', body_text, re.I)
            for name in att_m:
                if not any(name in a["name"] for a in attachments):
                    attachments.append({"name": name, "url": ""})
        except Exception as e:
            logger.warning(f"提取附件异常: {e}")
        return attachments


# ============================================================
# 附件下载器
# ============================================================

def download_attachment(url, save_path, driver=None):
    if not url:
        logger.warning(f"附件 URL 为空，跳过: {save_path}")
        return False

    try:
        headers = {"User-Agent": random.choice(USER_AGENTS)}
        resp = requests.get(url, headers=headers, timeout=60, stream=True)
        if resp.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(8192):
                    f.write(chunk)
            logger.info(f"✓ 附件已下载: {os.path.basename(save_path)}")
            return True
    except Exception:
        pass

    if driver:
        try:
            driver.get(url)
            time.sleep(5)
            logger.info(f"✓ 附件已触发下载: {os.path.basename(save_path)}")
            return True
        except Exception as e:
            logger.warning(f"附件下载失败: {e}")

    return False


# ============================================================
# 主爬虫
# ============================================================

class SfTaobaoScraper:
    def __init__(self, cookie, browser_type="edge", driver_path=None):
        self.cookie = cookie
        self.browser_type = browser_type
        self.driver_path = driver_path
        self.driver = None
        self.list_scraper = None
        self.detail_scraper = None
        self.checkpoint_list = Checkpoint("list")
        self.checkpoint_detail = Checkpoint("detail")
        self.stats = {
            "total_urls": 0,
            "total_saved": 0,
            "total_skipped": 0,
            "total_failed": 0,
            "start_time": None,
        }
        self.output_root = os.path.join(get_script_dir(), "output")

    def start(self):
        self.stats["start_time"] = time.time()

        if self.browser_type == "edge":
            opts = EdgeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            opts.add_argument("--window-size=1920,1080")
            svc = EdgeService(self.driver_path) if self.driver_path and os.path.isfile(self.driver_path) else EdgeService()
            self.driver = webdriver.Edge(service=svc, options=opts)
        else:
            opts = ChromeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            opts.add_argument("--window-size=1920,1080")
            svc = ChromeService(self.driver_path) if self.driver_path and os.path.isfile(self.driver_path) else ChromeService()
            self.driver = webdriver.Chrome(service=svc, options=opts)

        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })

        self._inject_cookie(self.cookie)

        self.list_scraper = ListPageScraper(self.driver)
        self.detail_scraper = DetailPageScraper(self.driver)
        logger.info(f"浏览器已启动 ({self.browser_type.upper()})")

    def stop(self):
        if self.driver:
            try: self.driver.quit()
            except: pass

    def _inject_cookie(self, cookie_str):
        self.driver.get("https://sf.taobao.com/")
        time.sleep(2)
        for item in cookie_str.split(";"):
            item = item.strip()
            if "=" in item:
                n, v = item.split("=", 1)
                try:
                    self.driver.add_cookie({
                        "name": n.strip(), "value": v.strip(),
                        "domain": ".taobao.com", "path": "/"
                    })
                except: pass
        self.driver.refresh()
        time.sleep(2)

    def _handle_blocked(self, reason="blocked"):
        print("\n" + "=" * 60)
        if reason == "blocked":
            print("⚠️  检测到反爬验证（验证码/滑块）！")
            print("正在打开系统浏览器，请手动完成验证...")
            webbrowser.open("https://sf.taobao.com/")
            print("\n完成验证后，请：")
            print("  1. 在系统浏览器中登录淘宝")
            print("  2. 按 F12 → Network → 刷新页面")
            print("  3. 复制任意请求的 Cookie 值")
        else:
            print("⚠️  Cookie 已失效，需要重新输入！")
            print("请：")
            print("  1. 用浏览器打开 https://sf.taobao.com/ 并登录")
            print("  2. 按 F12 → Network → 刷新页面")
            print("  3. 复制任意请求的 Cookie 值")

        print("\n" + "-" * 60)
        new_cookie = input("请粘贴新 Cookie（直接回车放弃）: ").strip()
        if not new_cookie:
            print("放弃，继续处理已采集的数据...")
            return False

        self.cookie = new_cookie
        try:
            self._inject_cookie(new_cookie)
            logger.info("✓ Cookie 已更新，继续爬取")
            return True
        except Exception as e:
            logger.error(f"Cookie 注入失败: {e}")
            return False

    def _save_item(self, detail, date_range_str):
        title = detail.get("title", "未知")
        safe_title = sanitize_filename(title)
        if not safe_title:
            safe_title = f"item_{int(time.time())}"

        item_dir = os.path.join(self.output_root, date_range_str, safe_title)
        safe_makedirs(item_dir)

        # 1. URL
        with open(os.path.join(item_dir, "url.txt"), "w", encoding="utf-8") as f:
            f.write(detail.get("url", ""))

        # 2. 拍卖信息
        info_lines = [
            f"拍卖类型: {detail.get('auction_type', '')}",
            f"拍卖物名称: {detail.get('title', '')}",
            f"拍卖状态: {detail.get('status', '')}",
            f"是否成交: {'是' if detail.get('is_success') else '否'}",
            f"详情页链接: {detail.get('url', '')}",
            f"爬取时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        ]
        with open(os.path.join(item_dir, "拍卖信息.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(info_lines))

        # 3. 标的物介绍 CSV（多列）
        intro = detail.get("intro_table", {})
        table_rows = intro.get("table_rows", [])
        if table_rows:
            with open(os.path.join(item_dir, "标的物介绍.csv"), "w",
                      newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                for row in table_rows:
                    writer.writerow(row)

        # 4. 标的物介绍原始文本
        raw_text = intro.get("raw", "")
        if raw_text:
            with open(os.path.join(item_dir, "标的物介绍.txt"), "w", encoding="utf-8") as f:
                f.write(raw_text)

        # 5. 附件下载
        attachments = detail.get("attachments", [])
        if attachments:
            att_dir = os.path.join(item_dir, "附件")
            safe_makedirs(att_dir)
            for att in attachments:
                att_name = sanitize_filename(att.get("name", "attachment"))
                att_url = att.get("url", "")
                if att_name:
                    save_path = os.path.join(att_dir, att_name)
                    if not os.path.exists(save_path):
                        download_attachment(att_url, save_path, self.driver)

        logger.info(f"✓ 已保存: {safe_title[:50]}")
        return item_dir

    def run(self, date_start="2024-01-01", date_end="2026-03-31", mode="full"):
        date_range_str = f"{date_start}_to_{date_end}"
        if mode == "full":
            self._run_full(date_start, date_end, date_range_str)
        elif mode == "detail_only":
            self._run_detail_only(date_range_str)
        self._print_stats()

    def _run_full(self, date_start, date_end, date_range_str):
        cp_urls = self.checkpoint_list.visited_urls
        cp_page = self.checkpoint_list.last_page
        start_page = 1

        if cp_urls:
            print(f"\n🔄 发现列表页断点：上次爬到第 {cp_page} 页，已 {len(cp_urls)} 条")
            choice = input("  继续上次(y) / 重新开始(n) [默认y]: ").strip().lower()
            if choice == "n":
                self.checkpoint_list.clear()
                cp_urls = set()
            else:
                start_page = cp_page + 1

        all_urls = list(cp_urls)

        logger.info("=" * 60)
        logger.info(f"阶段1: 收集拍品 URL")
        logger.info(f"日期: {date_start} ~ {date_end}")
        logger.info("=" * 60)

        for page in range(start_page, MAX_PAGES + 1):
            status, items = self.list_scraper.fetch_page(page, date_start, date_end)

            if status == "blocked":
                if self._handle_blocked("blocked"):
                    status, items = self.list_scraper.fetch_page(page, date_start, date_end)
                    if status != "ok":
                        logger.error("重试失败，停止")
                        break
                else:
                    break
            elif status == "timeout":
                time.sleep(10)
                continue

            if not items:
                logger.info("列表页无数据，停止翻页")
                break

            for item in items:
                url = item.get("itemUrl", "")
                if url and url not in cp_urls:
                    all_urls.append(url)
                    cp_urls.add(url)

            self.stats["total_urls"] = len(all_urls)
            self.checkpoint_list.save(cp_urls, page, len(all_urls))

            if len(items) < ITEMS_PER_PAGE:
                break
            random_delay()

        logger.info(f"\n共收集到 {len(all_urls)} 个拍品 URL")
        self._process_details(all_urls, date_range_str)

    def _run_detail_only(self, date_range_str):
        url_file = os.path.join(get_script_dir(), "urls.txt")
        if not os.path.isfile(url_file):
            logger.error(f"未找到 URL 列表文件: {url_file}")
            return

        with open(url_file, "r", encoding="utf-8") as f:
            all_urls = [l.strip() for l in f if l.strip() and l.strip().startswith("http")]

        logger.info(f"从 urls.txt 读取到 {len(all_urls)} 个 URL")
        self._process_details(all_urls, date_range_str)

    def _process_details(self, all_urls, date_range_str):
        cp_visited = self.checkpoint_detail.visited_urls
        cp_saved = self.checkpoint_detail.saved_count

        if cp_visited:
            print(f"\n🔄 发现详情页断点：已处理 {len(cp_visited)} 个，已保存 {cp_saved} 个")
            choice = input("  继续上次(y) / 重新开始(n) [默认y]: ").strip().lower()
            if choice == "n":
                self.checkpoint_detail.clear()
                cp_visited = set()
                cp_saved = 0

        self.stats["total_saved"] = cp_saved

        logger.info("\n" + "=" * 60)
        logger.info(f"阶段2: 处理详情页 ({len(all_urls)} 个)")
        logger.info("=" * 60)

        consecutive_failures = 0

        # 修改后的非住宅关键词（不含车位、车库）
        NON_RESIDENTIAL = ["商铺", "门面", "门市", "写字楼",
                           "办公楼", "商业", "土地", "厂房", "仓库", "在建工程",
                           "机器设备", "车辆", "股权", "债权"]

        for idx, url in enumerate(all_urls):
            if url in cp_visited:
                continue

            logger.info(f"\n[{idx+1}/{len(all_urls)}] {url[:80]}...")

            if "sf_item" not in url and "auction" not in url and "susong" not in url:
                logger.info(f"  跳过（非详情页URL）")
                cp_visited.add(url)
                continue

            random_delay()

            detail_result = self.detail_scraper.fetch_detail(url)

            if isinstance(detail_result, tuple):
                status, detail = detail_result
            else:
                status = "error"
                detail = detail_result

            if status == "blocked":
                consecutive_failures += 1
                if consecutive_failures >= 3:
                    logger.error("连续3次被拦截，暂停处理")
                    if self._handle_blocked("blocked"):
                        consecutive_failures = 0
                        continue
                    else:
                        break
                cp_visited.add(url)
                continue
            elif status == "need_login":
                if self._handle_blocked("need_login"):
                    detail_result = self.detail_scraper.fetch_detail(url)
                    if isinstance(detail_result, tuple):
                        status, detail = detail_result
                    else:
                        cp_visited.add(url)
                        continue
                else:
                    break
            elif status == "error":
                cp_visited.add(url)
                self.stats["total_failed"] += 1
                consecutive_failures += 1
                if consecutive_failures >= 10:
                    logger.error("连续10次失败，暂停")
                    break
                continue

            consecutive_failures = 0
            cp_visited.add(url)

            # 非住宅过滤（去掉了车位/车库）
            title = detail.get("title", "")
            is_non_residential = any(kw in title for kw in NON_RESIDENTIAL)
            if is_non_residential:
                logger.info(f"  ⏭ 跳过（非住宅）: {title[:40]}")
                self.stats["total_skipped"] += 1
                self.checkpoint_detail.save(cp_visited, 0, self.stats["total_saved"])
                continue

            if not detail.get("is_success"):
                logger.info(f"  ⏭ 跳过（未成交）: {detail.get('status', '未知状态')}")
                self.stats["total_skipped"] += 1
                self.checkpoint_detail.save(cp_visited, 0, self.stats["total_saved"])
                continue

            self._save_item(detail, date_range_str)
            self.stats["total_saved"] += 1
            self.checkpoint_detail.save(cp_visited, 0, self.stats["total_saved"])

        self.checkpoint_list.clear()
        self.checkpoint_detail.clear()

    def _print_stats(self):
        elapsed = time.time() - self.stats["start_time"] if self.stats["start_time"] else 0
        print("\n" + "=" * 60)
        print("📊 统计:")
        print(f"   总URL数:   {self.stats['total_urls']}")
        print(f"   已保存:    {self.stats['total_saved']} (成功拍卖)")
        print(f"   已跳过:    {self.stats['total_skipped']} (未成交)")
        print(f"   失败:      {self.stats['total_failed']}")
        print(f"   耗时:      {elapsed/60:.1f} 分钟")
        print(f"   输出目录:  {self.output_root}")
        print("=" * 60)

    def save_summary(self, date_range_str):
        summary_dir = os.path.join(self.output_root, date_range_str)
        if not os.path.isdir(summary_dir):
            return

        summary_file = os.path.join(summary_dir, "_汇总.csv")
        rows = []

        for item_name in sorted(os.listdir(summary_dir)):
            item_dir = os.path.join(summary_dir, item_name)
            if not os.path.isdir(item_dir) or item_name.startswith("_"):
                continue

            row = {"目录名": item_name}

            info_file = os.path.join(item_dir, "拍卖信息.txt")
            if os.path.isfile(info_file):
                with open(info_file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if ":" in line:
                            k, v = line.split(":", 1)
                            row[k.strip()] = v.strip()

            intro_txt_file = os.path.join(item_dir, "标的物介绍.txt")
            if os.path.isfile(intro_txt_file):
                try:
                    with open(intro_txt_file, "r", encoding="utf-8") as f:
                        intro_raw = f.read(500)
                    intro_clean = intro_raw.replace('\n', ' ').replace('\r', ' ')
                    row["标的物介绍摘要"] = intro_clean[:200]
                except:
                    pass

            att_dir = os.path.join(item_dir, "附件")
            if os.path.isdir(att_dir):
                att_count = len(os.listdir(att_dir))
                row["附件数量"] = str(att_count)

            rows.append(row)

        if rows:
            all_keys = []
            seen = set()
            for row in rows:
                for k in row.keys():
                    if k not in seen:
                        all_keys.append(k)
                        seen.add(k)

            with open(summary_file, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=all_keys, extrasaction="ignore")
                writer.writeheader()
                for row in rows:
                    writer.writerow(row)

            logger.info(f"\n汇总已保存: {summary_file} ({len(rows)} 条)")


# ============================================================
# 入口
# ============================================================

def main():
    print("""

╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║           阿里司法拍卖 - 法拍房数据爬取工具 (V0.10)          ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝

————————
 _ ___  _  _     ____ _  _ _ _  _ ____ _  _ ____ 
 | |  \ |_/      [__  |__| | |\ | |__| |\ | |  | 
_| |__/ | \_ ___ ___] |  | | | \| |  | | \| |__| 
                        
————————
版本更新日志:
    状态判定仅用明确关键词                                    
    标的物介绍优先抓取 #J_desc > table（支持多列/合并）
————————
准备工作：
    1.用浏览器打开 https://sf.taobao.com/ 并登录
    2.F12 → Network → 刷新 → 复制 Cookie
————————
""")

    if not HAS_SELENIUM:
        print("❌ 未安装 selenium！请运行: pip install selenium")
        sys.exit(1)

    browser_type, local_driver = detect_browser()
    if browser_type:
        print(f"🔍 检测到: {browser_type.upper()}")
        if local_driver:
            print(f"   Driver: {local_driver}")
        else:
            print(f"   Driver: 未找到，将自动下载")
    else:
        print("⚠️  未检测到 Edge/Chrome")

    cookie = input("\n请粘贴 Cookie: ").strip()
    if not cookie:
        print("Cookie 不能为空"); sys.exit(1)

    print("\n📅 日期范围：")
    date_start = input("  开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
    date_end = input("  结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"

    print("\n🔧 运行模式：")
    print("  1 = 完整模式（列表页 + 详情页）")
    print("  2 = 仅详情页（需要同目录下有 urls.txt）")
    mode_choice = input("  选择 (默认 1): ").strip()
    mode = "detail_only" if mode_choice == "2" else "full"

    driver_path = local_driver
    if not driver_path:
        if browser_type == "edge":
            driver_path = download_edgedriver()
            if not driver_path:
                print("❌ msedgedriver 下载失败"); sys.exit(1)
        else:
            dp = input("  chromedriver 路径: ").strip()
            if dp: driver_path = dp

    print(f"\n{'='*60}\n开始爬取...\n{'='*60}\n")
    scraper = SfTaobaoScraper(cookie, browser_type=browser_type or "edge", driver_path=driver_path)

    try:
        scraper.start()
        scraper.run(date_start, date_end, mode)
        date_range_str = f"{date_start}_to_{date_end}"
        scraper.save_summary(date_range_str)
    except KeyboardInterrupt:
        logger.info("\n中断，保存进度...")
        scraper.save_summary(f"{date_start}_to_{date_end}")
    except Exception as e:
        logger.error(f"错误: {e}")
        import traceback; traceback.print_exc()
    finally:
        scraper.stop()

    print("\n✅ 运行结束")


if __name__ == "__main__":
    main()
