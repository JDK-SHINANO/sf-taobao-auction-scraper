#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 昆明市法拍房数据爬取工具 (第六版)
========================================
全新策略：
  - 从列表页获取所有拍品 URL
  - 逐个打开详情页，用 Selenium 提取：
    1. 拍卖类型（一拍/二拍/变卖）
    2. 拍卖物名称
    3. 拍卖状态（用"已结束"判断成交，筛掉流拍/中止/撤回）
    4. 标的物介绍（整个表格，key-value 保存）
    5. 附件下载
  - 输出目录：脚本同目录 / 日期范围 / 拍卖物名称 /
      ├── url.txt
      ├── 拍卖信息.txt
      ├── 标的物介绍.csv
      └── 附件/ (下载的文件)
  - Cookie 过期/反爬拦截 → 呼出系统浏览器人工验证 → 继续执行
  - 断点续爬（checkpoint）

运行环境：Windows + Python 3.11 + Edge/Chrome
依赖：pip install selenium requests beautifulsoup4 lxml openpyxl
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
LIST_URL = "https://sf.taobao.com/list/0__2___{city_gbk}.htm"
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

# 拍卖成功的关键词（页面上出现这些表示已成交/已结束）
SUCCESS_KEYWORDS = ["已结束", "已成交", "成交确认", "竞价成功"]
# 拍卖失败的关键词
FAIL_KEYWORDS = ["流拍", "中止", "撤回", "已撤销", "已中止", "无人出价", "暂缓"]


# ============================================================
# 工具函数
# ============================================================

def get_script_dir():
    return str(Path(__file__).parent.resolve())

def random_delay(low=None, high=None):
    time.sleep(random.uniform(low or MIN_DELAY, high or MAX_DELAY))

def sanitize_filename(name):
    """清理文件名中的非法字符"""
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:200]  # 限制长度

def safe_makedirs(path):
    os.makedirs(path, exist_ok=True)

def detect_browser():
    """检测系统浏览器"""
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
# Checkpoint（断点续爬）
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
# 列表页爬取（Selenium）— 获取所有拍品 URL
# ============================================================

class ListPageScraper:
    """从列表页获取所有拍品的 URL 和基本信息"""

    def __init__(self, driver):
        self.driver = driver

    def fetch_page(self, page, date_start=None, date_end=None):
        """获取列表页的一页数据，返回 [{itemUrl, title, id, status}]"""
        city_gbk = urllib.parse.quote(LIST_CITY.encode("gbk"))
        url = LIST_URL.format(city_gbk=city_gbk)
        params = {"auction_source": "0", "st_param": "-1", "page": str(page)}
        if date_start: params["auctionStartFrom"] = date_start
        if date_end: params["auctionStartTo"] = date_end
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

        # 检查是否被拦截
        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url or "login" in cur_url:
            logger.error("⚠️ 触发反爬验证或 Cookie 失效！")
            return "blocked", []

        # 方法1: 从 script#sf-item-list-data 提取 JSON
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

        # 方法2: 从页面源码提取
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

        # 方法3: 从页面链接解析
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
        """统一列表页数据格式"""
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
# 详情页爬取（Selenium）— 全新策略
# ============================================================

class DetailPageScraper:
    """
    从详情页提取：
    1. 拍卖类型（一拍/二拍/变卖）
    2. 拍卖物名称
    3. 拍卖状态（判断是否成交）
    4. 标的物介绍（整个表格）
    5. 附件链接
    """

    def __init__(self, driver):
        self.driver = driver

    def fetch_detail(self, url):
        """
        返回 dict:
        {
            "url": str,
            "auction_type": str,       # 一拍/二拍/变卖
            "title": str,              # 拍卖物名称
            "status": str,             # 拍卖状态文本
            "is_success": bool,        # 是否成交
            "intro_table": dict,       # 标的物介绍 key→value
            "intro_raw": str,          # 标的物介绍原始文本
            "attachments": list,       # [{"name": str, "url": str}]
        }
        """
        result = {
            "url": url,
            "auction_type": "",
            "title": "",
            "status": "",
            "is_success": False,
            "intro_table": {},
            "intro_raw": "",
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

        # 检查是否被拦截
        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url:
            logger.error("  ⚠️ 触发反爬验证")
            return "blocked", result
        if "login" in cur_url:
            logger.error("  ⚠️ Cookie 失效，需要重新登录")
            return "need_login", result

        page_text = self.driver.find_element(By.TAG_NAME, "body").text

        # --- 1. 拍卖物名称 ---
        result["title"] = self._extract_title()

        # --- 2. 拍卖类型 ---
        result["auction_type"] = self._extract_auction_type(page_text)

        # --- 3. 拍卖状态 ---
        status_info = self._extract_status(page_text)
        result["status"] = status_info["text"]
        result["is_success"] = status_info["success"]

        # --- 4. 标的物介绍表格 ---
        intro = self._extract_intro_table()
        result["intro_table"] = intro["table"]
        result["intro_raw"] = intro["raw"]

        # --- 5. 附件 ---
        result["attachments"] = self._extract_attachments()

        return "ok", result

    def _extract_title(self):
        """提取拍卖物名称"""
        # 尝试多种选择器
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

        # 从 title 标签提取
        title_tag = self.driver.title
        if title_tag:
            # 通常格式："拍卖物名称 - 淘宝司法拍卖"
            return title_tag.split("-")[0].strip().split("—")[0].strip()
        return ""

    def _extract_auction_type(self, page_text):
        """提取拍卖类型：一拍/二拍/变卖"""
        # 从标题中提取
        title = self._extract_title()
        for kw in ["一拍", "二拍", "三拍", "变卖"]:
            if kw in title:
                return kw

        # 从页面文本中提取（通常在顶部区域）
        m = re.search(r'(一拍|二拍|三拍|变卖)', page_text[:500])
        if m:
            return m.group(1)

        # 从 URL 参数推断
        cur_url = self.driver.current_url
        if "auction_type" in cur_url:
            m = re.search(r'auction_type=(\d)', cur_url)
            if m:
                type_map = {"1": "一拍", "2": "二拍", "3": "三拍", "5": "变卖"}
                return type_map.get(m.group(1), "")

        return ""

    def _extract_status(self, page_text):
        """
        提取拍卖状态，判断是否成交。
        返回 {"text": str, "success": bool}
        """
        status_text = ""

        # 方法1: 从页面文本中查找状态关键词
        # "本场已结束" / "竞价成功" / "已成交" 等
        for kw in SUCCESS_KEYWORDS:
            if kw in page_text:
                status_text = kw
                break

        if not status_text:
            for kw in FAIL_KEYWORDS:
                if kw in page_text:
                    status_text = kw
                    break

        # 方法2: 从特定元素提取
        if not status_text:
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
                        break
                except NoSuchElementException:
                    continue

        # 判断是否成功
        is_success = False
        if status_text:
            is_success = any(kw in status_text for kw in SUCCESS_KEYWORDS)
            # 如果有"已结束"但没有失败关键词，默认认为成功
            if "已结束" in status_text and not any(kw in status_text for kw in FAIL_KEYWORDS):
                is_success = True

        return {"text": status_text, "success": is_success}

    def _extract_intro_table(self):
        """
        提取"标的物介绍"部分的整个表格。
        返回 {"table": {key: value}, "raw": str}
        """
        table_data = {}
        raw_text = ""

        # 策略1: 查找包含"标的物介绍"的标题，然后提取其后的表格
        try:
            # 找到"标的物介绍"文本所在的元素
            body = self.driver.find_element(By.TAG_NAME, "body")
            page_text = body.text

            # 方法A: 直接从页面文本用正则提取 key-value 对
            # 常见格式："建筑总面积\n123.45㎡" 或 "建筑总面积：123.45㎡"
            intro_patterns = [
                # 表格行格式：key 在左列，value 在右列
                r'([^\n：:]+)[：:]\s*([^\n]+)',
            ]

            # 找到"标的物介绍"区域
            intro_start = -1
            for marker in ["标的物介绍", "拍卖标的基本情况", "标的物详情"]:
                idx = page_text.find(marker)
                if idx != -1:
                    intro_start = idx + len(marker)
                    break

            if intro_start != -1:
                # 从"标的物介绍"到下一个大标题之间的文本
                intro_text = page_text[intro_start:]
                # 截断到下一个大标题
                for end_marker in ["竞买公告", "竞买须知", "拍卖规则", "附件", "相关附件",
                                   "重要提示", "竞买人条件", "咨询时间", "拍卖余款"]:
                    end_idx = intro_text.find(end_marker)
                    if end_idx != -1:
                        intro_text = intro_text[:end_idx]
                        break

                raw_text = intro_text.strip()

                # 解析 key-value
                lines = [l.strip() for l in intro_text.split('\n') if l.strip()]
                i = 0
                while i < len(lines):
                    line = lines[i]
                    # 跳过纯标题行
                    if len(line) < 2:
                        i += 1
                        continue

                    # 检查 "key：value" 或 "key:value" 格式
                    m = re.match(r'^([^：:]+)[：:]\s*(.+)$', line)
                    if m:
                        key = m.group(1).strip()
                        val = m.group(2).strip()
                        if key and val:
                            table_data[key] = val
                    elif i + 1 < len(lines):
                        # 可能是 key 和 value 分两行
                        # 判断：当前行像 key（较短、无数字开头），下一行像 value
                        next_line = lines[i + 1]
                        if (len(line) < 20 and not re.match(r'^[\d.]+', line)
                            and len(next_line) > 0 and line != next_line):
                            table_data[line] = next_line
                            i += 1  # 跳过下一行
                    i += 1

        except Exception as e:
            logger.warning(f"  提取标的物介绍异常: {e}")

        # 策略2: 从 HTML table 标签提取
        if not table_data:
            try:
                tables = self.driver.find_elements(By.CSS_SELECTOR, "table")
                for table in tables:
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 2:
                            key = cells[0].text.strip()
                            val = cells[1].text.strip()
                            if key and val and len(key) < 30:
                                table_data[key] = val
                    # 如果找到了数据，只取第一个有意义的表格
                    if len(table_data) > 3:
                        break
            except Exception as e:
                logger.warning(f"  HTML表格提取异常: {e}")

        # 策略3: 用 CSS 选择器找特定结构
        if not table_data:
            try:
                # 有些页面用 div 布局模拟表格
                rows = self.driver.find_elements(By.CSS_SELECTOR,
                    "[class*='info-row'], [class*='detail-row'], [class*='attr-item'], "
                    "[class*='prop-item'], [class*='spec-item']"
                )
                for row in rows:
                    try:
                        label = row.find_element(By.CSS_SELECTOR,
                            "[class*='label'], [class*='key'], [class*='name'], dt"
                        ).text.strip()
                        value = row.find_element(By.CSS_SELECTOR,
                            "[class*='value'], [class*='val'], [class*='content'], dd"
                        ).text.strip()
                        if label and value:
                            table_data[label] = value
                    except NoSuchElementException:
                        continue
            except Exception:
                pass

        return {"table": table_data, "raw": raw_text}

    def _extract_attachments(self):
        """提取附件下载链接"""
        attachments = []
        try:
            # 查找所有可能是附件的链接
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

            # 从页面文本找附件名
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            att_m = re.findall(r'相关附件[：:]\s*(\S+\.(?:pdf|docx?|xlsx?|zip|rar))', body_text, re.I)
            for name in att_m:
                # 检查是否已经在列表中
                if not any(name in a["name"] for a in attachments):
                    attachments.append({"name": name, "url": ""})

        except Exception as e:
            logger.warning(f"  提取附件异常: {e}")

        return attachments


# ============================================================
# 附件下载器
# ============================================================

def download_attachment(url, save_path, driver=None):
    """下载附件到指定路径"""
    if not url:
        logger.warning(f"  附件 URL 为空，跳过: {save_path}")
        return False

    try:
        # 方法1: 用 requests（如果有直接链接）
        headers = {"User-Agent": random.choice(USER_AGENTS)}
        resp = requests.get(url, headers=headers, timeout=60, stream=True)
        if resp.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(8192):
                    f.write(chunk)
            logger.info(f"  ✓ 附件已下载: {os.path.basename(save_path)}")
            return True
    except Exception:
        pass

    # 方法2: 用 Selenium 触发下载
    if driver:
        try:
            driver.get(url)
            time.sleep(5)
            # 浏览器会自动下载到默认目录
            logger.info(f"  ✓ 附件已触发下载: {os.path.basename(save_path)}")
            return True
        except Exception as e:
            logger.warning(f"  附件下载失败: {e}")

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
        # 输出根目录
        self.output_root = os.path.join(get_script_dir(), "output")

    def start(self):
        """启动浏览器"""
        self.stats["start_time"] = time.time()

        if self.browser_type == "edge":
            opts = EdgeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            opts.add_argument("--window-size=1920,1080")
            # 不完全禁用图片，有些页面需要看验证码
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

        # 隐藏 webdriver 特征
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })

        # 注入 Cookie
        self._inject_cookie(self.cookie)

        self.list_scraper = ListPageScraper(self.driver)
        self.detail_scraper = DetailPageScraper(self.driver)
        logger.info(f"浏览器已启动 ({self.browser_type.upper()})")

    def stop(self):
        if self.driver:
            try: self.driver.quit()
            except: pass

    def _inject_cookie(self, cookie_str):
        """注入 Cookie 到浏览器"""
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
        """
        处理反爬拦截/Cookie失效。
        呼出系统浏览器让用户手动验证，然后重新输入 Cookie。
        返回 True 表示恢复成功，False 表示用户放弃。
        """
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

        # 更新 Cookie
        self.cookie = new_cookie
        try:
            self._inject_cookie(new_cookie)
            logger.info("✓ Cookie 已更新，继续爬取")
            return True
        except Exception as e:
            logger.error(f"Cookie 注入失败: {e}")
            return False

    def _save_item(self, detail, date_range_str):
        """
        保存单条拍品数据到文件系统。
        目录结构：output/日期范围/拍卖物名称/
        """
        title = detail.get("title", "未知")
        safe_title = sanitize_filename(title)
        if not safe_title:
            safe_title = f"item_{int(time.time())}"

        # 创建目录
        item_dir = os.path.join(self.output_root, date_range_str, safe_title)
        safe_makedirs(item_dir)

        # 1. 保存 URL
        with open(os.path.join(item_dir, "url.txt"), "w", encoding="utf-8") as f:
            f.write(detail.get("url", ""))

        # 2. 保存拍卖信息
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

        # 3. 保存标的物介绍表格
        intro_table = detail.get("intro_table", {})
        if intro_table:
            with open(os.path.join(item_dir, "标的物介绍.csv"), "w",
                      newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["项目", "内容"])
                for k, v in intro_table.items():
                    writer.writerow([k, v])

        # 4. 保存标的物介绍原始文本
        intro_raw = detail.get("intro_raw", "")
        if intro_raw:
            with open(os.path.join(item_dir, "标的物介绍.txt"), "w", encoding="utf-8") as f:
                f.write(intro_raw)

        # 5. 下载附件
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

        logger.info(f"  ✓ 已保存: {safe_title[:50]}")
        return item_dir

    def run(self, date_start="2024-01-01", date_end="2026-03-31", mode="full"):
        """
        主流程。
        mode: "full" = 列表页+详情页, "detail_only" = 只处理已有URL列表
        """
        date_range_str = f"{date_start}_to_{date_end}"

        if mode == "full":
            self._run_full(date_start, date_end, date_range_str)
        elif mode == "detail_only":
            self._run_detail_only(date_range_str)

        self._print_stats()

    def _run_full(self, date_start, date_end, date_range_str):
        """完整流程：列表页 → 详情页"""
        # 检查列表页断点
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

        all_urls = list(cp_urls)  # 已收集的 URL

        # 阶段1: 从列表页收集所有 URL
        logger.info("=" * 60)
        logger.info(f"阶段1: 收集拍品 URL")
        logger.info(f"日期: {date_start} ~ {date_end}")
        logger.info("=" * 60)

        for page in range(start_page, MAX_PAGES + 1):
            status, items = self.list_scraper.fetch_page(page, date_start, date_end)

            if status == "blocked":
                if self._handle_blocked("blocked"):
                    # 重试当前页
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

        # 阶段2: 逐个处理详情页
        self._process_details(all_urls, date_range_str)

    def _run_detail_only(self, date_range_str):
        """只处理已有 URL 列表（从文件读取）"""
        url_file = os.path.join(get_script_dir(), "urls.txt")
        if not os.path.isfile(url_file):
            logger.error(f"未找到 URL 列表文件: {url_file}")
            logger.info("请创建 urls.txt，每行一个详情页 URL")
            return

        with open(url_file, "r", encoding="utf-8") as f:
            all_urls = [l.strip() for l in f if l.strip() and l.strip().startswith("http")]

        logger.info(f"从 urls.txt 读取到 {len(all_urls)} 个 URL")
        self._process_details(all_urls, date_range_str)

    def _process_details(self, all_urls, date_range_str):
        """处理所有详情页"""
        # 检查详情页断点
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

        for idx, url in enumerate(all_urls):
            if url in cp_visited:
                continue

            logger.info(f"\n[{idx+1}/{len(all_urls)}] {url[:80]}...")

            # 检查是否是详情页 URL
            if "sf_item" not in url and "auction" not in url and "susong" not in url:
                logger.info(f"  跳过（非详情页URL）")
                cp_visited.add(url)
                continue

            random_delay()

            # 获取详情页
            detail_result = self.detail_scraper.fetch_detail(url)

            if isinstance(detail_result, tuple):
                status, detail = detail_result
            else:
                # 兼容旧返回格式
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
                    # 重试
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

            # 筛选：只保存成功拍卖的
            if not detail.get("is_success"):
                logger.info(f"  ⏭ 跳过（未成交）: {detail.get('status', '未知状态')}")
                self.stats["total_skipped"] += 1
                self.checkpoint_detail.save(cp_visited, 0, self.stats["total_saved"])
                continue

            # 保存数据
            self._save_item(detail, date_range_str)
            self.stats["total_saved"] += 1
            self.checkpoint_detail.save(cp_visited, 0, self.stats["total_saved"])

        # 完成后清除断点
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
        """保存汇总 CSV"""
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

            # 读取拍卖信息
            info_file = os.path.join(item_dir, "拍卖信息.txt")
            if os.path.isfile(info_file):
                with open(info_file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if ":" in line:
                            k, v = line.split(":", 1)
                            row[k.strip()] = v.strip()

            # 读取标的物介绍
            intro_file = os.path.join(item_dir, "标的物介绍.csv")
            if os.path.isfile(intro_file):
                with open(intro_file, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    next(reader, None)  # 跳过表头
                    for csv_row in reader:
                        if len(csv_row) >= 2:
                            row[f"介绍_{csv_row[0]}"] = csv_row[1]

            # 附件数量
            att_dir = os.path.join(item_dir, "附件")
            if os.path.isdir(att_dir):
                att_count = len(os.listdir(att_dir))
                row["附件数量"] = str(att_count)

            rows.append(row)

        if rows:
            # 收集所有字段名
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
║       阿里司法拍卖 - 法拍房数据爬取工具 (第六版)             ║
║                                                              ║
║  ✨ 全新策略：逐页提取 | 标的物介绍整表 | 附件下载            ║
║  ✨ 自动处理 Cookie 过期 | 断点续爬                           ║
╚══════════════════════════════════════════════════════════════╝

📌 准备工作：
   1. 用浏览器打开 https://sf.taobao.com/ 并登录
   2. F12 → Network → 刷新 → 复制 Cookie
""")

    if not HAS_SELENIUM:
        print("❌ 未安装 selenium！请运行: pip install selenium")
        sys.exit(1)

    # 检测浏览器
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

    # 下载 driver
    driver_path = local_driver
    if not driver_path:
        if browser_type == "edge":
            driver_path = download_edgedriver()
            if not driver_path:
                print("❌ msedgedriver 下载失败"); sys.exit(1)
        else:
            dp = input("  chromedriver 路径: ").strip()
            if dp: driver_path = dp

    # 启动
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
