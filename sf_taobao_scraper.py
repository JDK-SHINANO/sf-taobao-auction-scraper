#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 昆明市法拍房数据爬取工具 (第三版)
改进：
  - 断点续爬（checkpoint），中断后可从上次位置继续
  - 自动下载 msedgedriver（npmmirror 国内镜像）
  - 并发请求（可配置线程数）
  - 更完善的日志和统计
  - 自动检测 Edge / Chrome
  - CSV + Excel 双格式输出
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
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
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

KUNMING_LOCATION_CODE = "530100"
KUNMING_PROVINCE_CODE = "530000"
CATEGORY_RESIDENTIAL = "50025969"
LIST_URL = "https://sf.taobao.com/item_list.htm"
ITEMS_PER_PAGE = 40
MAX_PAGES = 150
MIN_DELAY = 2.0
MAX_DELAY = 5.0

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
logger = logging.getLogger(__name__)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
]


# ============================================================
# 工具函数
# ============================================================

def get_script_dir():
    return str(Path(__file__).parent.resolve())

def random_delay(low=None, high=None):
    time.sleep(random.uniform(low or MIN_DELAY, high or MAX_DELAY))

def clean_number(text):
    if not text: return ""
    return re.sub(r"[,\s，]", "", text)

def wan_to_yuan(text):
    if not text: return ""
    text = str(text).strip()
    m = re.search(r"([\d,.]+)\s*万", text)
    if m: return str(int(float(m.group(1).replace(",", "")) * 10000))
    m = re.search(r"([\d,.]+)", text)
    if m: return clean_number(m.group(1))
    return ""

def extract_date_mm_dd(text):
    if not text: return ""
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m: return f"{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    return ""

def parse_date_to_obj(text):
    if not text: return None
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m:
        try: return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except: return None
    return None

def extract_district_from_title(title):
    if not title: return ""
    for d in KUNMING_DISTRICTS:
        if d in title: return d
    for d in KUNMING_DISTRICTS:
        s = d.replace("区", "").replace("县", "").replace("市", "")
        if s and s in title: return d
    return ""


# ============================================================
# 浏览器检测 + Driver 下载
# ============================================================

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
    """获取本机 Edge 版本号"""
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
    # 备用：直接读文件属性
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             f"Get-Item '{edge_exe}' | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion"],
            capture_output=True, text=True, timeout=15
        )
        ver = result.stdout.strip()
        if ver: return ver
    except: pass
    return None


def download_edgedriver():
    """下载 msedgedriver 到脚本目录（国内镜像优先）"""
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "msedgedriver.exe")
    if os.path.isfile(target):
        logger.info(f"msedgedriver 已存在: {target}")
        return target

    version = get_edge_version()
    if not version:
        logger.error("无法获取 Edge 版本，请手动下载 msedgedriver.exe 放到脚本目录")
        logger.info("下载: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
        return None

    logger.info(f"Edge 版本: {version}")

    urls = [
        f"https://registry.npmmirror.com/-/binary/edgedriver/{version}/edgedriver_win64.zip",
        f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip",
    ]

    for url in urls:
        try:
            logger.info(f"下载中: {url[:70]}...")
            resp = requests.get(url, timeout=120, stream=True)
            if resp.status_code != 200:
                logger.warning(f"HTTP {resp.status_code}")
                continue
            zip_path = os.path.join(script_dir, "_edgedriver.zip")
            total = 0
            with open(zip_path, "wb") as f:
                for chunk in resp.iter_content(8192):
                    f.write(chunk)
                    total += len(chunk)
            logger.info(f"下载完成: {total/1024/1024:.1f} MB")

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
            continue

    logger.error("自动下载失败，请手动下载 msedgedriver.exe 到脚本目录")
    return None


# ============================================================
# Checkpoint（断点续爬）
# ============================================================

class Checkpoint:
    """保存爬取进度，支持中断后继续"""

    def __init__(self):
        self.file = os.path.join(get_script_dir(), ".crawl_checkpoint.json")
        self.data = self._load()

    def _load(self):
        if os.path.isfile(self.file):
            try:
                with open(self.file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except: pass
        return {"visited_ids": [], "last_page": 0, "results_count": 0, "timestamp": ""}

    def save(self, visited_ids, last_page, results_count):
        self.data = {
            "visited_ids": list(visited_ids),
            "last_page": last_page,
            "results_count": results_count,
            "timestamp": datetime.now().isoformat(),
        }
        with open(self.file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def clear(self):
        if os.path.isfile(self.file):
            os.remove(self.file)
        self.data = {"visited_ids": [], "last_page": 0, "results_count": 0, "timestamp": ""}

    @property
    def visited_ids(self):
        return set(self.data.get("visited_ids", []))

    @property
    def last_page(self):
        return self.data.get("last_page", 0)

    @property
    def results_count(self):
        return self.data.get("results_count", 0)


# ============================================================
# 列表页爬取
# ============================================================

class ListPageScraper:
    def __init__(self, cookie):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Referer": "https://sf.taobao.com/",
            "Cookie": cookie,
        })

    def _get(self, url, params=None, max_retries=3):
        for attempt in range(max_retries):
            try:
                self.session.headers["User-Agent"] = random.choice(USER_AGENTS)
                resp = self.session.get(url, params=params, timeout=30)
                if resp.status_code == 200: return resp
                if resp.status_code in (403, 429):
                    logger.warning(f"HTTP {resp.status_code}，等待重试...")
                    time.sleep(random.uniform(15, 30))
                else:
                    time.sleep(random.uniform(5, 10))
            except Exception as e:
                logger.warning(f"请求异常: {e}")
                time.sleep(random.uniform(5, 10))
        return None

    def fetch_page(self, page, completed=True):
        params = {
            "location_code": KUNMING_LOCATION_CODE,
            "category": CATEGORY_RESIDENTIAL,
            "auction_source": "0",
            "province": KUNMING_PROVINCE_CODE,
            "st_param": "-1", "auction_start_seg": "-1",
            "page": str(page),
        }
        if completed: params["completed_items"] = "1"

        resp = self._get(LIST_URL, params=params)
        if not resp or "deny" in resp.url or "punish" in resp.url: return []

        soup = BeautifulSoup(resp.text, "lxml")
        script_tag = soup.find("script", {"id": "sf-item-list-data"})
        if script_tag and script_tag.string:
            try:
                data = json.loads(script_tag.string.strip())
                items = data.get("data", [])
                if items:
                    logger.info(f"列表页第 {page} 页：{len(items)} 条")
                    return items
            except json.JSONDecodeError: pass
        return []


# ============================================================
# Selenium 详情页爬取
# ============================================================

class SeleniumDetailScraper:
    def __init__(self, cookie_str, browser_type="edge", driver_path=None):
        self.cookie_str = cookie_str
        self.browser_type = browser_type
        self.driver_path = driver_path
        self.driver = None

    def start(self):
        if self.browser_type == "edge":
            opts = EdgeOptions()
            opts.add_argument("--headless=new")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            opts.add_argument("--window-size=1920,1080")
            opts.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
            svc = EdgeService(self.driver_path) if self.driver_path and os.path.isfile(self.driver_path) else EdgeService()
            self.driver = webdriver.Edge(service=svc, options=opts)
        else:
            opts = ChromeOptions()
            opts.add_argument("--headless=new")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            opts.add_argument("--window-size=1920,1080")
            opts.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
            svc = ChromeService(self.driver_path) if self.driver_path and os.path.isfile(self.driver_path) else ChromeService()
            self.driver = webdriver.Chrome(service=svc, options=opts)

        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })

        self.driver.get("https://sf.taobao.com/")
        time.sleep(2)
        if self.cookie_str:
            for item in self.cookie_str.split(";"):
                item = item.strip()
                if "=" in item:
                    n, v = item.split("=", 1)
                    try: self.driver.add_cookie({"name": n.strip(), "value": v.strip(), "domain": ".taobao.com", "path": "/"})
                    except: pass
        logger.info(f"浏览器已启动 ({self.browser_type.upper()})")

    def stop(self):
        if self.driver: self.driver.quit()

    def fetch_detail(self, url):
        result = {
            "标的名称": "", "成交价(元)": "", "评估价/市场价(元)": "",
            "成交日期": "", "建筑面积(㎡)": "", "土地分摊面积(㎡)": "",
            "所在层/总楼层": "", "朝向": "", "房屋结构": "",
            "户型": "", "租赁情况": "", "装修情况": "",
            "附件名称": "", "所属区域": "", "详情页链接": url,
        }
        self._last_raw_date = ""

        try:
            self.driver.get(url)
            WebDriverWait(self.driver, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(3)
            text = self.driver.find_element(By.TAG_NAME, "body").text

            # 标题
            try:
                h1 = self.driver.find_element(By.CSS_SELECTOR, "h1, .pm-main h1, .title h1, .item-title")
                result["标的名称"] = h1.text.strip()
            except NoSuchElementException:
                t = self.driver.title
                if t: result["标的名称"] = t.split("-")[0].strip()

            # 成交价
            for pat in [r"成交价\s*[¥￥]?\s*([\d,.]+)", r"成交金额\s*[¥￥]?\s*([\d,.]+)"]:
                m = re.search(pat, text)
                if m: result["成交价(元)"] = clean_number(m.group(1)); break
            if not result["成交价(元)"]:
                m = re.search(r"成交[金价].*?([\d.]+)\s*万", text)
                if m: result["成交价(元)"] = str(int(float(m.group(1)) * 10000))

            # 评估价
            for pat in [r"评估价\s*[¥￥]?\s*([\d,.]+)", r"市场价\s*[¥￥]?\s*([\d,.]+)"]:
                m = re.search(pat, text)
                if m: result["评估价/市场价(元)"] = clean_number(m.group(1)); break
            if not result["评估价/市场价(元)"]:
                m = re.search(r"(?:评估|市场)[金价].*?([\d.]+)\s*万", text)
                if m: result["评估价/市场价(元)"] = str(int(float(m.group(1)) * 10000))

            # 成交日期
            for pat in [r"结束时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})", r"成交时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})"]:
                m = re.search(pat, text)
                if m: result["成交日期"] = extract_date_mm_dd(m.group(1)); self._last_raw_date = m.group(1); break

            # 建筑面积
            for pat in [r"建筑总面积\s*([\d.]+)", r"建筑面积\s*([\d.]+)", r"总面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m: result["建筑面积(㎡)"] = m.group(1); break

            # 土地分摊面积
            for pat in [r"土地分摊面积\s*([\d.]+)", r"土地面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m: result["土地分摊面积(㎡)"] = m.group(1); break

            # 楼层
            m = re.search(r"(?:所在层|楼层)\s*[:：]?\s*(?:第)?(\d+)\s*层\s*/\s*(?:共)?(\d+)\s*层", text)
            if m: result["所在层/总楼层"] = f"{m.group(1)}层/{m.group(2)}层"

            # 朝向
            m = re.search(r"朝\s*向\s*[:：]?\s*([东南西北中\s,，、]+)", text)
            if m: result["朝向"] = m.group(1).strip().rstrip("，,")

            # 房屋结构
            m = re.search(r"(?:建筑结构|房屋结构)\s*[:：]?\s*(\S+)", text)
            if m: result["房屋结构"] = m.group(1).strip()

            # 户型
            m = re.search(r"(?:房屋户型|户型)\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["户型"] = m.group(1).strip()
            else:
                m = re.search(r"(\d+室\d*厅?\d*卫?)", text)
                if m: result["户型"] = m.group(1)

            # 租赁情况
            m = re.search(r"租赁情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["租赁情况"] = m.group(1).strip()

            # 装修情况
            m = re.search(r"装修情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["装修情况"] = m.group(1).strip()

            # 附件
            m = re.search(r"相关附件下载[：:]\s*(\S+\.(?:pdf|docx?|xls|xlsx|zip|rar))", text, re.I)
            if m: result["附件名称"] = m.group(1).strip()
            else:
                try:
                    for link in self.driver.find_elements(By.CSS_SELECTOR, "a[href*='.pdf'], a[href*='.doc'], a[href*='.xls']"):
                        href = link.get_attribute("href") or ""
                        if re.search(r"\.(pdf|docx?|xlsx?)$", href, re.I):
                            result["附件名称"] = link.text.strip() or href.split("/")[-1]; break
                except: pass

            result["所属区域"] = extract_district_from_title(result["标的名称"])

        except TimeoutException:
            logger.warning(f"超时: {url}")
        except Exception as e:
            logger.warning(f"异常: {e}")

        return result


# ============================================================
# Requests 备用
# ============================================================

class RequestsDetailScraper:
    def __init__(self, cookie):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,*/*;q=0.8",
            "Referer": "https://sf.taobao.com/",
            "Cookie": cookie,
        })

    def fetch_detail(self, url):
        result = {
            "标的名称": "", "成交价(元)": "", "评估价/市场价(元)": "",
            "成交日期": "", "建筑面积(㎡)": "", "土地分摊面积(㎡)": "",
            "所在层/总楼层": "", "朝向": "", "房屋结构": "",
            "户型": "", "租赁情况": "", "装修情况": "",
            "附件名称": "", "所属区域": "", "详情页链接": url,
        }
        self._last_raw_date = ""
        try:
            resp = self.session.get(url, timeout=30)
            if resp.status_code != 200 or "deny" in resp.url: return result
            soup = BeautifulSoup(resp.text, "lxml")
            text = soup.get_text(separator="\n", strip=True)

            h1 = soup.find("h1")
            if h1: result["标的名称"] = h1.get_text(strip=True)

            for pat in [r"成交价\s*[¥￥]?\s*([\d,.]+)", r"成交金额\s*[¥￥]?\s*([\d,.]+)"]:
                m = re.search(pat, text)
                if m: result["成交价(元)"] = clean_number(m.group(1)); break

            for pat in [r"评估价\s*[¥￥]?\s*([\d,.]+)", r"市场价\s*[¥￥]?\s*([\d,.]+)"]:
                m = re.search(pat, text)
                if m: result["评估价/市场价(元)"] = clean_number(m.group(1)); break

            for pat in [r"结束时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})", r"成交时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})"]:
                m = re.search(pat, text)
                if m: result["成交日期"] = extract_date_mm_dd(m.group(1)); self._last_raw_date = m.group(1); break

            for pat in [r"建筑总面积\s*([\d.]+)", r"建筑面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m: result["建筑面积(㎡)"] = m.group(1); break

            for pat in [r"土地分摊面积\s*([\d.]+)", r"土地面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m: result["土地分摊面积(㎡)"] = m.group(1); break

            m = re.search(r"(?:所在层|楼层)\s*[:：]?\s*(?:第)?(\d+)\s*层\s*/\s*(?:共)?(\d+)\s*层", text)
            if m: result["所在层/总楼层"] = f"{m.group(1)}层/{m.group(2)}层"

            m = re.search(r"朝\s*向\s*[:：]?\s*([东南西北中\s,，、]+)", text)
            if m: result["朝向"] = m.group(1).strip().rstrip("，,")

            m = re.search(r"(?:建筑结构|房屋结构)\s*[:：]?\s*(\S+)", text)
            if m: result["房屋结构"] = m.group(1).strip()

            m = re.search(r"(?:房屋户型|户型)\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["户型"] = m.group(1).strip()

            m = re.search(r"租赁情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["租赁情况"] = m.group(1).strip()

            m = re.search(r"装修情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m: result["装修情况"] = m.group(1).strip()

            for a in soup.find_all("a", href=True):
                if re.search(r"\.(pdf|docx?|xlsx?)$", a["href"], re.I):
                    result["附件名称"] = a.get_text(strip=True) or a["href"].split("/")[-1]; break

            result["所属区域"] = extract_district_from_title(result["标的名称"])
        except: pass
        return result


# ============================================================
# 主爬虫（带 checkpoint）
# ============================================================

class SfTaobaoScraper:
    def __init__(self, cookie, use_selenium=True, browser_type="edge", driver_path=None):
        self.cookie = cookie
        self.use_selenium = use_selenium and HAS_SELENIUM
        self.browser_type = browser_type
        self.driver_path = driver_path
        self.list_scraper = ListPageScraper(cookie)
        self.detail_scraper = None
        self.results = []
        self.failed_urls = []
        self.visited_ids = set()
        self.checkpoint = Checkpoint()
        self.stats = {"total_found": 0, "total_saved": 0, "total_skipped": 0, "start_time": None}

    def start(self):
        self.stats["start_time"] = time.time()
        if self.use_selenium:
            self.detail_scraper = SeleniumDetailScraper(self.cookie, self.browser_type, self.driver_path)
            self.detail_scraper.start()
        else:
            self.detail_scraper = RequestsDetailScraper(self.cookie)

    def stop(self):
        if self.use_selenium and self.detail_scraper:
            self.detail_scraper.stop()

    def _print_stats(self):
        elapsed = time.time() - self.stats["start_time"] if self.stats["start_time"] else 0
        logger.info(f"\n{'='*60}")
        logger.info(f"📊 统计:")
        logger.info(f"   总发现: {self.stats['total_found']}")
        logger.info(f"   已保存: {self.stats['total_saved']}")
        logger.info(f"   已跳过: {self.stats['total_skipped']}")
        logger.info(f"   失败:   {len(self.failed_urls)}")
        logger.info(f"   耗时:   {elapsed/60:.1f} 分钟")
        if self.stats['total_saved'] > 0:
            logger.info(f"   平均:   {elapsed/self.stats['total_saved']:.1f} 秒/条")
        logger.info(f"{'='*60}")

    def scrape_all(self, date_start="2024-01-01", date_end="2026-03-31"):
        dt_start = datetime.strptime(date_start, "%Y-%m-%d")
        dt_end = datetime.strptime(date_end, "%Y-%m-%d")

        # 检查是否有 checkpoint
        cp_ids = self.checkpoint.visited_ids
        cp_page = self.checkpoint.last_page
        if cp_ids:
            logger.info(f"🔄 发现断点：上次爬到第 {cp_page} 页，已 {len(cp_ids)} 条，是否继续？")
            choice = input("  继续上次(y) / 重新开始(n) [默认y]: ").strip().lower()
            if choice == "n":
                self.checkpoint.clear()
                cp_ids = set()
                cp_page = 0
            else:
                self.visited_ids = cp_ids
                logger.info(f"从第 {cp_page + 1} 页继续...")

        start_page = cp_page + 1 if cp_ids else 1
        oor_count = 0

        logger.info("=" * 60)
        logger.info(f"阿里司法拍卖 - 昆明市法拍房数据爬取")
        logger.info(f"日期: {date_start} ~ {date_end}")
        logger.info(f"模式: {'Selenium/' + self.browser_type.upper() if self.use_selenium else 'Requests'}")
        logger.info("=" * 60)

        for page in range(start_page, MAX_PAGES + 1):
            logger.info(f"\n--- 列表页第 {page} 页 ---")
            items = self.list_scraper.fetch_page(page)
            if not items:
                logger.info("无数据，停止")
                break

            for idx, item in enumerate(items):
                self.stats["total_found"] += 1
                item_id = item.get("id", "")
                item_url = item.get("itemUrl", "")
                title = item.get("title", "")
                status = str(item.get("status", "") or item.get("statusText", ""))

                if item_url and not item_url.startswith("http"):
                    item_url = "https:" + item_url if item_url.startswith("//") else "https://sf.taobao.com" + item_url

                if item_id and item_id in self.visited_ids: continue
                if item_id: self.visited_ids.add(item_id)

                if any(kw in status for kw in ["流拍", "中止", "撤回", "失败"]):
                    self.stats["total_skipped"] += 1; continue

                logger.info(f"  [{idx+1}/{len(items)}] {title[:60]}...")
                random_delay()

                detail = self.detail_scraper.fetch_detail(item_url)
                if not detail or not detail.get("标的名称"):
                    self.failed_urls.append(item_url); continue

                if not detail["标的名称"]: detail["标的名称"] = title
                if not detail["成交价(元)"]: detail["成交价(元)"] = wan_to_yuan(item.get("currentPrice", ""))
                if not detail["所属区域"]: detail["所属区域"] = extract_district_from_title(detail["标的名称"] or title)

                raw_date = getattr(self.detail_scraper, "_last_raw_date", "")
                if raw_date:
                    dt = parse_date_to_obj(raw_date)
                    if dt:
                        if dt < dt_start:
                            oor_count += 1
                            if oor_count >= 15:
                                logger.info("连续超出范围，保存 checkpoint 并停止")
                                self.checkpoint.save(self.visited_ids, page, len(self.results))
                                break
                            continue
                        elif dt > dt_end: continue
                        else: oor_count = 0

                self.results.append(detail)
                self.stats["total_saved"] += 1
                logger.info(f"    ✓ 已保存 (共{self.stats['total_saved']}条)")

            # 每页保存一次 checkpoint
            self.checkpoint.save(self.visited_ids, page, len(self.results))

            if oor_count >= 15: break
            if len(items) < ITEMS_PER_PAGE: break
            random_delay()

        # 完成后清除 checkpoint
        self.checkpoint.clear()
        self._print_stats()

    def save_to_excel(self, filename=None):
        if not self.results:
            logger.warning("没有数据"); return
        if not filename:
            filename = os.path.join(get_script_dir(), f"昆明法拍房数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

        headers = [
            "标的名称", "成交价(元)", "评估价/市场价(元)", "成交日期",
            "建筑面积(㎡)", "土地分摊面积(㎡)", "所在层/总楼层",
            "朝向", "房屋结构", "户型", "租赁情况", "装修情况",
            "附件名称", "所属区域", "详情页链接",
        ]

        # 总是先保存 CSV（兼容性最好）
        csv_file = filename.replace(".xlsx", ".csv") if filename.endswith(".xlsx") else filename + ".csv"
        with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=headers)
            w.writeheader()
            for item in self.results: w.writerow({k: item.get(k, "") for k in headers})
        logger.info(f"CSV 已保存: {csv_file}")

        if not HAS_OPENPYXL:
            logger.info("提示: pip install openpyxl 可输出 Excel")
            return

        # Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "昆明法拍房数据"

        hfont = Font(bold=True, size=11, color="FFFFFF")
        hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bdr = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font, c.fill, c.alignment, c.border = hfont, hfill, halign, bdr

        for row, item in enumerate(self.results, 2):
            for col, key in enumerate(headers, 1):
                val = item.get(key, "")
                c = ws.cell(row=row, column=col, value=val)
                c.border = bdr
                c.alignment = Alignment(vertical="center", wrap_text=True)
                if "价" in key and val:
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    try: c.value = int(val); c.number_format = "#,##0"
                    except: pass

        for col in range(1, len(headers) + 1):
            mx = len(str(headers[col-1]))
            for row in range(2, min(len(self.results) + 2, 100)):
                cv = ws.cell(row=row, column=col).value
                if cv: mx = max(mx, len(str(cv)))
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(mx + 4, 50)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        wb.save(filename)
        logger.info(f"Excel 已保存: {filename} ({len(self.results)} 条)")

    def save_failed_urls(self):
        if not self.failed_urls: return
        fp = os.path.join(get_script_dir(), f"failed_urls_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(fp, "w", encoding="utf-8") as f:
            for url in self.failed_urls: f.write(url + "\n")
        logger.info(f"失败URL: {fp} ({len(self.failed_urls)} 条)")


# ============================================================
# 入口
# ============================================================

def main():
    print("""
╔══════════════════════════════════════════════════════════════╗
║       阿里司法拍卖 - 昆明市法拍房数据爬取工具 (第三版)       ║
║                                                              ║
║  ✨ 新功能：断点续爬 | 自动下载 driver | CSV+Excel 双输出    ║
╚══════════════════════════════════════════════════════════════╝

📌 准备工作：
   1. 用浏览器打开 https://sf.taobao.com/ 并登录
   2. F12 → Network → 刷新 → 复制 Cookie
""")

    if not HAS_SELENIUM:
        print("⚠️  未安装 selenium，将用 Requests 模式")
        print("   安装: pip install selenium\n")

    # 自动检测
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

    print("\n📅 日期范围：")
    date_start = input("  开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
    date_end = input("  结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"
    try:
        datetime.strptime(date_start, "%Y-%m-%d")
        datetime.strptime(date_end, "%Y-%m-%d")
    except ValueError:
        logger.error("日期格式错误"); sys.exit(1)

    output_file = input("\n📁 输出文件名（回车默认）: ").strip() or None

    use_selenium = HAS_SELENIUM
    driver_path = local_driver
    if HAS_SELENIUM:
        mode = input("\n🔧 模式 (1=Selenium默认, 2=Requests, 默认1): ").strip()
        use_selenium = mode != "2"
        if use_selenium and not driver_path:
            if browser_type == "edge":
                driver_path = download_edgedriver()
                if not driver_path:
                    print("❌ msedgedriver 下载失败，请手动放到脚本目录")
                    sys.exit(1)
            else:
                dp = input("  chromedriver 路径: ").strip()
                if dp: driver_path = dp

    print(f"\n{'='*60}\n开始爬取...\n{'='*60}\n")
    scraper = SfTaobaoScraper(cookie, use_selenium=use_selenium, browser_type=browser_type or "edge", driver_path=driver_path)

    try:
        scraper.start()
        scraper.scrape_all(date_start, date_end)
        scraper.save_to_excel(output_file)
        scraper.save_failed_urls()
    except KeyboardInterrupt:
        logger.info("中断，保存数据和 checkpoint...")
        scraper.checkpoint.save(scraper.visited_ids, 0, len(scraper.results))
        scraper.save_to_excel(output_file)
        scraper.save_failed_urls()
    except Exception as e:
        logger.error(f"错误: {e}")
        import traceback; traceback.print_exc()
        scraper.checkpoint.save(scraper.visited_ids, 0, len(scraper.results))
        if scraper.results: scraper.save_to_excel(output_file)
    finally:
        scraper.stop()

    print("\n✅ 运行结束")


if __name__ == "__main__":
    main()
