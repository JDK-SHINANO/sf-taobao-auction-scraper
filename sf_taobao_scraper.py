#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 法拍房数据爬取工具 v0.18.1
============================================
- 修复菜单选项编号，统一按键逻辑
- 支持多城市/房屋类型/区县/资产来源筛选
- 交互式菜单，单键选择（Windows 支持 msvcrt）
- 动态 URL 生成
- 保留原有断点续传、反爬处理、表A/B提取等功能
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

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    pass

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
# 城市/房屋类型/资产来源配置
# ============================================================

CITY_CONFIG = {
    "昆明": {
        "gbk": "%C0%A5%C3%F7",
        "code": "530100",
        "districts": {
            "全市": None,
            "五华区": "530102",
            "盘龙区": "530103",
            "官渡区": "530111",
            "西山区": "530112",
            "东川区": "530113",
            "呈贡区": "530114",
            "晋宁区": "530115",
            "富民县": "530124",
            "宜良县": "530125",
            "石林彝族自治县": "530126",
            "嵩明县": "530127",
            "禄劝彝族苗族自治县": "530128",
            "寻甸回族彝族自治县": "530129",
            "安宁市": "530181",
        }
    },
    "深圳": {
        "gbk": "%C9%EE%DB%DA",
        "code": "440300",
        "districts": {
            "全市": None,
            "罗湖区": "440303",
            "福田区": "440304",
            "南山区": "440305",
            "宝安区": "440306",
            "龙岗区": "440307",
            "盐田区": "440308",
            "龙华区": "440309",
            "坪山区": "440310",
            "光明区": "440311",
            "大鹏新区": "440312",
        }
    }
}

CATEGORY_CONFIG = {
    "住宅用房": 50025969,
    "商业用房": 200782003,
    "工业用房": 200788003,
    "其他用房": 200798003,
}

BIZ_TYPE_CONFIG = {
    "不限资产类型": -1,
    "涉刑资产": 4,
    "诉讼资产": 6,
    "破产资产": 8,
    "自行处置": 10,
}

# ============================================================
# 日志
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
    script_dir = get_script_dir()
    if browser_type == "edge":
        return os.path.join(script_dir, "msedgedriver.exe")
    return os.path.join(script_dir, "chromedriver.exe")

def download_chromedriver():
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "chromedriver.exe")
    if os.path.isfile(target):
        return target
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True, stderr=subprocess.STDOUT
        )
        version = re.search(r'REG_SZ\s+([\d.]+)', output.decode()).group(1)
    except:
        try:
            result = subprocess.run(
                ["powershell", "-Command",
                 "(Get-Item 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe').VersionInfo.ProductVersion"],
                capture_output=True, text=True, timeout=15
            )
            version = result.stdout.strip()
        except:
            logger.error("无法获取 Chrome 版本")
            return None

    logger.info(f"Chrome 版本: {version}")
    major = version.split('.')[0]
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
    script_dir = get_script_dir()
    target = os.path.join(script_dir, "msedgedriver.exe")
    if os.path.isfile(target):
        return target
    version = None
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
        logger.error("无法获取 Edge 版本")
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
    logger.error("自动下载 EdgeDriver 失败")
    return None

def detect_browser():
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

    if shutil.which("msedge"): return "edge", None
    if shutil.which("google-chrome") or shutil.which("chrome"): return "chrome", None
    return None, None

# 控制台单键输入支持
def getch():
    if sys.platform == 'win32':
        import msvcrt
        while True:
            ch = msvcrt.getch()
            if ch == b'\xe0':  # 功能键前导字节，再读一次
                msvcrt.getch()
                continue
            return ch.decode('utf-8', errors='ignore')
    else:
        import tty, termios
        fd = sys.stdin.fileno()
        old = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            return sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old)

# ============================================================
# 菜单工具
# ============================================================

def menu_select(descriptions, prompt="请选择:", allow_back=True):
    """单键菜单选择，选项自动从1开始编号，0为返回/退出。
       参数:
           descriptions: 选项描述列表
           prompt: 提示文字
           allow_back: True 表示0为“返回上一步”，False 表示0为“退出”
       返回:
           选项索引 (0-based)，-1 表示返回/退出
    """
    while True:
        print(f"\n{prompt}")
        for idx, desc in enumerate(descriptions, 1):
            print(f"  [{idx}] {desc}")
        if allow_back:
            print("  [0] 返回上一步")
        else:
            print("  [0] 退出")
        key = getch()
        if key.isdigit():
            num = int(key)
            if 1 <= num <= len(descriptions):
                return num - 1  # 转为索引
            elif num == 0:
                return -1
        print("无效输入，请重新按键")

def input_with_default(prompt, default):
    """输入带默认值的字符串"""
    val = input(f"{prompt} (默认为 {default}): ").strip()
    return val if val else default

def confirm(prompt="确认? (Y/N): "):
    """确认提示，返回 True/False"""
    while True:
        print(prompt, end="", flush=True)
        if sys.platform == 'win32':
            ch = getch().upper()
            print(ch)
            if ch == 'Y':
                return True
            elif ch == 'N':
                return False
        else:
            ans = input().strip().upper()
            if ans == 'Y':
                return True
            elif ans == 'N':
                return False
        print("请输入 Y 或 N")

# ============================================================
# URL 生成
# ============================================================

def build_list_url(city, category, district, biz_type, date_start, date_end):
    """根据筛选参数构建列表页 URL"""
    city_info = CITY_CONFIG[city]
    gbk_city = city_info["gbk"]
    category_id = CATEGORY_CONFIG[category]
    base = f"https://sf.taobao.com/list/{category_id}__2___{gbk_city}.htm"
    params = {
        "auction_source": "0",
        "st_param": "-1",
        "auction_start_seg": "0",
        "auction_start_from": date_start,
        "auction_start_to": date_end,
    }
    if biz_type != -1:
        params["item_biz_type"] = biz_type
    district_code = city_info["districts"].get(district)
    if district_code:
        params["location_code"] = district_code
    return base + "?" + urllib.parse.urlencode(params)

# ============================================================
# 配置管理
# ============================================================

DEFAULT_CONFIG = {
    "cookie": "",
    "city": "昆明",
    "category": "住宅用房",
    "district": "全市",
    "biz_type": "不限资产类型",
    "date_start": "2024-01-01",
    "date_end": "2026-03-31",
    "mode": "full",
    "task_status": "not_started",
    "list_progress": {
        "last_page": 0,
        "all_urls": []
    },
    "detail_progress": {
        "visited_urls": [],
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
                for key, val in DEFAULT_CONFIG.items():
                    if key not in config:
                        config[key] = val
                return config
            except:
                pass
        return dict(DEFAULT_CONFIG)

    def save(self):
        self.data["last_update"] = datetime.now().isoformat()
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def is_task_completed(self):
        return self.data.get("task_status") == "completed"

    def is_task_in_progress(self):
        return self.data.get("task_status") == "in_progress"

    def mark_completed(self):
        self.data["task_status"] = "completed"
        self.data["list_progress"]["all_urls"] = []
        self.data["list_progress"]["last_page"] = 0
        self.data["detail_progress"]["visited_urls"] = []
        self.save()

    def reset_for_new_task(self):
        cookie = self.data.get("cookie", "")
        city = self.data.get("city", "昆明")
        category = self.data.get("category", "住宅用房")
        district = self.data.get("district", "全市")
        biz_type = self.data.get("biz_type", "不限资产类型")
        date_start = self.data.get("date_start", "2024-01-01")
        date_end = self.data.get("date_end", "2026-03-31")
        self.data = dict(DEFAULT_CONFIG)
        self.data["cookie"] = cookie
        self.data["city"] = city
        self.data["category"] = category
        self.data["district"] = district
        self.data["biz_type"] = biz_type
        self.data["date_start"] = date_start
        self.data["date_end"] = date_end
        self.save()

# ============================================================
# 列表页爬取
# ============================================================

class ListPageScraper:
    def __init__(self, driver):
        self.driver = driver

    def fetch_page(self, page_url):
        """传入完整 URL，爬取一页数据"""
        logger.info(f"列表页: {page_url[:90]}...")
        try:
            self.driver.get(page_url)
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

        page_source = self.driver.page_source
        m = re.search(r'id="sf-item-list-data"[^>]*>(.*?)</script>', page_source, re.DOTALL)
        if m:
            try:
                data = json.loads(m.group(1).strip())
                items = data.get("data", [])
                if items:
                    return "ok", self._normalize(items)
            except: pass

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
# 详情页爬取（与 v0.17 完全一致）
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
            "end_time": "",
            "deal_price": "",
            "assessment_price": "",
            "basic_info_table": [],
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
            return result

        cur_url = self.driver.current_url
        if "deny" in cur_url or "punish" in cur_url:
            return "blocked", result
        if "login" in cur_url:
            return "need_login", result

        result["title"], result["auction_type"] = self._extract_title_and_type()
        result["end_time"] = self._extract_end_time()
        result["deal_price"] = self._extract_deal_price()
        result["assessment_price"] = self._extract_assessment_price()
        status_info = self._extract_status()
        result["status"] = status_info["text"]
        result["is_success"] = status_info["success"]
        result["basic_info_table"] = self._extract_basic_info_table()
        result["intro_table"] = self._extract_intro_table()
        result["attachments"] = self._extract_attachments()
        return "ok", result

    def _extract_title_and_type(self):
        title = ""
        auction_type = ""
        try:
            h1 = self.driver.find_element(By.TAG_NAME, "h1")
            full_text = h1.text.strip()
            status_spans = h1.find_elements(By.CSS_SELECTOR, "span.item-status")
            if status_spans:
                auction_type = status_spans[0].text.strip()
                title = full_text.replace(auction_type, "").strip()
            else:
                title = full_text
        except NoSuchElementException:
            pass
        if not title:
            title = self._extract_title_fallback()
        if not auction_type:
            auction_type = self._extract_auction_type_fallback(title)
        return title, auction_type

    def _extract_title_fallback(self):
        try:
            return self.driver.title.split("-")[0].strip()
        except:
            return ""

    def _extract_auction_type_fallback(self, title):
        for kw in ["一拍", "二拍", "三拍", "变卖"]:
            if kw in title: return kw
        body = self.driver.find_element(By.TAG_NAME, "body").text
        m = re.search(r'(一拍|二拍|三拍|变卖)', body[:500])
        if m: return m.group(1)
        return ""

    def _extract_end_time(self):
        try:
            elem = self.driver.find_element(By.CSS_SELECTOR, "span.countdown.J_TimeLeft")
            return elem.text.strip()
        except NoSuchElementException:
            return ""

    def _extract_deal_price(self):
        try:
            elem = self.driver.find_element(By.CSS_SELECTOR, "span.pm-current-price em")
            return elem.text.strip().replace(",", "") + "元"
        except NoSuchElementException:
            return "无"

    def _extract_assessment_price(self):
        try:
            table_a = self.driver.find_element(By.CSS_SELECTOR, "table[pm-ui*='rule-info']")
            rows = table_a.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                keys = row.find_elements(By.CSS_SELECTOR, "span.key--lGkywRDE")
                for key_elem in keys:
                    key_text = key_elem.text.strip()
                    if any(kw in key_text for kw in ["评估价", "市场价", "询价"]):
                        parent_td = key_elem.find_element(By.XPATH, "./ancestor::td[1]")
                        try:
                            price_elem = parent_td.find_element(By.CSS_SELECTOR, "span.family-tahoma")
                            return price_elem.text.strip().replace(",", "") + "元"
                        except:
                            pass
        except NoSuchElementException:
            pass

        try:
            table_b = self._find_intro_table()
            if table_b:
                rows = table_b.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    text = row.text
                    if "询价总价" in text or "网络询价结果" in text:
                        nums = re.findall(r'[\d.]+', text)
                        if nums:
                            return nums[-1] + "元"
        except:
            pass
        return "无"

    def _extract_status(self):
        status_text = ""
        try:
            elem = self.driver.find_element(By.CSS_SELECTOR, "span.firstDesc--G_zP45iR")
            status_text = elem.text.strip()
        except NoSuchElementException:
            pass

        if not status_text:
            body = self.driver.find_element(By.TAG_NAME, "body").text
            for kw in FAIL_KEYWORDS:
                if kw in body[:2000]:
                    return {"text": kw, "success": False}
            for kw in SUCCESS_KEYWORDS:
                if kw in body[:2000]:
                    return {"text": kw, "success": True}
            status_text = "未知"

        is_success = "已结束" in status_text and not any(fk in status_text for fk in ["流拍", "无人出价"])
        return {"text": status_text, "success": is_success}

    def _extract_basic_info_table(self):
        try:
            table_a = self.driver.find_element(By.CSS_SELECTOR, "table[pm-ui*='rule-info']")
            return self._parse_table_to_grid(table_a)
        except NoSuchElementException:
            return []

    def _find_intro_table(self):
        tables = self.driver.find_elements(By.TAG_NAME, "table")
        for table in tables:
            if "标的物情况调查表" in table.text:
                return table
        try:
            intro_headers = self.driver.find_elements(By.XPATH, "//*[contains(text(), '标的物介绍')]")
            if intro_headers:
                following_tables = intro_headers[0].find_elements(By.XPATH, "./following::table")
                if following_tables:
                    return following_tables[0]
        except:
            pass
        try:
            return self.driver.find_element(By.CSS_SELECTOR, "#J_desc > table")
        except NoSuchElementException:
            return None

    def _extract_intro_table(self):
        table_rows = []
        raw_text = ""
        target_table = self._find_intro_table()
        if target_table:
            table_rows = self._parse_table_to_grid(target_table)
            raw_text = target_table.text
        else:
            body = self.driver.find_element(By.TAG_NAME, "body").text
            intro_start = -1
            for mk in ["标的物介绍", "拍卖标的基本情况", "标的物详情"]:
                idx = body.find(mk)
                if idx != -1:
                    intro_start = idx + len(mk)
                    break
            if intro_start != -1:
                intro = body[intro_start:]
                for em in ["竞买公告", "竞买须知", "拍卖规则", "附件", "相关附件",
                           "重要提示", "竞买人条件", "咨询时间", "拍卖余款"]:
                    idx = intro.find(em)
                    if idx != -1: intro = intro[:idx]
                raw_text = intro.strip()
                table_rows = self._fallback_text_parse(intro)
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

    def _fallback_text_parse(self, text):
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        data = {}
        i = 0
        while i < len(lines):
            line = lines[i]
            m = re.match(r'^([^：:]+)[：:]\s*(.+)$', line)
            if m:
                data[m.group(1).strip()] = m.group(2).strip()
            elif i+1 < len(lines):
                nxt = lines[i+1]
                if len(line) < 20 and not re.match(r'^[\d.]+', line) and nxt and line != nxt:
                    data[line] = nxt
                    i += 1
            i += 1
        return [["项目", "内容"]] + [[k, v] for k, v in data.items()] if data else []

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
        self.saved_urls = []

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

        info_lines = [
            f"拍卖物名称：{title}",
            f"拍卖类型：{detail.get('auction_type')}",
            f"拍卖状态：{detail.get('status')}",
            f"是否成交：{'是' if detail.get('is_success') else '否'}",
            f"成交日期：{detail.get('end_time')}",
            f"成交价：{detail.get('deal_price')}",
            f"评估价：{detail.get('assessment_price')}",
            f"详情页链接：{detail.get('url')}",
            f"爬取时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
        ]
        with open(os.path.join(item_dir, "拍卖信息.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(info_lines))
        self.saved_urls.append(detail.get("url", ""))

        basic_rows = detail.get("basic_info_table", [])
        if basic_rows:
            with open(os.path.join(item_dir, "基本信息.csv"), "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerows(basic_rows)

        intro = detail.get("intro_table", {})
        rows = intro.get("table_rows", [])
        if rows:
            with open(os.path.join(item_dir, "标的物介绍.csv"), "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerows(rows)

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

    def run(self, city, category, district, biz_type, date_start, date_end, mode):
        date_range_str = f"{date_start}_to_{date_end}"
        self._setup_log_file(date_range_str)

        if mode == "full":
            list_url = build_list_url(city, category, district, biz_type, date_start, date_end)
            self._run_full(list_url, date_range_str)
        else:
            self._run_detail_only(date_range_str)

        self._print_stats()
        self._save_urls_summary(date_range_str)

    def _run_full(self, list_base_url, date_range_str):
        config = self.config.data
        list_prog = config.get("list_progress", {})
        last_page = list_prog.get("last_page", 0)
        all_urls = list_prog.get("all_urls", [])
        start_page = last_page + 1 if last_page > 0 else 1

        logger.info("阶段1：收集拍品 URL")
        user_aborted = False
        for page in range(start_page, MAX_PAGES + 1):
            if "?" in list_base_url:
                page_url = list_base_url + f"&page={page}"
            else:
                page_url = list_base_url + f"?page={page}"
            status, items = self.list_scraper.fetch_page(page_url)
            if status == "blocked":
                if self._handle_blocked():
                    status, items = self.list_scraper.fetch_page(page_url)
                else:
                    user_aborted = True
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
            list_prog["last_page"] = page
            list_prog["all_urls"] = all_urls
            self.config.save()
            self.stats["total_urls"] = len(all_urls)
            if len(items) < ITEMS_PER_PAGE:
                break
            random_delay()

        if user_aborted:
            logger.info("列表爬取被用户中断，进度已保存")
            self.config.data["task_status"] = "in_progress"
            self.config.save()
            return

        all_urls = self._deduplicate_urls(all_urls)
        logger.info(f"去重后 URL 数量: {len(all_urls)}")
        self.config.data["list_progress"]["all_urls"] = all_urls
        self.config.data["list_progress"]["last_page"] = MAX_PAGES
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
        self.config.data["list_progress"]["all_urls"] = urls
        self.config.save()
        self._process_details(urls, date_range_str)

    def _process_details(self, urls, date_range_str):
        config = self.config.data
        detail_prog = config.get("detail_progress", {})
        visited = set(detail_prog.get("visited_urls", []))
        saved_count = detail_prog.get("saved_count", 0)
        self.stats["total_saved"] = saved_count

        pending = [u for u in urls if u not in visited]
        total = len(pending)
        logger.info(f"阶段2：处理详情页 ({total} 个待处理)")

        completed_normally = True
        consecutive_fail = 0
        RETRYABLE = ("blocked", "need_login", "timeout")

        for idx, url in enumerate(pending):
            logger.info(f"[{idx+1}/{total}] {url[:80]}")
            random_delay()

            result = self.detail_scraper.fetch_detail(url)
            if isinstance(result, tuple):
                status, detail = result
            else:
                status = "error"
                detail = result

            if status in RETRYABLE:
                consecutive_fail += 1
                if consecutive_fail >= 3:
                    if self._handle_blocked():
                        consecutive_fail = 0
                        result = self.detail_scraper.fetch_detail(url)
                        if isinstance(result, tuple):
                            status, detail = result
                        else:
                            status = "error"
                            detail = result
                        if status in RETRYABLE:
                            logger.warning("重试后仍失败，跳过本次（未标记已处理）")
                            self.config.save()
                            continue
                    else:
                        completed_normally = False
                        break
                else:
                    logger.warning(f"异常状态 {status}，未标记已处理，将重试")
                    self.config.save()
                    continue

            elif status == "error":
                self.stats["total_failed"] += 1
                consecutive_fail += 1
                if consecutive_fail >= 10:
                    logger.error("连续10次失败，暂停")
                    completed_normally = False
                    break
                self.config.save()
                continue

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

            detail_prog["visited_urls"] = list(visited)
            detail_prog["saved_count"] = self.stats["total_saved"]
            self.config.save()

        if completed_normally:
            self.config.mark_completed()
            logger.info("所有 URL 处理完毕，任务完成")
        else:
            self.config.data["task_status"] = "in_progress"
            self.config.save()
            logger.info("详情处理中断，进度已保存")

    def _save_urls_summary(self, date_range_str):
        if self.saved_urls:
            summary_dir = os.path.join(self.output_root, date_range_str)
            urls_file = os.path.join(summary_dir, "urls.txt")
            with open(urls_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.saved_urls))
            logger.info(f"URL 汇总已保存: {urls_file}")

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
            intro_csv = os.path.join(d, "标的物介绍.csv")
            if os.path.isfile(intro_csv):
                try:
                    with open(intro_csv, "r", encoding="utf-8-sig") as f:
                        reader = csv.reader(f)
                        for line in reader:
                            if line:
                                row["标的物摘要"] = line[0][:50]
                                break
                except:
                    pass
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
        try:
            self.config.save()
        except:
            pass

# ============================================================
# ASCII 艺术字
# ============================================================

BANNER = r"""
      ___           ___                       ___           ___           ___           ___     
     /\  \         /\__\          ___        /\__\         /\  \         /\__\         /\  \    
    /::\  \       /:/  /         /\  \      /::|  |       /::\  \       /::|  |       /::\  \   
   /:/\ \  \     /:/__/          \:\  \    /:|:|  |      /:/\:\  \     /:|:|  |      /:/\:\  \  
  _\:\~\ \  \   /::\  \ ___      /::\__\  /:/|:|  |__   /::\~\:\  \   /:/|:|  |__   /:/  \:\  \ 
 /\ \:\ \ \__\ /:/\:\  /\__\  __/:/\/__/ /:/ |:| /\__\ /:/\:\ \:\__\ /:/ |:| /\__\ /:/__/ \:\__\
 \:\ \:\ \/__/ \/__\:\/:/  / /\/:/  /    \/__|:|/:/  / \/__\:\/:/  / \/__|:|/:/  / \:\  \ /:/  /
  \:\ \:\__\        \::/  /  \::/__/         |:/:/  /       \::/  /      |:/:/  /   \:\  /:/  / 
   \:\/:/  /        /:/  /    \:\__\         |::/  /        /:/  /       |::/  /     \:\/:/  /  
    \::/  /        /:/  /      \/__/         /:/  /        /:/  /        /:/  /       \::/  /   
     \/__/         \/__/                     \/__/         \/__/         \/__/         \/__/    
"""

def install_dependencies():
    print("正在安装依赖：selenium requests beautifulsoup4 ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "requests", "beautifulsoup4"])
        return True
    except Exception as e:
        print(f"安装失败: {e}")
        return False

def run_task_selection(config):
    """任务参数选择菜单，返回 (city, category, district, biz_type, date_start, date_end, mode)"""
    # 选择城市
    cities = list(CITY_CONFIG.keys())
    city_idx = menu_select(cities, "选择城市:")
    if city_idx == -1: return None
    city = cities[city_idx]

    # 选择房屋类型
    categories = list(CATEGORY_CONFIG.keys())
    cat_idx = menu_select(categories, "选择房屋类型:")
    if cat_idx == -1: return run_task_selection(config)
    category = categories[cat_idx]

    # 选择区县
    districts = list(CITY_CONFIG[city]["districts"].keys())
    dist_idx = menu_select(districts, "选择区县:")
    if dist_idx == -1: return run_task_selection(config)
    district = districts[dist_idx]

    # 选择资产来源
    biz_types = list(BIZ_TYPE_CONFIG.keys())
    biz_idx = menu_select(biz_types, "选择资产来源:")
    if biz_idx == -1: return run_task_selection(config)
    biz_type = biz_types[biz_idx]

    # 输入日期
    print()
    date_start = input_with_default("开始日期 (YYYY-MM-DD)", "2024-01-01")
    date_end = input_with_default("结束日期 (YYYY-MM-DD)", "2026-03-31")

    # 模式选择
    modes = ["完整模式（列表页 + 详情页）", "仅详情页（需要同目录下有 urls.txt）"]
    mode_idx = menu_select(modes, "选择运行模式:")
    if mode_idx == -1: return run_task_selection(config)
    mode = "detail_only" if mode_idx == 1 else "full"

    # 确认
    print(f"\n当前设置:")
    print(f"  城市: {city}")
    print(f"  房屋类型: {category}")
    print(f"  区县: {district}")
    print(f"  资产来源: {biz_type}")
    print(f"  日期: {date_start} ~ {date_end}")
    print(f"  模式: {mode}")
    if not confirm("确认开始？(Y/N): "):
        return run_task_selection(config)
    return city, category, district, biz_type, date_start, date_end, mode

def main():
    print(BANNER)
    print("阿里司法拍卖 - 法拍房数据爬取工具 v0.18.1\n")

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

    config = ConfigManager()

    # 主菜单
    while True:
        options = []
        if config.is_task_in_progress():
            options.append("恢复未完成任务")
        options.append("新建任务")
        choice = menu_select(options, "主菜单", allow_back=False)
        if choice == -1:  # 0 退出
            sys.exit(0)

        # 确定是新建还是恢复
        if options[choice] == "恢复未完成任务":
            # 直接使用配置中的参数
            city = config.data["city"]
            category = config.data["category"]
            district = config.data["district"]
            biz_type = config.data["biz_type"]
            date_start = config.data["date_start"]
            date_end = config.data["date_end"]
            mode = config.data["mode"]
            print(f"\n恢复任务：{city} {category} {district} {biz_type} {date_start} ~ {date_end} {mode}")
            break
        else:
            # 新建任务
            config.reset_for_new_task()
            params = run_task_selection(config)
            if params is None:
                continue  # 用户取消了选择，回到主菜单
            city, category, district, biz_type, date_start, date_end, mode = params
            config.data["city"] = city
            config.data["category"] = category
            config.data["district"] = district
            config.data["biz_type"] = biz_type
            config.data["date_start"] = date_start
            config.data["date_end"] = date_end
            config.data["mode"] = mode
            config.data["task_status"] = "in_progress"
            config.save()
            break

    scraper = SfTaobaoScraper(config, browser_type, driver_path)
    try:
        scraper.start()
        scraper.run(city, category, district, biz_type, date_start, date_end, mode)
        date_range_str = f"{date_start}_to_{date_end}"
        scraper.save_summary(date_range_str)
    except KeyboardInterrupt:
        logger.info("用户中断，进度已保存")
        date_range_str = f"{date_start}_to_{date_end}"
        scraper.save_summary(date_range_str)
    except Exception as e:
        logger.error(f"异常退出: {e}")
        import traceback; traceback.print_exc()
        date_range_str = f"{date_start}_to_{date_end}"
        scraper.save_summary(date_range_str)
    finally:
        scraper.stop()
    print("\n✅ 程序结束")

if __name__ == "__main__":
    main()