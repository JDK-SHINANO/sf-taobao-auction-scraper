#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阿里司法拍卖 - 昆明市法拍房数据爬取工具
目标：爬取 sf.taobao.com 昆明市已成交的法拍住宅数据
日期范围：2024年1月 ~ 2026年3月（可自定义）

使用方法：
  1. 安装依赖：pip install requests beautifulsoup4 lxml openpyxl selenium webdriver-manager
  2. 运行脚本：python sf_taobao_scraper.py
  3. 按提示输入 Cookie 和日期范围
  4. 等待爬取完成，结果保存到 Excel 文件

如何获取 Cookie：
  用 Chrome 浏览器打开 https://sf.taobao.com/ 并登录淘宝账号
  按 F12 → Network → 刷新页面 → 点任意请求 → Headers → 复制 Cookie

⚠️ 重要说明：
  sf.taobao.com 的详情页使用 React 渲染，纯 requests 无法获取完整 HTML。
  本脚本提供两种模式：
    模式1（推荐）：Selenium 模拟浏览器 - 完整获取所有字段
    模式2（轻量）：仅用 requests - 从列表页 JSON 提取部分字段
"""

import re
import os
import sys
import json
import time
import random
import logging
from datetime import datetime
from urllib.parse import urljoin, quote
from pathlib import Path

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
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

try:
    from webdriver_manager.chrome import ChromeDriverManager
    HAS_WDM = True
except ImportError:
    HAS_WDM = False

# ============================================================
# 配置区
# ============================================================

# 昆明市区域代码映射
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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
]


# ============================================================
# 工具函数
# ============================================================

def random_delay(low=None, high=None):
    time.sleep(random.uniform(low or MIN_DELAY, high or MAX_DELAY))

def clean_number(text):
    if not text: return ""
    return re.sub(r"[,\s，]", "", text)

def wan_to_yuan(text):
    if not text: return ""
    text = str(text).strip()
    m = re.search(r"([\d,.]+)\s*万", text)
    if m:
        return str(int(float(m.group(1).replace(",", "")) * 10000))
    m = re.search(r"([\d,.]+)", text)
    if m:
        return clean_number(m.group(1))
    return ""

def extract_date_mm_dd(text):
    if not text: return ""
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m: return f"{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    if m: return f"{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    return ""

def parse_date_to_obj(text):
    if not text: return None
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m:
        try: return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return None
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    if m:
        try: return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return None
    return None

def extract_district_from_title(title):
    if not title: return ""
    for district in KUNMING_DISTRICTS:
        if district in title: return district
    for district in KUNMING_DISTRICTS:
        short = district.replace("区", "").replace("县", "").replace("市", "")
        if short and short in title: return district
    return ""


# ============================================================
# 列表页爬取（requests）
# ============================================================

class ListPageScraper:
    """使用 requests 爬取列表页"""

    def __init__(self, cookie):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Referer": "https://sf.taobao.com/",
            "Cookie": cookie,
        })

    def _rotate_ua(self):
        self.session.headers["User-Agent"] = random.choice(USER_AGENTS)

    def _get(self, url, params=None, max_retries=3):
        for attempt in range(max_retries):
            try:
                self._rotate_ua()
                resp = self.session.get(url, params=params, timeout=30)
                if resp.status_code == 200:
                    return resp
                elif resp.status_code in (403, 429):
                    logger.warning(f"HTTP {resp.status_code}，等待重试...")
                    time.sleep(random.uniform(10, 20))
                else:
                    logger.warning(f"HTTP {resp.status_code}，重试 {attempt+1}/{max_retries}")
                    time.sleep(random.uniform(5, 10))
            except requests.exceptions.RequestException as e:
                logger.warning(f"请求异常: {e}，重试 {attempt+1}/{max_retries}")
                time.sleep(random.uniform(5, 10))
        return None

    def fetch_page(self, page, completed=True):
        """获取列表页数据，返回拍品列表"""
        params = {
            "location_code": KUNMING_LOCATION_CODE,
            "category": CATEGORY_RESIDENTIAL,
            "auction_source": "0",
            "city": "",
            "province": KUNMING_PROVINCE_CODE,
            "st_param": "-1",
            "auction_start_seg": "-1",
            "page": str(page),
        }
        if completed:
            params["completed_items"] = "1"

        resp = self._get(LIST_URL, params=params)
        if not resp:
            logger.error(f"列表页第 {page} 页获取失败")
            return []

        if "deny" in resp.url or "punish" in resp.url:
            logger.error("触发反爬验证，请检查 Cookie 是否有效")
            return []

        soup = BeautifulSoup(resp.text, "lxml")

        # 从 script#sf-item-list-data 提取 JSON
        script_tag = soup.find("script", {"id": "sf-item-list-data"})
        if script_tag and script_tag.string:
            try:
                data = json.loads(script_tag.string.strip())
                items = data.get("data", [])
                if items:
                    logger.info(f"列表页第 {page} 页：{len(items)} 条记录")
                    return items
            except json.JSONDecodeError:
                pass

        # 备用：从 HTML 解析
        items = []
        for a in soup.find_all("a", href=re.compile(r"(?:susong-item|sf-item).*\.htm")):
            href = a.get("href", "")
            if not href.startswith("http"):
                href = "https:" + href if href.startswith("//") else "https://sf.taobao.com" + href
            title = a.get_text(strip=True)
            item_id_m = re.search(r"(\d+)\.htm", href)
            items.append({
                "itemUrl": href,
                "title": title,
                "id": item_id_m.group(1) if item_id_m else "",
                "currentPrice": "",
                "status": "",
                "statusText": "",
            })
        if items:
            logger.info(f"列表页第 {page} 页：{len(items)} 条记录 (HTML解析)")
        return items


# ============================================================
# 详情页爬取 - Selenium 模式
# ============================================================

class SeleniumDetailScraper:
    """使用 Selenium 爬取详情页（推荐方式）"""

    def __init__(self, cookie_str):
        self.cookie_str = cookie_str
        self.driver = None

    def start(self):
        """启动浏览器"""
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
        chrome_options.add_argument("--window-size=1920,1080")
        # 禁用图片加载以加快速度
        prefs = {"profile.managed_default_content_settings.images": 2}
        chrome_options.add_experimental_option("prefs", prefs)

        if HAS_WDM:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        else:
            self.driver = webdriver.Chrome(options=chrome_options)

        # 隐藏 webdriver 特征
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })

        # 先访问主站设置 cookie
        self.driver.get("https://sf.taobao.com/")
        time.sleep(2)

        # 设置 cookie
        if self.cookie_str:
            for cookie_item in self.cookie_str.split(";"):
                cookie_item = cookie_item.strip()
                if "=" in cookie_item:
                    name, value = cookie_item.split("=", 1)
                    try:
                        self.driver.add_cookie({
                            "name": name.strip(),
                            "value": value.strip(),
                            "domain": ".taobao.com",
                            "path": "/",
                        })
                    except Exception:
                        pass

        logger.info("Selenium 浏览器已启动")

    def stop(self):
        if self.driver:
            self.driver.quit()
            logger.info("Selenium 浏览器已关闭")

    def fetch_detail(self, url) -> dict:
        """获取详情页完整数据"""
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
            # 等待页面关键元素加载
            WebDriverWait(self.driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(3)  # 额外等待 React 渲染

            page_text = self.driver.find_element(By.TAG_NAME, "body").text
            page_source = self.driver.page_source

            # --- 标的名称 ---
            try:
                h1 = self.driver.find_element(By.CSS_SELECTOR, "h1, .pm-main h1, .title h1, .item-title")
                result["标的名称"] = h1.text.strip()
            except NoSuchElementException:
                # 从 title 标签获取
                title_tag = self.driver.title
                if title_tag:
                    result["标的名称"] = title_tag.split("-")[0].strip()

            # --- 成交价 ---
            price_patterns = [
                r"成交价\s*[¥￥]?\s*([\d,.]+)",
                r"成交金额\s*[¥￥]?\s*([\d,.]+)",
                r"最终成交价\s*[¥￥]?\s*([\d,.]+)",
            ]
            for pat in price_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["成交价(元)"] = clean_number(m.group(1))
                    break
            if not result["成交价(元)"]:
                m = re.search(r"成交[金价].*?([\d.]+)\s*万", page_text)
                if m:
                    result["成交价(元)"] = str(int(float(m.group(1)) * 10000))

            # --- 评估价/市场价 ---
            eval_patterns = [
                r"评估价\s*[¥￥]?\s*([\d,.]+)",
                r"市场价\s*[¥￥]?\s*([\d,.]+)",
                r"评估总价\s*[¥￥]?\s*([\d,.]+)",
            ]
            for pat in eval_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["评估价/市场价(元)"] = clean_number(m.group(1))
                    break
            if not result["评估价/市场价(元)"]:
                m = re.search(r"(?:评估|市场)[金价].*?([\d.]+)\s*万", page_text)
                if m:
                    result["评估价/市场价(元)"] = str(int(float(m.group(1)) * 10000))

            # --- 成交日期 ---
            date_patterns = [
                r"结束时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
                r"成交时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
                r"拍卖结束\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
                r"竞价结束\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
            ]
            for pat in date_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["成交日期"] = extract_date_mm_dd(m.group(1))
                    self._last_raw_date = m.group(1)
                    break

            # --- 建筑面积 ---
            area_patterns = [
                r"建筑总面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
                r"建筑面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
                r"总面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
                r"房屋面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
            ]
            for pat in area_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["建筑面积(㎡)"] = m.group(1)
                    break

            # --- 土地分摊面积 ---
            land_patterns = [
                r"土地分摊面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
                r"土地面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
                r"分摊面积\s*([\d.]+)\s*(?:㎡|平方米|m²|m2)",
            ]
            for pat in land_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["土地分摊面积(㎡)"] = m.group(1)
                    break

            # --- 所在层/总楼层 ---
            floor_patterns = [
                r"(?:所在层|楼层)\s*[:：]?\s*(?:第)?(\d+)\s*层\s*/\s*(?:共)?(\d+)\s*层",
                r"(?:所在层|楼层)\s*[:：]?\s*(\d+)\s*/\s*(\d+)",
                r"第(\d+)层\s*/\s*共(\d+)层",
            ]
            for pat in floor_patterns:
                m = re.search(pat, page_text)
                if m:
                    result["所在层/总楼层"] = f"{m.group(1)}层/{m.group(2)}层"
                    break

            # --- 朝向 ---
            m = re.search(r"朝\s*向\s*[:：]?\s*([东南西北中\s,，、]+)", page_text)
            if m:
                result["朝向"] = m.group(1).strip().rstrip("，,")
            else:
                m = re.search(r"朝向\s*[:：]?\s*(\S+)", page_text)
                if m:
                    result["朝向"] = m.group(1).strip()

            # --- 房屋结构 ---
            m = re.search(r"(?:建筑结构|房屋结构)\s*[:：]?\s*(\S+)", page_text)
            if m:
                result["房屋结构"] = m.group(1).strip()

            # --- 户型 ---
            m = re.search(r"(?:房屋户型|户型)\s*[:：]?\s*([^\n,，。]+)", page_text)
            if m:
                result["户型"] = m.group(1).strip()
            else:
                m = re.search(r"(\d+室\d*厅?\d*卫?)", page_text)
                if m:
                    result["户型"] = m.group(1)

            # --- 租赁情况 ---
            m = re.search(r"租赁情况\s*[:：]?\s*([^\n,，。]+)", page_text)
            if m:
                result["租赁情况"] = m.group(1).strip()

            # --- 装修情况 ---
            m = re.search(r"装修情况\s*[:：]?\s*([^\n,，。]+)", page_text)
            if m:
                result["装修情况"] = m.group(1).strip()

            # --- 附件名称 ---
            m = re.search(r"相关附件下载[：:]\s*(\S+\.(?:pdf|docx?|xls|xlsx|zip|rar))", page_text, re.IGNORECASE)
            if m:
                result["附件名称"] = m.group(1).strip()
            else:
                # 从链接中找
                try:
                    links = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='.pdf'], a[href*='.doc'], a[href*='.xls']")
                    for link in links:
                        href = link.get_attribute("href") or ""
                        name = link.text.strip()
                        if re.search(r"\.(pdf|docx?|xlsx?|zip|rar)$", href, re.IGNORECASE):
                            result["附件名称"] = name or href.split("/")[-1]
                            break
                except Exception:
                    pass

            # --- 所属区域 ---
            result["所属区域"] = extract_district_from_title(result["标的名称"])

        except TimeoutException:
            logger.warning(f"详情页加载超时: {url}")
        except Exception as e:
            logger.warning(f"详情页解析异常: {e}")

        return result


# ============================================================
# 详情页爬取 - Requests 模式（备用）
# ============================================================

class RequestsDetailScraper:
    """使用 requests 爬取详情页（备用方案，可能获取不到完整数据）"""

    def __init__(self, cookie):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "Referer": "https://sf.taobao.com/",
            "Cookie": cookie,
        })

    def fetch_detail(self, url) -> dict:
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
            if resp.status_code != 200:
                return result

            # 检查是否被拦截
            if "deny" in resp.url or "punish" in resp.url:
                return result

            soup = BeautifulSoup(resp.text, "lxml")
            text = soup.get_text(separator="\n", strip=True)

            # 标题
            h1 = soup.find("h1")
            if h1:
                result["标的名称"] = h1.get_text(strip=True)
            if not result["标的名称"]:
                title_tag = soup.find("title")
                if title_tag:
                    result["标的名称"] = title_tag.get_text(strip=True).split("-")[0].strip()

            # 尝试从 script 标签中提取数据（某些版本可能有内嵌 JSON）
            for script in soup.find_all("script"):
                script_text = script.string or ""
                # 尝试找 JSON 数据
                if "detailData" in script_text or "auctionInfo" in script_text:
                    m = re.search(r"(?:detailData|auctionInfo)\s*=\s*(\{.*?\});", script_text, re.DOTALL)
                    if m:
                        try:
                            detail_json = json.loads(m.group(1))
                            # 提取字段
                            if "title" in detail_json:
                                result["标的名称"] = detail_json["title"]
                            if "price" in detail_json:
                                result["成交价(元)"] = str(detail_json["price"])
                            break
                        except json.JSONDecodeError:
                            pass

            # 从页面文本提取字段（与 Selenium 模式相同的正则）
            for pat, key in [
                (r"成交价\s*[¥￥]?\s*([\d,.]+)", "成交价(元)"),
                (r"成交金额\s*[¥￥]?\s*([\d,.]+)", "成交价(元)"),
            ]:
                m = re.search(pat, text)
                if m:
                    result[key] = clean_number(m.group(1))
                    break

            for pat in [r"评估价\s*[¥￥]?\s*([\d,.]+)", r"市场价\s*[¥￥]?\s*([\d,.]+)"]:
                m = re.search(pat, text)
                if m:
                    result["评估价/市场价(元)"] = clean_number(m.group(1))
                    break

            for pat in [r"结束时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})", r"成交时间\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})"]:
                m = re.search(pat, text)
                if m:
                    result["成交日期"] = extract_date_mm_dd(m.group(1))
                    self._last_raw_date = m.group(1)
                    break

            for pat in [r"建筑总面积\s*([\d.]+)", r"建筑面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m:
                    result["建筑面积(㎡)"] = m.group(1)
                    break

            for pat in [r"土地分摊面积\s*([\d.]+)", r"土地面积\s*([\d.]+)"]:
                m = re.search(pat, text)
                if m:
                    result["土地分摊面积(㎡)"] = m.group(1)
                    break

            m = re.search(r"(?:所在层|楼层)\s*[:：]?\s*(?:第)?(\d+)\s*层\s*/\s*(?:共)?(\d+)\s*层", text)
            if m:
                result["所在层/总楼层"] = f"{m.group(1)}层/{m.group(2)}层"

            m = re.search(r"朝\s*向\s*[:：]?\s*([东南西北中\s,，、]+)", text)
            if m:
                result["朝向"] = m.group(1).strip().rstrip("，,")

            m = re.search(r"(?:建筑结构|房屋结构)\s*[:：]?\s*(\S+)", text)
            if m:
                result["房屋结构"] = m.group(1).strip()

            m = re.search(r"(?:房屋户型|户型)\s*[:：]?\s*([^\n,，。]+)", text)
            if m:
                result["户型"] = m.group(1).strip()

            m = re.search(r"租赁情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m:
                result["租赁情况"] = m.group(1).strip()

            m = re.search(r"装修情况\s*[:：]?\s*([^\n,，。]+)", text)
            if m:
                result["装修情况"] = m.group(1).strip()

            for a in soup.find_all("a", href=True):
                href = a.get("href", "")
                if re.search(r"\.(pdf|docx?|xlsx?)$", href, re.IGNORECASE):
                    result["附件名称"] = a.get_text(strip=True) or href.split("/")[-1]
                    break

            result["所属区域"] = extract_district_from_title(result["标的名称"])

        except Exception as e:
            logger.warning(f"详情页解析异常: {e}")

        return result


# ============================================================
# 主爬虫类
# ============================================================

class SfTaobaoScraper:
    """阿里司法拍卖爬虫主控"""

    def __init__(self, cookie, use_selenium=True):
        self.cookie = cookie
        self.use_selenium = use_selenium and HAS_SELENIUM
        self.list_scraper = ListPageScraper(cookie)
        self.detail_scraper = None
        self.results = []
        self.failed_urls = []
        self.visited_ids = set()

    def start(self):
        if self.use_selenium:
            self.detail_scraper = SeleniumDetailScraper(self.cookie)
            self.detail_scraper.start()
        else:
            self.detail_scraper = RequestsDetailScraper(self.cookie)

    def stop(self):
        if self.use_selenium and self.detail_scraper:
            self.detail_scraper.stop()

    def scrape_all(self, date_start="2024-01-01", date_end="2026-03-31"):
        dt_start = datetime.strptime(date_start, "%Y-%m-%d")
        dt_end = datetime.strptime(date_end, "%Y-%m-%d")

        logger.info("=" * 60)
        logger.info("阿里司法拍卖 - 昆明市法拍房数据爬取")
        logger.info(f"日期范围: {date_start} ~ {date_end}")
        logger.info(f"模式: {'Selenium (推荐)' if self.use_selenium else 'Requests (轻量)'}")
        logger.info("=" * 60)

        total_found = 0
        total_saved = 0
        total_skipped = 0
        out_of_range_count = 0

        for page in range(1, MAX_PAGES + 1):
            logger.info(f"\n--- 列表页第 {page} 页 ---")
            items = self.list_scraper.fetch_page(page)

            if not items:
                logger.info(f"第 {page} 页无数据，停止翻页")
                break

            for idx, item in enumerate(items):
                total_found += 1
                item_id = item.get("id", "")
                item_url = item.get("itemUrl", "")
                title = item.get("title", "")
                status = str(item.get("status", "") or item.get("statusText", ""))

                # 补全 URL
                if item_url and not item_url.startswith("http"):
                    item_url = "https:" + item_url if item_url.startswith("//") else "https://sf.taobao.com" + item_url

                # 去重
                if item_id and item_id in self.visited_ids:
                    continue
                if item_id:
                    self.visited_ids.add(item_id)

                # 过滤流拍
                fail_kw = ["流拍", "中止", "撤回", "失败", "取消"]
                if any(kw in status for kw in fail_kw):
                    total_skipped += 1
                    continue

                # 从列表页提取价格（备用）
                list_price = item.get("currentPrice", "")

                logger.info(f"  [{idx+1}/{len(items)}] {title[:60]}...")
                random_delay()

                # 爬取详情页
                detail = self.detail_scraper.fetch_detail(item_url)
                if not detail or not detail.get("标的名称"):
                    self.failed_urls.append(item_url)
                    continue

                # 补充列表页数据
                if not detail["标的名称"]:
                    detail["标的名称"] = title
                if not detail["成交价(元)"] and list_price:
                    detail["成交价(元)"] = wan_to_yuan(list_price)
                if not detail["所属区域"]:
                    detail["所属区域"] = extract_district_from_title(detail["标的名称"] or title)

                # 日期范围过滤
                raw_date = getattr(self.detail_scraper, "_last_raw_date", "")
                if raw_date:
                    dt = parse_date_to_obj(raw_date)
                    if dt:
                        if dt < dt_start:
                            out_of_range_count += 1
                            if out_of_range_count >= 15:
                                logger.info("连续多条早于起始日期，可能已超出范围，停止")
                                break
                            continue
                        elif dt > dt_end:
                            continue
                        else:
                            out_of_range_count = 0

                self.results.append(detail)
                total_saved += 1
                logger.info(f"    ✓ 已保存 (共{total_saved}条)")

            if out_of_range_count >= 15:
                break

            if len(items) < ITEMS_PER_PAGE:
                logger.info("当前页不足40条，已到最后一页")
                break

            random_delay()

        logger.info("\n" + "=" * 60)
        logger.info(f"爬取完成！总发现: {total_found} | 已保存: {total_saved} | 跳过: {total_skipped} | 失败: {len(self.failed_urls)}")
        logger.info("=" * 60)

    def save_to_excel(self, filename=None):
        if not self.results:
            logger.warning("没有数据可保存")
            return

        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"昆明法拍房数据_{timestamp}.xlsx"

        headers = [
            "标的名称", "成交价(元)", "评估价/市场价(元)", "成交日期",
            "建筑面积(㎡)", "土地分摊面积(㎡)", "所在层/总楼层",
            "朝向", "房屋结构", "户型", "租赁情况", "装修情况",
            "附件名称", "所属区域", "详情页链接",
        ]

        if not HAS_OPENPYXL:
            self._save_csv(filename.replace(".xlsx", ".csv"), headers)
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "昆明法拍房数据"

        hdr_font = Font(bold=True, size=11, color="FFFFFF")
        hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )

        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
            c.border = border

        for row, item in enumerate(self.results, 2):
            for col, key in enumerate(headers, 1):
                val = item.get(key, "")
                c = ws.cell(row=row, column=col, value=val)
                c.border = border
                c.alignment = Alignment(vertical="center", wrap_text=True)
                if "价" in key and val:
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    try:
                        c.value = int(val)
                        c.number_format = "#,##0"
                    except (ValueError, TypeError):
                        pass

        for col in range(1, len(headers) + 1):
            max_len = len(str(headers[col-1]))
            for row in range(2, min(len(self.results) + 2, 100)):
                cv = ws.cell(row=row, column=col).value
                if cv:
                    max_len = max(max_len, len(str(cv)))
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(max_len + 4, 50)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        wb.save(filename)
        logger.info(f"数据已保存: {filename} ({len(self.results)} 条)")

    def _save_csv(self, filename, headers):
        import csv
        with open(filename, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=headers)
            w.writeheader()
            for item in self.results:
                w.writerow({k: item.get(k, "") for k in headers})
        logger.info(f"数据已保存: {filename} (CSV, {len(self.results)} 条)")
        logger.info("提示: pip install openpyxl 可输出 Excel 格式")

    def save_failed_urls(self, filename=None):
        if not self.failed_urls:
            return
        if not filename:
            filename = f"failed_urls_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(filename, "w", encoding="utf-8") as f:
            for url in self.failed_urls:
                f.write(url + "\n")
        logger.info(f"失败URL已保存: {filename} ({len(self.failed_urls)} 条)")


# ============================================================
# 入口
# ============================================================

def main():
    print("""
╔══════════════════════════════════════════════════════════════╗
║         阿里司法拍卖 - 昆明市法拍房数据爬取工具              ║
╚══════════════════════════════════════════════════════════════╝

📌 准备工作：
   1. 用 Chrome 打开 https://sf.taobao.com/ 并登录淘宝账号
   2. F12 → Network → 刷新 → 点任意请求 → 复制 Cookie 值
""")

    # 检查依赖
    if not HAS_SELENIUM:
        print("⚠️  未检测到 selenium，将以 Requests 模式运行（部分字段可能获取不全）")
        print("   推荐安装: pip install selenium webdriver-manager\n")

    # Cookie
    cookie = input("请粘贴 Cookie（直接回车尝试无 Cookie 访问）: ").strip()

    # 日期范围
    print("\n📅 日期范围（直接回车使用默认值）：")
    date_start = input("  开始日期 (默认 2024-01-01): ").strip() or "2024-01-01"
    date_end = input("  结束日期 (默认 2026-03-31): ").strip() or "2026-03-31"

    try:
        datetime.strptime(date_start, "%Y-%m-%d")
        datetime.strptime(date_end, "%Y-%m-%d")
    except ValueError:
        logger.error("日期格式错误，请使用 YYYY-MM-DD")
        sys.exit(1)

    # 输出文件
    output_file = input("\n📁 输出文件名（回车使用默认）: ").strip() or None

    # 模式选择
    use_selenium = HAS_SELENIUM
    if HAS_SELENIUM:
        mode = input("\n🔧 选择模式 (1=Selenium推荐, 2=Requests轻量, 默认1): ").strip()
        use_selenium = mode != "2"

    # 开始
    print("\n" + "=" * 60)
    print("开始爬取，请耐心等待...")
    print("=" * 60 + "\n")

    scraper = SfTaobaoScraper(cookie, use_selenium=use_selenium)

    try:
        scraper.start()
        scraper.scrape_all(date_start=date_start, date_end=date_end)
        scraper.save_to_excel(output_file)
        scraper.save_failed_urls()
    except KeyboardInterrupt:
        logger.info("\n用户中断，保存已爬取数据...")
        scraper.save_to_excel(output_file)
        scraper.save_failed_urls()
    except Exception as e:
        logger.error(f"错误: {e}")
        import traceback
        traceback.print_exc()
        if scraper.results:
            scraper.save_to_excel(output_file)
    finally:
        scraper.stop()

    print("\n✅ 运行结束")


if __name__ == "__main__":
    main()
