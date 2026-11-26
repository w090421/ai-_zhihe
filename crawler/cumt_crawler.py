#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国矿业大学公开信息爬虫（项目角色 B：学校信息库 / 最小爬虫）

数据源（均为公开页面，无需登录）：
  - 中国矿业大学新闻网  https://news.cumt.edu.cn/
  - 计算机科学与技术学院/人工智能学院  https://cs.cumt.edu.cn/

产出：
  - data/excel/矿大信息库.xlsx   单工作簿多 Sheet
  - data/seed/school-info.csv    给校园综合互助平台信息库使用
  - data/seed/teacher-info.csv   教师简表
  - data/raw/crawl_result.json   本次抓取的原始结构化结果（可复现）
  - data/crawl_stats.json        统计与失败记录
  - crawler/crawl.log            运行日志

设计约束：
  - 只抓公开页面，不登录、不绕过反爬、不请求非公开接口。
  - 每个请求间隔 >= delay 秒；同 URL 最多 3 次尝试。
  - 按 URL 去重，重复跑会全量刷新并去重，不累加重复数据。
  - 仅保留公开的专业信息；教师信息只采集官网已公开的姓名/院系/职称/研究方向等。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import posixpath
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse, urlunparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
EXCEL_DIR = DATA_DIR / "excel"
SEED_DIR = DATA_DIR / "seed"
RAW_DIR = DATA_DIR / "raw"
CACHE_DIR = DATA_DIR / "cache"
LOG_DIR = ROOT / "crawler"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

ALLOWED_HOSTS = {"cs.cumt.edu.cn", "news.cumt.edu.cn"}

SECTIONS: list[dict[str, str]] = [
    # 学校新闻网
    {"site": "学校", "name": "要闻矿大", "category": "学校新闻", "url": "https://news.cumt.edu.cn/ywk.htm"},
    {"site": "学校", "name": "资讯矿大", "category": "学校新闻", "url": "https://news.cumt.edu.cn/sdxw.htm"},
    {"site": "学校", "name": "校园快讯", "category": "校园新闻", "url": "https://news.cumt.edu.cn/xykx.htm"},
    {"site": "学校", "name": "学术聚焦", "category": "学术信息", "url": "https://news.cumt.edu.cn/xsjj.htm"},
    # 学院官网
    {"site": "学院", "name": "新闻动态", "category": "学院新闻", "url": "https://cs.cumt.edu.cn/xwdt.htm"},
    {"site": "学院", "name": "通知公告", "category": "通知公告", "url": "https://cs.cumt.edu.cn/tzgg.htm"},
    {"site": "学院", "name": "学术通知", "category": "学术信息", "url": "https://cs.cumt.edu.cn/xsky/xstz.htm"},
]

TEACHER_PAGES: list[dict[str, str]] = [
    {"url": "https://cs.cumt.edu.cn/szdw/xyj1/jsjkxyjsx1.htm"},
    {"url": "https://cs.cumt.edu.cn/szdw/xyj1/wlkjaqx.htm"},
    {"url": "https://cs.cumt.edu.cn/szdw/xyj1/sjkxygcx.htm"},
    {"url": "https://cs.cumt.edu.cn/szdw/xyj1/rgznx_znkxyjsx.htm"},
]

TITLE_RANK = {"教授": 1, "副教授": 2, "准聘副教授": 2, "准聘副高": 2, "讲师": 3, "实验师": 4, "助理实验师": 4, "师资博士后": 5, "行健博士后": 5}

logger = logging.getLogger("cumt")


class Crawler:
    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.delay = args.delay
        self.session = self._make_session()
        self.stats: dict[str, Any] = {
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "finished_at": None,
            "elapsed_seconds": None,
            "sections": {},
            "teachers": {},
            "requests": 0,
            "cache_hits": 0,
            "failed_urls": [],
            "duplicates_merged": 0,
        }
        self.articles: dict[str, dict[str, Any]] = {}
        self.teachers: dict[str, dict[str, Any]] = {}

    # ---------------------------------------------------------------- HTTP
    @staticmethod
    def _make_session() -> requests.Session:
        session = requests.Session()
        retry = Retry(
            total=2,  # 初次 + 2 次重试 = 最多 3 次
            backoff_factor=1.2,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.6",
                "Connection": "keep-alive",
            }
        )
        return session

    def _cache_path(self, url: str) -> Path:
        key = hashlib.sha1(url.encode("utf-8")).hexdigest()
        return CACHE_DIR / f"{key}.html"

    def fetch(self, url: str, referer: str | None = None) -> str | None:
        cache_file = self._cache_path(url)
        if not self.args.no_cache and cache_file.exists():
            try:
                text = cache_file.read_text(encoding="utf-8")
                if text.strip():
                    self.stats["cache_hits"] += 1
                    return text
            except OSError:
                pass
        self.stats["requests"] += 1
        if self.delay > 0:
            time.sleep(self.delay)
        headers = {"Referer": referer} if referer else None
        try:
            resp = self.session.get(url, headers=headers, timeout=25)
        except requests.RequestException as exc:
            logger.warning("请求失败 %s -> %s", url, exc)
            self.stats["failed_urls"].append({"url": url, "error": type(exc).__name__})
            return None
        if resp.status_code != 200:
            logger.warning("HTTP %s %s", resp.status_code, url)
            self.stats["failed_urls"].append({"url": url, "error": f"HTTP {resp.status_code}"})
            return None
        if not resp.encoding or resp.encoding.lower() in ("iso-8859-1", "ascii"):
            resp.encoding = resp.apparent_encoding or "utf-8"
        text = resp.text
        if not self.args.no_cache:
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            cache_file.write_text(text, encoding="utf-8")
        return text

    # ----------------------------------------------------------- pagination
    @staticmethod
    def _first_page_parts(first_url: str) -> tuple[str, str, str]:
        parsed = urlparse(first_url)
        dirpath, name = posixpath.split(parsed.path)
        dirpath = dirpath.rstrip("/")
        stem = name[:-4] if name.endswith(".htm") else name
        return parsed.scheme + "://" + parsed.netloc, dirpath, stem

    @classmethod
    def total_pages(cls, first_url: str, soup: BeautifulSoup) -> int:
        origin, dirpath, stem = cls._first_page_parts(first_url)
        pattern = re.compile(rf"^{re.escape(dirpath)}/{re.escape(stem)}/(\d+)\.htm$")
        nums: list[int] = []
        for a in soup.find_all("a", href=True):
            path = urlparse(urljoin(first_url, a["href"])).path
            m = pattern.match(path)
            if m:
                nums.append(int(m.group(1)))
        return max(nums) + 1 if nums else 1

    @classmethod
    def page_url(cls, first_url: str, page: int, total: int) -> str:
        if page <= 1:
            return first_url
        parsed = urlparse(first_url)
        _, dirpath, stem = cls._first_page_parts(first_url)
        new_path = f"{dirpath}/{stem}/{total - page + 1}.htm" if dirpath else f"/{stem}/{total - page + 1}.htm"
        return urlunparse(parsed._replace(path=new_path, query="", fragment=""))

    # ------------------------------------------------------------- parsing
    def crawl_section(self, section: dict[str, str]) -> None:
        first_url = section["url"]
        html = self.fetch(first_url)
        if not html:
            self.stats["sections"][section["name"]] = {"list_pages": 0, "items": 0, "error": "first_page_failed"}
            return
        soup = BeautifulSoup(html, "html.parser")
        raw_total = self.total_pages(first_url, soup)
        total = min(raw_total, self.args.max_pages)
        page_items = 0
        logger.info("[%s] %s 共 %s 页待抓（上限 %s）", section["site"], section["name"], total, self.args.max_pages)

        for page in range(1, total + 1):
            url = first_url
            if page > 1:
                url = self.page_url(first_url, page, raw_total)
                html = self.fetch(url, referer=first_url)
                if not html:
                    continue
                soup = BeautifulSoup(html, "html.parser")
            rows = self.parse_cs_list(soup, url) if section["site"] == "学院" else self.parse_news_list(soup, url)
            if self.args.per_section:
                rows = rows[: self.args.per_section]
            for row in rows:
                self.add_article(row, section)
                page_items += 1
            logger.info("  - 第 %s/%s 页：%s 条", page, total, len(rows))
            if self.args.per_section and page_items >= self.args.per_section:
                break

        self.stats["sections"][section["name"]] = {"list_pages": total, "items": page_items}

    @staticmethod
    def _detail_url(page_url: str, href: str) -> str:
        return urljoin(page_url, href.strip()).split("#")[0]

    @staticmethod
    def _parse_cs_date(date_el: Any) -> str:
        if not date_el:
            return ""
        raw = date_el.get_text(" ", strip=True)
        full = re.search(r"(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})", raw)
        if full:
            return f"{full.group(1)}-{int(full.group(2)):02d}-{int(full.group(3)):02d}"
        span = date_el.find("span")
        ym = re.search(r"(\d{4})\s*-\s*(\d{1,2})", raw)
        if span and ym and re.fullmatch(r"\d{1,2}", span.get_text(strip=True)):
            return f"{ym.group(1)}-{int(ym.group(2)):02d}-{int(span.get_text(strip=True)):02d}"
        return ""

    def parse_cs_list(self, soup: BeautifulSoup, page_url: str) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        for a in soup.select("div.img-list a.item[href], div.text-list a.item[href]"):
            title = (a.get("title") or "").strip()
            if not title:
                el = a.select_one(".tit")
                title = clean_inline(el.get_text(" ", strip=True)) if el else ""
            if not title:
                continue
            des = a.select_one(".des")
            summary = clean_inline(des.get_text(" ", strip=True)) if des else ""
            rows.append(
                {
                    "title": title,
                    "date": self._parse_cs_date(a.select_one(".date")),
                    "summary": summary,
                    "url": self._detail_url(page_url, a.get("href", "")),
                }
            )
        return rows

    @staticmethod
    def parse_news_list(soup: BeautifulSoup, page_url: str) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        for li in soup.select("ul.ky-list li"):
            a = li.select_one(".ky-list-tt a[href]")
            if not a:
                continue
            h3 = li.select_one(".ky-list-tt h3")
            p = li.select_one(".ky-list-tt p")
            title = clean_inline(h3.get_text(" ", strip=True)) if h3 else ""
            if not title:
                continue
            date = ""
            d = li.select_one(".ky-list-date")
            if d:
                y, m, day = d.find("em"), d.find("span"), d.find("font")
                if y and m and day:
                    date = f"{y.get_text(strip=True)}-{int(m.get_text(strip=True)):02d}-{int(day.get_text(strip=True)):02d}"
            summary = clean_inline(p.get_text(" ", strip=True)) if p else ""
            rows.append(
                {
                    "title": title,
                    "date": date,
                    "summary": summary,
                    "url": urljoin(page_url, a["href"]).split("#")[0],
                }
            )
        return rows

    def add_article(self, row: dict[str, str], section: dict[str, str]) -> None:
        url = row["url"]
        if not url:
            return
        if url in self.articles:
            self.stats["duplicates_merged"] += 1
            existing = self.articles[url]
            if section["name"] not in existing["栏目"]:
                existing["栏目"] += "、" + section["name"]
            return
        self.articles[url] = {
            "site": section["site"],
            "栏目": section["name"],
            "分类": section["category"],
            "标题": row["title"],
            "发布日期": row["date"],
            "摘要": row["summary"],
            "正文": "",
            "作者": "",
            "来源": "",
            "链接": url,
            "抓取时间": "",
        }

    # ------------------------------------------------------------- details
    def enrich_articles(self) -> None:
        total = len(self.articles)
        for idx, (url, rec) in enumerate(self.articles.items(), start=1):
            if self.args.no_details:
                break
            host = urlparse(url).netloc
            if host not in ALLOWED_HOSTS:
                logger.info("[%s/%s] 跳过外部详情：%s", idx, total, url)
                rec["来源"] = host
                if not rec["摘要"]:
                    rec["摘要"] = "详见原文链接"
                continue
            html = self.fetch(url)
            if not html:
                continue
            soup = BeautifulSoup(html, "html.parser")
            raw_text = soup.get_text(" ", strip=True)
            title = detail_title(soup)
            if title and len(title) >= len(rec["标题"]):
                rec["标题"] = title
            date = detail_date(raw_text)
            if date:
                rec["发布日期"] = date
            author, source = detail_meta(raw_text)
            if host == "cs.cumt.edu.cn":
                # 学院页面的“新闻来源”大多是撰稿人，来源统一记为学院官网
                if source and not is_org_name(source):
                    author = author or source
                source = "计算机科学与技术学院/人工智能学院"
            rec["作者"] = author
            rec["来源"] = source or host
            body = extract_content(soup)
            rec["正文"] = body
            if (not rec["摘要"]) or rec["摘要"].endswith(("...", "…")):
                rec["摘要"] = make_summary(body, 180)
            if not rec["摘要"]:
                rec["摘要"] = "详见原文链接"
            if idx % 20 == 0 or idx == total:
                logger.info("详情进度 %s/%s", idx, total)

    # ------------------------------------------------------------ teachers
    def crawl_teachers(self) -> None:
        for page in TEACHER_PAGES:
            html = self.fetch(page["url"])
            if not html:
                continue
            soup = BeautifulSoup(html, "html.parser")
            dept = (soup.title.get_text(strip=True) if soup.title else "").split("-")[0].strip() or "计算机科学与技术学院"
            found = 0
            for content in soup.select("div.tea-list div.content"):
                group_el = content.select_one("div.title")
                group = clean_inline(group_el.get_text(" ", strip=True)) if group_el else ""
                for a in content.select("a.item[href]"):
                    name = (a.get("title") or "").strip()
                    if not name:
                        tit = a.select_one(".tit")
                        name = clean_inline(tit.get_text(" ", strip=True)) if tit else ""
                    if not name:
                        continue
                    url = urljoin(page["url"], a["href"]).split("#")[0]
                    self.add_teacher(name, dept, group, url)
                    found += 1
            self.stats["teachers"][dept] = found
            logger.info("教师列表 %s：%s 人", dept, found)

        if self.args.teacher_limit:
            for key in list(self.teachers)[self.args.teacher_limit :]:
                self.teachers.pop(key, None)

        total = len(self.teachers)
        for idx, (url, rec) in enumerate(self.teachers.items(), start=1):
            html = self.fetch(url)
            if not html:
                continue
            soup = BeautifulSoup(html, "html.parser")
            raw_text = soup.get_text(" ", strip=True)
            bio = extract_content(soup)
            rec["简介"] = bio[:900]
            rec["发布日期"] = detail_date(raw_text)
            rec["研究方向"] = extract_research(bio)
            rec["毕业院校"] = extract_graduated(bio)
            rec["邮箱"] = extract_email(bio or raw_text)
            rec["导师类别"] = extract_advisor(bio)
            if idx % 20 == 0 or idx == total:
                logger.info("教师详情进度 %s/%s", idx, total)

    def add_teacher(self, raw_name: str, dept: str, group: str, url: str) -> None:
        name, suffix = split_name_suffix(raw_name)
        title = suffix or group
        if url in self.teachers:
            old = self.teachers[url]
            if dept not in old["所属院系"]:
                old["所属院系"] += "、" + dept
            return
        self.teachers[url] = {
            "姓名": name,
            "所属院系": dept,
            "职称": title,
            "导师类别": "",
            "研究方向": "",
            "毕业院校": "",
            "邮箱": "",
            "简介": "",
            "发布日期": "",
            "详情链接": url,
            "来源": "计算机科学与技术学院/人工智能学院官网",
            "抓取时间": "",
            "_rank": TITLE_RANK.get(title, 9),
        }


# ---------------------------------------------------------------- helpers
def clean_meta(value: str) -> str:
    value = value.replace("\u200b", "").replace("\xa0", " ").replace("\u3000", " ")
    return mask_personal(re.sub(r"\s+", " ", value).strip())


_PHONE_RE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
_LANDLINE_RE = re.compile(r"(?<!\d)0\d{2,3}-?\d{7,8}(?!\d)")
_ID_RE = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")


def mask_personal(text: str) -> str:
    """脱敏：手机号、座机、身份证号一律替换为 [已脱敏]。"""
    if not text:
        return text
    text = _ID_RE.sub("[已脱敏]", text)
    text = _PHONE_RE.sub("[已脱敏]", text)
    text = _LANDLINE_RE.sub("[已脱敏]", text)
    return text


def clean_inline(value: str) -> str:
    value = value.replace("\u200b", "").replace("\xa0", " ").replace("\u3000", " ")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", value)
    value = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[，。；：、！？）》」』])", "", value)
    value = re.sub(r"(?<=[，。；：、！？）》」』])\s+(?=[\u4e00-\u9fff\dA-Za-z])", "", value)
    value = re.sub(r"(?<=[（《「『])\s+", "", value)
    value = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=\d)", "", value)
    value = re.sub(r"(?<=[\u4e00-\u9fff])\s*[—–-]\s*(?=[\u4e00-\u9fff])", "-", value)
    value = re.sub(r"(?<=\d)\s+(?=\d)", "", value)
    value = re.sub(r"\b([A-Z])\s+([A-Z]{2,})\b", r"\1\2", value)
    return mask_personal(value.strip())


def extract_content(soup: BeautifulSoup) -> str:
    for selector in ("#vsb_content_1001", "#vsb_content_2", "#vsb_content", "div.v_news_content", "div.news_bd"):
        el = soup.select_one(selector)
        if not el:
            continue
        paragraphs = [clean_inline(p.get_text(" ", strip=True)) for p in el.find_all("p")]
        paragraphs = [p for p in paragraphs if p and not re.match(r"^(上一篇|下一篇)[：:]?$", p)]
        text = "\n".join(paragraphs) if paragraphs else clean_inline(el.get_text(" ", strip=True))
        if len(text) >= 20:
            return text
    return ""


def detail_title(soup: BeautifulSoup) -> str:
    for selector in ("div.con_box h3", "h1", "h3"):
        el = soup.select_one(selector)
        if el:
            title = clean_inline(el.get_text(" ", strip=True))
            if title:
                return title
    if soup.title:
        return clean_inline(soup.title.get_text(strip=True).split("-")[0])
    return ""


def detail_date(raw_text: str) -> str:
    for label in ("发布时间", "发布日期", "日期", "时间"):
        m = re.search(rf"{label}[：:]\s*(\d{{4}})[-年/.](\d{{1,2}})[-月/.](\d{{1,2}})", raw_text)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def detail_meta(raw_text: str) -> tuple[str, str]:
    author = ""
    source = ""
    m = re.search(r"作者[：:]\s*(.+?)\s+来源[：:]", raw_text)
    if m:
        author = clean_meta(m.group(1))
    m = re.search(r"来源[：:]\s*(.+?)\s+(?:摄影|浏览数|发布时间|编辑|责任编辑)", raw_text)
    if m:
        source = clean_meta(m.group(1))
    if not source:
        m = re.search(r"新闻来源[：:]\s*(.+?)\s+(?:责任编辑|一审|二审|三审|下一篇|摄影|$)", raw_text)
        if m:
            source = clean_meta(m.group(1))
    return author, source


def is_org_name(value: str) -> bool:
    return any(k in value for k in ("学院", "大学", "中心", "研究院", "部", "处", "办公室", "实验室"))


def make_summary(text: str, limit: int) -> str:
    paragraphs = [re.sub(r"\s+", " ", p).strip() for p in text.splitlines() if p.strip()]
    if not paragraphs:
        return ""
    start = 0
    for i, para in enumerate(paragraphs):
        if len(para) >= 30 or "。" in para:
            start = i
            break
    text = " ".join(paragraphs[start:]).strip()
    if not text:
        return ""
    return text[:limit].rstrip() + ("…" if len(text) > limit else "")


def split_name_suffix(raw_name: str) -> tuple[str, str]:
    m = re.search(r"[（(]([^）)]+)[）)]", raw_name)
    if not m:
        return raw_name.strip(), ""
    suffix = clean_inline(m.group(1))
    name = clean_inline(raw_name[: m.start()])
    return name or raw_name.strip(), suffix


def extract_research(bio: str) -> str:
    text = re.sub(r"\s+", "", bio)
    patterns = [
        r"(?:主要从事|从事|长期致力于|致力于|主要研究|研究方向(?:为|是|：|:)?|研究领域(?:为|是|：|:)?|研究兴趣(?:为|是|：|:)?)([^。；;]{4,120})",
        r"(?:专长|擅长)(?:为|是|：|:)?([^。；;]{4,120})",
        r"[Rr]esearch interests?(?:include|are|focus on)?[：:\s]+([^.;]{4,160})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            value = m.group(1).strip("，。；; 、")
            value = re.split(r"(?:等方面|等方向|等理论|等应用|的研究工作|研究工作|的理论与应用研究)", value)[0]
            value = value.strip("，。；; 、")
            if value:
                return value
    return ""


def extract_graduated(bio: str) -> str:
    m = re.search(r"毕业于([^，。；;\n]{2,40})", bio)
    return clean_inline(m.group(1)) if m else ""


def extract_email(text: str) -> str:
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text or "")
    return m.group(0) if m else ""


def extract_advisor(bio: str) -> str:
    has_phd = any(k in bio for k in ("博士生导师", "博导"))
    has_master = any(k in bio for k in ("硕士生导师", "硕导"))
    if has_phd and has_master:
        return "博士生导师/硕士生导师"
    if has_phd:
        return "博士生导师"
    if has_master:
        return "硕士生导师"
    return ""


# ---------------------------------------------------------------- writers
def safe_text(value: Any) -> str:
    text = "" if value is None else str(value)
    if text[:1] in ("=", "+", "-", "@"):
        return "'" + text
    return text


def write_excel(path: Path, articles: list[dict[str, Any]], teachers: list[dict[str, Any]], stats: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_sheet(ws, headers: list[str], widths: list[int]) -> None:
        for c, (h, w) in enumerate(zip(headers, widths), start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
        ws.row_dimensions[1].height = 24

    # Sheet 1: 学校新闻
    ws = wb.active
    ws.title = "学校新闻"
    headers = ["序号", "栏目", "分类", "标题", "发布日期", "摘要", "正文", "作者", "来源", "链接", "抓取时间"]
    widths = [6, 12, 12, 42, 12, 50, 80, 16, 24, 50, 20]
    ws.append(headers)
    school = [r for r in articles if r["site"] == "学校"]
    for i, rec in enumerate(school, start=1):
        row = [i, rec["栏目"], rec["分类"], safe_text(rec["标题"]), rec["发布日期"], safe_text(rec["摘要"]), safe_text(rec["正文"]), safe_text(rec["作者"]), safe_text(rec["来源"]), rec["链接"], rec["抓取时间"]]
        ws.append(row)
        ws.cell(row=ws.max_row, column=10).hyperlink = rec["链接"]
        ws.cell(row=ws.max_row, column=10).style = "Hyperlink"
    style_sheet(ws, headers, widths)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in (4, 6, 7, 9, 10))

    # Sheet 2: 学院动态
    ws = wb.create_sheet("学院动态")
    headers2 = ["序号", "栏目", "分类", "标题", "发布日期", "摘要", "正文", "来源", "链接", "抓取时间"]
    ws.append(headers2)
    college = [r for r in articles if r["site"] == "学院"]
    for i, rec in enumerate(college, start=1):
        row = [i, rec["栏目"], rec["分类"], safe_text(rec["标题"]), rec["发布日期"], safe_text(rec["摘要"]), safe_text(rec["正文"]), safe_text(rec["来源"]), rec["链接"], rec["抓取时间"]]
        ws.append(row)
        ws.cell(row=ws.max_row, column=9).hyperlink = rec["链接"]
        ws.cell(row=ws.max_row, column=9).style = "Hyperlink"
    style_sheet(ws, headers2, [6, 12, 12, 42, 12, 50, 80, 24, 50, 20])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in (4, 6, 7, 8, 9))

    # Sheet 3: 教师信息
    ws = wb.create_sheet("教师信息")
    headers3 = ["序号", "姓名", "所属院系", "职称", "导师类别", "研究方向", "毕业院校", "邮箱", "简介", "发布日期", "详情链接", "来源", "抓取时间"]
    ws.append(headers3)
    for i, rec in enumerate(teachers, start=1):
        row = [i, safe_text(rec["姓名"]), rec["所属院系"], rec["职称"], rec["导师类别"], safe_text(rec["研究方向"]), safe_text(rec["毕业院校"]), safe_text(rec["邮箱"]), safe_text(rec["简介"]), rec["发布日期"], rec["详情链接"], rec["来源"], rec["抓取时间"]]
        ws.append(row)
        ws.cell(row=ws.max_row, column=11).hyperlink = rec["详情链接"]
        ws.cell(row=ws.max_row, column=11).style = "Hyperlink"
    style_sheet(ws, headers3, [6, 14, 30, 14, 20, 40, 24, 28, 80, 12, 50, 32, 20])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in (3, 5, 6, 7, 9, 11, 12))

    # Sheet 4: 数据说明
    ws = wb.create_sheet("数据说明")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110
    info = [
        ("数据集", "中国矿业大学公开信息库（校园综合互助平台 · 角色 B 产出）"),
        ("数据源", "中国矿业大学新闻网 https://news.cumt.edu.cn/；计算机科学与技术学院/人工智能学院 https://cs.cumt.edu.cn/"),
        ("抓取时间", stats.get("started_at", "")),
        ("抓取范围", f"学校新闻/学术聚焦各近 {stats.get('max_pages')} 页；学院新闻/通知/学术通知各近 {stats.get('max_pages')} 页；学院教师列表全部"),
        ("新闻条数", str(len(articles))),
        ("教师条数", str(len(teachers))),
        ("团队契约文件", "data/seed/school-info.csv：id,title,category,summary,source_url；category 已映射到教务/生活/场馆/奖助/就业/其他；UTF-8 无 BOM。"),
        ("追溯文件", "data/seed/school-info.full.csv 保留原始栏目/来源/日期；data/seed/teacher-info.csv 为教师简表。"),
        ("去重规则", "按详情页 URL 去重；同一 URL 在多个栏目出现时合并栏目名；id 由 URL 哈希生成，重复跑保持稳定。"),
        ("合规说明", "仅抓取无需登录的公开页面；未绕过任何访问控制或反爬机制；content.jsp 等提示需 VPN 的页面（HTTP 403）未尝试绕过，仅保留标题与链接；请求间隔>=%.1f 秒；教师信息仅保留官网公开的专业资料。" % stats.get("delay", 0.8)),
        ("字段说明", "摘要为列表页摘要或正文首个实质段落前 180 字；正文为详情页正文；链接为官网原文地址；发布日期以官网显示为准。"),
        ("维护方式", "重新运行 crawler/cumt_crawler.py 即可全量刷新并按 URL 去重；本地 HTML 缓存在 data/cache/，断点续跑不重复请求。"),
    ]
    for r, (k, v) in enumerate(info, start=1):
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=v).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 32

    wb.save(path)


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str], bom: bool = True) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig" if bom else "utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


INFO_CATEGORY_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("就业", ("招聘", "就业", "双选", "岗位", "实习", "求职", "宣讲会")),
    ("奖助", ("奖学金", "助学金", "奖助", "资助", "补助", "助学贷款")),
    ("场馆", ("图书馆", "体育馆", "场馆", "报告厅", "运动场", "游泳馆")),
    ("生活", ("食堂", "宿舍", "后勤", "校医院", "医保", "体检", "生活", "快递")),
    ("教务", ("教务", "教学", "课程", "考试", "学籍", "成绩", "培养", "学位", "选课", "开学", "放假", "推免", "转专业")),
]


def map_info_category(rec: dict[str, Any]) -> str:
    # 只用标题做关键词映射，避免正文里“教务部/就业中心”等词造成误判
    title = rec["标题"]
    for category, keywords in INFO_CATEGORY_KEYWORDS:
        if any(k in title for k in keywords):
            return category
    if rec["分类"] in ("通知公告", "学术信息"):
        return "教务"
    return "其他"


def stable_id(prefix: str, url: str) -> str:
    return f"{prefix}-{hashlib.sha1(url.encode('utf-8')).hexdigest()[:10]}"


def build_school_csv(articles: list[dict[str, Any]]) -> list[dict[str, str]]:
    """团队契约格式：id,title,category,summary,source_url（UTF-8 无 BOM）。"""
    rows: list[dict[str, str]] = []
    for rec in articles:
        rows.append(
            {
                "id": stable_id("cumt-news", rec["链接"]),
                "title": rec["标题"],
                "category": map_info_category(rec),
                "summary": rec["摘要"],
                "source_url": rec["链接"],
            }
        )
    return rows


def build_school_full_csv(articles: list[dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for rec in articles:
        rows.append(
            {
                "id": stable_id("cumt-news", rec["链接"]),
                "title": rec["标题"],
                "summary": rec["摘要"],
                "category": rec["分类"],
                "info_category": map_info_category(rec),
                "source_url": rec["链接"],
                "source_name": "中国矿业大学新闻网" if rec["site"] == "学校" else "计算机科学与技术学院/人工智能学院",
                "published_at": rec["发布日期"],
                "tags": rec["栏目"],
            }
        )
    return rows


def build_teacher_csv(teachers: list[dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for rec in teachers:
        rows.append(
            {
                "id": stable_id("cumt-teacher", rec["详情链接"]),
                "name": rec["姓名"],
                "department": rec["所属院系"],
                "title": rec["职称"],
                "advisor_type": rec["导师类别"],
                "research_area": rec["研究方向"],
                "graduated_from": rec["毕业院校"],
                "email": rec["邮箱"],
                "bio": rec["简介"],
                "source_url": rec["详情链接"],
                "source_name": rec["来源"],
                "published_at": rec["发布日期"],
                "crawled_at": rec["抓取时间"],
            }
        )
    return rows


# ------------------------------------------------------------------ main
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="中国矿业大学公开信息爬虫（角色 B）")
    parser.add_argument("--max-pages", type=int, default=5, help="每个栏目最多抓取的列表页数（默认 5）")
    parser.add_argument("--delay", type=float, default=0.8, help="请求间隔秒数（默认 0.8）")
    parser.add_argument("--per-section", type=int, default=0, help="每个栏目最多条数，0 表示不限制（调试用）")
    parser.add_argument("--teacher-limit", type=int, default=0, help="教师最多抓取人数，0 表示全部（调试用）")
    parser.add_argument("--no-details", action="store_true", help="只抓列表，不抓详情")
    parser.add_argument("--no-teachers", action="store_true", help="跳过教师抓取")
    parser.add_argument("--no-cache", action="store_true", help="禁用本地 HTML 缓存（默认启用，便于断点续跑）")
    parser.add_argument("--smoke", action="store_true", help="冒烟模式：1 页、每栏目 2 条、教师 3 人")
    return parser.parse_args()


def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(LOG_DIR / "crawl.log", encoding="utf-8", mode="w")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    args = parse_args()
    if args.smoke:
        args.max_pages = 1
        args.per_section = 2
        args.teacher_limit = 3
        args.delay = min(args.delay, 0.2)
    setup_logging()

    started = time.time()
    crawler = Crawler(args)
    crawler.stats["max_pages"] = args.max_pages
    crawler.stats["delay"] = args.delay
    logger.info("开始抓取：max_pages=%s delay=%.1fs", args.max_pages, args.delay)

    for section in SECTIONS:
        crawler.crawl_section(section)

    crawler.enrich_articles()
    if not args.no_teachers:
        crawler.crawl_teachers()

    now = datetime.now().isoformat(timespec="seconds")
    for rec in crawler.articles.values():
        rec["抓取时间"] = now
    for rec in crawler.teachers.values():
        rec["抓取时间"] = now

    articles = sorted(crawler.articles.values(), key=lambda r: (r["发布日期"] or "", r["site"], r["标题"]), reverse=True)
    teachers = sorted(crawler.teachers.values(), key=lambda r: (r["所属院系"], r["_rank"], r["姓名"]))
    for rec in teachers:
        rec.pop("_rank", None)

    stats = crawler.stats
    stats["finished_at"] = now
    stats["elapsed_seconds"] = round(time.time() - started, 1)
    stats["article_count"] = len(articles)
    stats["teacher_count"] = len(teachers)
    stats["article_by_site"] = {
        "学校": sum(1 for r in articles if r["site"] == "学校"),
        "学院": sum(1 for r in articles if r["site"] == "学院"),
    }
    stats["article_by_category"] = {}
    for rec in articles:
        stats["article_by_category"][rec["分类"]] = stats["article_by_category"].get(rec["分类"], 0) + 1
    stats["info_category"] = {}
    for rec in articles:
        category = map_info_category(rec)
        stats["info_category"][category] = stats["info_category"].get(category, 0) + 1

    excel_path = EXCEL_DIR / "矿大信息库.xlsx"
    write_excel(excel_path, articles, teachers, stats)
    # 团队契约格式（与 data/seed/school-info.csv 现有表头一致，UTF-8 无 BOM）
    write_csv(
        SEED_DIR / "school-info.csv",
        build_school_csv(articles),
        ["id", "title", "category", "summary", "source_url"],
        bom=False,
    )
    # 追溯用完整字段
    write_csv(
        SEED_DIR / "school-info.full.csv",
        build_school_full_csv(articles),
        ["id", "title", "summary", "category", "info_category", "source_url", "source_name", "published_at", "tags"],
        bom=True,
    )
    write_csv(
        SEED_DIR / "teacher-info.csv",
        build_teacher_csv(teachers),
        ["id", "name", "department", "title", "advisor_type", "research_area", "graduated_from", "email", "bio", "source_url", "source_name", "published_at", "crawled_at"],
        bom=True,
    )
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    (RAW_DIR / "crawl_result.json").write_text(
        json.dumps({"stats": stats, "articles": articles, "teachers": teachers}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (DATA_DIR / "crawl_stats.json").write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")

    logger.info("完成：新闻 %s 条，教师 %s 人，耗时 %.1fs", len(articles), len(teachers), stats["elapsed_seconds"])
    logger.info("Excel: %s", excel_path)
    logger.info("CSV: %s", SEED_DIR / "school-info.csv")
    if stats["failed_urls"]:
        logger.warning("失败请求 %s 个，详情见 data/crawl_stats.json", len(stats["failed_urls"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
