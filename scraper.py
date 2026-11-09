"""Books to Scrape の一覧・詳細取得と差分エクスポート。"""

import argparse
import csv
import json
import logging
import math
import os
import re
import time
import uuid
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import urljoin, urlsplit
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


FIELDS = ["url", "title", "price_gbp", "stock", "rating", "category"]
EXPORT_FIELDS = ["change", *FIELDS]
RATINGS = {name: n for n, name in enumerate(
    ["One", "Two", "Three", "Four", "Five"], 1)}


def load_config(path):
    config = json.loads(Path(path).read_text(encoding="utf-8"))
    if config["base_url"] != "https://books.toscrape.com/":
        raise ValueError("base_url は https://books.toscrape.com/ に限定しています")
    for key in ("request_interval_seconds", "timeout_seconds", "backoff_seconds"):
        value = config[key]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
            raise ValueError(f"{key} は有限の正数にしてください")
    if config["request_interval_seconds"] < 1:
        raise ValueError("リクエスト間隔は1秒以上必要です")
    if type(config["max_retries"]) is not int or not 0 <= config["max_retries"] <= 10:
        raise ValueError("max_retries は0〜10の整数にしてください")
    if not config["user_agent"].strip():
        raise ValueError("user_agent は空にできません")
    root = Path(__file__).resolve().parent
    output = (root / config["output_dir"]).resolve()
    if not output.is_relative_to(root) or output == root:
        raise ValueError("output_dir はプロジェクト内のサブフォルダにしてください")
    config["output_dir"] = output
    return config


class Client:
    def __init__(self, config, session=None, clock=time.monotonic, sleep=time.sleep):
        self.config = config
        self.session = session or requests.Session()
        self.session.headers.update({"User-Agent": config["user_agent"]})
        self.clock, self.sleep = clock, sleep
        self.interval = config["request_interval_seconds"]
        self.last_finished = None
        self.robots = None

    def _request(self, url):
        """リダイレクトを含む自動の追加リクエストは行わない。"""
        for attempt in range(self.config["max_retries"] + 1):
            if self.last_finished is not None:
                self.sleep(max(0, self.interval - (self.clock() - self.last_finished)))
            response = None
            try:
                response = self.session.get(url, timeout=self.config["timeout_seconds"], allow_redirects=False)
            except requests.RequestException:
                if attempt == self.config["max_retries"]:
                    raise
            finally:
                self.last_finished = self.clock()
            if response is not None and response.status_code not in (408, 429) and response.status_code < 500:
                return response
            if attempt == self.config["max_retries"]:
                raise RuntimeError(f"再試行上限: {url}")
            delay = self.config["backoff_seconds"] * (2 ** attempt)
            if response is not None and response.headers.get("Retry-After"):
                value = response.headers["Retry-After"]
                try:
                    delay = max(delay, float(value))
                except ValueError:
                    try:
                        delay = max(delay, (parsedate_to_datetime(value) - datetime.now(timezone.utc)).total_seconds())
                    except (ValueError, TypeError, OverflowError):
                        pass
            logging.warning("再試行 %s/%s: %s", attempt + 1, self.config["max_retries"], url)
            self.sleep(delay)
        raise RuntimeError("取得失敗")

    def prepare_robots(self):
        url = urljoin(self.config["base_url"], "robots.txt")
        response = self._request(url)
        rules = RobotFileParser()
        if response.status_code in (404, 410):
            logging.info("robots.txt は HTTP %s（制限なしとして扱います）", response.status_code)
            rules.parse([])
        elif response.status_code == 200:
            if "<html" in response.text.lower():
                raise RuntimeError("robots.txt がHTMLのため安全に判定できません")
            rules.parse(response.text.splitlines())
        else:
            raise RuntimeError(f"robots.txt を確認できません: HTTP {response.status_code}")
        agent = self.config["user_agent"]
        delay = rules.crawl_delay(agent)
        rate = rules.request_rate(agent)
        self.interval = max(self.interval, delay or 0, rate.seconds / rate.requests if rate else 0)
        self.robots = rules

    def get(self, url):
        if self.robots is None:
            raise RuntimeError("robots.txt の確認が必要です")
        base, target = urlsplit(self.config["base_url"]), urlsplit(url)
        if (target.scheme, target.netloc) != (base.scheme, base.netloc):
            raise RuntimeError(f"取得対象外のURL: {url}")
        if not self.robots.can_fetch(self.config["user_agent"], url):
            raise RuntimeError(f"robots.txt により取得禁止: {url}")
        response = self._request(url)
        if response.status_code != 200:
            raise RuntimeError(f"HTTP {response.status_code}: {url}")
        response.encoding = "utf-8"
        return response.text


def parse_listing(html, url):
    soup = BeautifulSoup(html, "html.parser")
    links = [urljoin(url, node["href"]) for node in soup.select("article.product_pod h3 a[href]")]
    if not links:
        raise ValueError(f"一覧の書籍リンクがありません: {url}")
    next_link = soup.select_one("li.next a[href]")
    return links, urljoin(url, next_link["href"]) if next_link else None


def parse_book(html, url):
    soup = BeautifulSoup(html, "html.parser")

    def required(selector):
        node = soup.select_one(selector)
        if node is None or not node.get_text(strip=True):
            raise ValueError(f"詳細の必須項目がありません: {selector}: {url}")
        return " ".join(node.get_text(" ", strip=True).split())

    price = required(".product_main .price_color")
    if not re.fullmatch(r"£\d+\.\d{2}", price):
        raise ValueError(f"価格の形式が不正: {url}")
    rating_node = soup.select_one(".product_main .star-rating")
    ratings = [] if rating_node is None else [RATINGS[c] for c in rating_node.get("class", []) if c in RATINGS]
    if len(ratings) != 1:
        raise ValueError(f"評価が不正: {url}")
    return dict(zip(FIELDS, [url, required(".product_main h1"), price[1:],
                            required(".product_main .availability"), ratings[0],
                            required("ul.breadcrumb li:nth-last-child(2) a")]))


def crawl(client):
    client.prepare_robots()
    page = client.config["base_url"]
    pages, books = set(), {}
    expected = None
    while page:
        if page in pages:
            raise ValueError("ページ送りの循環を検出しました")
        pages.add(page)
        html = client.get(page)
        total = BeautifulSoup(html, "html.parser").select_one("form.form-horizontal strong")
        if total is None or not total.get_text(strip=True).isdigit():
            raise ValueError("一覧の総件数を確認できません")
        page_total = int(total.get_text(strip=True))
        if expected is None:
            expected = page_total
        elif page_total != expected:
            raise ValueError("取得中に一覧の総件数が変わりました。再実行してください")
        links, next_page = parse_listing(html, page)
        for url in links:
            if url not in books:
                books[url] = parse_book(client.get(url), url)
        logging.info("一覧 %s ページ、書籍 %s 件を取得", len(pages), len(books))
        page = next_page
    if len(books) != expected:
        raise ValueError(f"一覧の総件数と取得数が一致しません: {expected} / {len(books)}")
    return [books[url] for url in sorted(books)]


def diff(previous, current):
    old, new = ({row["url"]: row for row in rows} for rows in (previous, current))
    changes = []
    for url in sorted(old.keys() | new.keys()):
        if url not in old:
            changes.append({"change": "added", **new[url]})
        elif url not in new:
            changes.append({"change": "removed", **old[url]})
        elif old[url] != new[url]:
            changes.append({"change": "updated", **new[url]})
    return changes


def safe_csv(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def export(changes, folder):
    with (folder / "changes.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=EXPORT_FIELDS)
        writer.writeheader()
        writer.writerows({key: safe_csv(row[key]) for key in EXPORT_FIELDS} for row in changes)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Changes"
    sheet.append(EXPORT_FIELDS)
    for row in changes:
        sheet.append([row[key] for key in EXPORT_FIELDS])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for col, width in zip("ABCDEFG", [14, 60, 55, 16, 28, 12, 24]):
        sheet.column_dimensions[col].width = width
    workbook.save(folder / "changes.xlsx")


def save_run(books, output):
    output.mkdir(parents=True, exist_ok=True)
    latest = output / "latest.json"
    previous = json.loads(latest.read_text(encoding="utf-8"))["books"] if latest.exists() else []
    changes = diff(previous, books)
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ") + "-" + uuid.uuid4().hex[:8]
    folder = output / "runs" / run_id
    folder.mkdir(parents=True)
    try:
        export(changes, folder)
        state = {"run_id": run_id, "books": books}
        temporary = output / f".{run_id}.tmp"
        temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        # この置換が成功した時点で、今回の出力と比較基準が確定する。
        os.replace(temporary, latest)
    except Exception:
        logging.error("保存未完了。latest.json が指す実行のみを成功として扱ってください")
        raise
    return folder, len(changes)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default=str(Path(__file__).with_name("config.json")))
    parser.add_argument("--check-robots", action="store_true", help="robots.txt のみ確認して終了")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    try:
        config = load_config(args.config)
        output = config["output_dir"]
        output.mkdir(parents=True, exist_ok=True)
        lock = output / ".lock"
        try:
            handle = lock.open("x")
        except FileExistsError:
            raise RuntimeError("実行ロックがあります。別の実行がないことを確認してください") from None
        try:
            with handle, requests.Session() as session:
                client = Client(config, session=session)
                if args.check_robots:
                    client.prepare_robots()
                    logging.info("robots確認済み: 一覧取得可=%s、間隔=%s秒", client.robots.can_fetch(config["user_agent"], config["base_url"]), client.interval)
                else:
                    books = crawl(client)
                    folder, count = save_run(books, output)
                    logging.info("完了: 全%s件、差分%s件、出力=%s", len(books), count, folder.relative_to(Path(__file__).resolve().parent))
        finally:
            lock.unlink()
    except (OSError, ValueError, KeyError, RuntimeError, requests.RequestException) as error:
        logging.error("処理を中止しました: %s", error)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
