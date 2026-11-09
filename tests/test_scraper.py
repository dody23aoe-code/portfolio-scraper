import csv
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import requests
from openpyxl import load_workbook

from scraper import Client, crawl, diff, load_config, parse_book, parse_listing, save_run


ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).with_name("fixtures")
BASE = "https://books.toscrape.com/"
BOOK = BASE + "catalogue/sample-book_1/index.html"


def fixture(name):
    return (FIXTURES / name).read_text(encoding="utf-8")


def response(text="", status=200, headers=None):
    result = requests.Response()
    result.status_code = status
    result._content = text.encode("utf-8")
    result.encoding = "utf-8"
    result.headers.update(headers or {})
    return result


class FakeClock:
    def __init__(self):
        self.now = 0

    def time(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


class FakeSession:
    def __init__(self, results, clock):
        self.results = iter(results)
        self.clock = clock
        self.headers = {}
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append((url, self.clock.time(), kwargs))
        result = next(self.results)
        if isinstance(result, Exception):
            raise result
        return result


class ScraperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        (ROOT / ".tmp").mkdir(exist_ok=True)

    def setUp(self):
        # 不注意で実HTTPを呼んだテストも必ず失敗させる。
        blocker = patch("requests.sessions.Session.request", side_effect=AssertionError("network disabled"))
        blocker.start()
        self.addCleanup(blocker.stop)
        self.config = load_config(ROOT / "config.json")

    def client(self, results):
        clock = FakeClock()
        session = FakeSession(results, clock)
        return Client(self.config, session, clock.time, clock.sleep), session

    def test_parse_saved_html(self):
        links, next_page = parse_listing(fixture("listing.html"), BASE)
        self.assertEqual(links, [BOOK])
        self.assertEqual(next_page, BASE + "catalogue/page-2.html")
        self.assertEqual(parse_book(fixture("book.html"), BOOK), {
            "url": BOOK, "title": "Sample Book & Stories", "price_gbp": "51.77",
            "stock": "In stock (22 available)", "rating": 3, "category": "Poetry"})

    def test_bad_html_aborts(self):
        for html in ("<html></html>", fixture("book.html").replace("£51.77", "unknown"),
                     fixture("book.html").replace("Three", "Unknown")):
            with self.assertRaises(ValueError):
                parse_book(html, BOOK)
        with self.assertRaises(ValueError):
            parse_listing("<html></html>", BASE)

    def test_saved_site_html(self):
        links, next_page = parse_listing(fixture("site_listing.html"), BASE)
        self.assertEqual(len(links), 20)
        self.assertEqual(next_page, BASE + "catalogue/page-2.html")
        book = parse_book(fixture("site_book.html"), links[0])
        self.assertEqual(book["title"], "A Light in the Attic")
        self.assertEqual(book["price_gbp"], "51.77")
        self.assertEqual(book["category"], "Poetry")

    def test_incomplete_listing_aborts(self):
        html = fixture("listing.html").replace('<li class="next"><a href="catalogue/page-2.html">next</a></li>', '')
        client, _ = self.client([response(), response(html), response(fixture("book.html"))])
        with self.assertRaisesRegex(ValueError, "総件数と取得数"):
            crawl(client)

    def test_crawl_all_pages_deduplicates_and_reads_robots_first(self):
        client, session = self.client([
            response(fixture("robots.txt")), response(fixture("listing.html")),
            response(fixture("book.html")), response(fixture("listing_last.html")),
            response(fixture("book.html").replace("Poetry", "Fiction"))])
        books = crawl(client)
        self.assertEqual(len(books), 2)
        self.assertEqual({book["category"] for book in books}, {"Poetry", "Fiction"})
        self.assertEqual(session.calls[0][0], BASE + "robots.txt")
        self.assertEqual(len(session.calls), 5)
        self.assertTrue(all(b[1] - a[1] >= 3 for a, b in zip(session.calls, session.calls[1:])))
        self.assertTrue(all(not call[2]["allow_redirects"] for call in session.calls))

    def test_robots_denial_and_external_urls_never_requested(self):
        client, session = self.client([response(fixture("robots.txt"))])
        with self.assertRaises(RuntimeError):
            client.get(BASE)
        client.prepare_robots()
        for url in (BASE + "private/secret", "https://example.org/"):
            with self.assertRaises(RuntimeError):
                client.get(url)
        self.assertEqual(len(session.calls), 1)

    def test_robots_missing_allowed_but_errors_stop(self):
        for status in (404, 410):
            client, _ = self.client([response(status=status)])
            client.prepare_robots()
            self.assertTrue(client.robots.can_fetch(self.config["user_agent"], BASE))
        for status in (401, 403, 302, 500):
            client, _ = self.client([response(status=status)] * 4)
            with self.assertRaises(RuntimeError):
                client.prepare_robots()

    def test_retries_bounded_and_spaced(self):
        client, session = self.client([requests.Timeout()] * 4)
        with self.assertRaises(requests.Timeout):
            client.prepare_robots()
        self.assertEqual(len(session.calls), 4)
        self.assertTrue(all(b[1] - a[1] >= 1 for a, b in zip(session.calls, session.calls[1:])))

    def test_retry_after_and_success(self):
        client, session = self.client([response(status=429, headers={"Retry-After": "7"}), response()])
        client.prepare_robots()
        self.assertGreaterEqual(session.calls[1][1] - session.calls[0][1], 7)

    def test_non_retryable_error_and_redirect_stop(self):
        for status in (404, 302):
            client, session = self.client([response(), response(status=status)])
            client.prepare_robots()
            with self.assertRaises(RuntimeError):
                client.get(BASE)
            self.assertEqual(len(session.calls), 2)

    def test_added_updated_removed_and_unchanged(self):
        row = parse_book(fixture("book.html"), BOOK)
        self.assertEqual(diff([row], [dict(row)]), [])
        for key in ("title", "price_gbp", "stock", "rating", "category"):
            self.assertEqual(diff([row], [{**row, key: "changed"}])[0]["change"], "updated")
        self.assertEqual(diff([], [row])[0]["change"], "added")
        self.assertEqual(diff([row], [])[0]["change"], "removed")

    def test_csv_excel_and_second_run_empty(self):
        row = parse_book(fixture("book.html"), BOOK)
        with tempfile.TemporaryDirectory(dir=ROOT / ".tmp") as temporary:
            output = Path(temporary)
            folder, count = save_run([row], output)
            self.assertEqual(count, 1)
            with (folder / "changes.csv").open(encoding="utf-8-sig", newline="") as stream:
                rows = list(csv.DictReader(stream))
            self.assertEqual(rows[0]["title"], row["title"])
            workbook = load_workbook(folder / "changes.xlsx")
            self.assertEqual(workbook.active["C2"].value, row["title"])
            self.assertEqual(workbook.active["F2"].value, 3)
            workbook.close()
            folder, count = save_run([row], output)
            self.assertEqual(count, 0)
            with (folder / "changes.csv").open(encoding="utf-8-sig", newline="") as stream:
                self.assertEqual(list(csv.DictReader(stream)), [])
            workbook = load_workbook(folder / "changes.xlsx")
            self.assertEqual(workbook.active.max_row, 1)
            workbook.close()

    def test_export_failure_preserves_baseline(self):
        row = parse_book(fixture("book.html"), BOOK)
        with tempfile.TemporaryDirectory(dir=ROOT / ".tmp") as temporary:
            output = Path(temporary)
            save_run([row], output)
            previous = (output / "latest.json").read_bytes()
            with patch("scraper.export", side_effect=OSError("disk full")):
                with self.assertRaises(OSError):
                    save_run([{**row, "price_gbp": "1.00"}], output)
            self.assertEqual((output / "latest.json").read_bytes(), previous)

    def test_spreadsheet_formula_is_text(self):
        row = {**parse_book(fixture("book.html"), BOOK), "title": "=1+1"}
        with tempfile.TemporaryDirectory(dir=ROOT / ".tmp") as temporary:
            folder, _ = save_run([row], Path(temporary))
            workbook = load_workbook(folder / "changes.xlsx")
            self.assertEqual(workbook.active["C2"].data_type, "s")
            workbook.close()
            self.assertIn("'=1+1", (folder / "changes.csv").read_text(encoding="utf-8-sig"))

    def test_invalid_config(self):
        original = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        for key, value in [("request_interval_seconds", 0.9), ("max_retries", -1),
                           ("max_retries", 1.5), ("output_dir", "../outside"),
                           ("timeout_seconds", float("inf"))]:
            with tempfile.TemporaryDirectory(dir=ROOT / ".tmp") as temporary:
                path = Path(temporary) / "config.json"
                path.write_text(json.dumps({**original, key: value}), encoding="utf-8")
                with self.assertRaises(ValueError):
                    load_config(path)


if __name__ == "__main__":
    unittest.main()
