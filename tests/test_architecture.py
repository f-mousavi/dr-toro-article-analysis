import tempfile
import unittest
from datetime import date
from pathlib import Path

from doctoreto.aggregate import monthly_categories
from doctoreto.config import CrawlConfig
from doctoreto.excel_writer import CELL_LIMIT, build_workbook, save_workbook
from doctoreto.fetchers import FetchError, RequestsFetcher, SeleniumFetcher, create_fetcher
from doctoreto.models import ArchiveEntry, ArticleRecord
from doctoreto.parser import parse_archive, parse_article
from doctoreto.pipeline import collect_records, discover_articles


ARTICLE_HTML = """
<meta property="article:published_time" content="2026-11-20T10:00:00" />
<div class="post-info__row"><span class="post-info__label">دسته‌بندی</span>
<div class="post-info__chips"><a>سلامت</a><a>سلامت</a><a>تغذیه</a></div></div>
<div class="post-info__row"><span class="post-info__label">برچسب</span>
<div class="post-info__chips"><a>آموزش</a></div></div>
<div class="entry-content"><p>متن مقاله</p><aside>تبلیغ</aside><img src="x"></div>
"""


class FakeFetcher:
    def __init__(self, pages):
        self.pages = pages
        self.closed = False

    def fetch(self, url):
        if url not in self.pages:
            raise FetchError(url)
        return self.pages[url]

    def close(self):
        self.closed = True


class ArchitectureTests(unittest.TestCase):
    def test_parser_and_deduplication(self):
        archive = '<article class="post-card"><h3 class="post-card__title"><a href="/blog/a/">عنوان</a></h3></article>'
        entries = parse_archive(archive, "https://doctoreto.com/blog/all-posts/")
        self.assertEqual(entries, [ArchiveEntry("https://doctoreto.com/blog/a/", "عنوان")])
        record = parse_article(ARTICLE_HTML, entries[0])
        self.assertEqual(record.categories, ["سلامت", "تغذیه"])
        self.assertEqual(record.tags, ["آموزش"])
        self.assertEqual(record.published_date, date(2026, 11, 20))
        self.assertIn("متن مقاله", record.html)
        self.assertNotIn("تبلیغ", record.html)
        self.assertEqual(record.status, "ok")

    def test_missing_fields_and_fetch_error(self):
        entry = ArchiveEntry("https://example.com/a", "A")
        record = parse_article("<html></html>", entry)
        self.assertEqual(record.status, "error")
        self.assertEqual(len(record.issues), 3)
        failed = collect_records(FakeFetcher({}), [entry], 1)
        self.assertEqual(failed[0].status, "error")
        self.assertIn("خطای دریافت", failed[0].issues[0])

    def test_pipeline_order_and_monthly_aggregation(self):
        archive_url = "https://example.com/all-posts/"
        archive = '<article class="post-card"><h3 class="post-card__title"><a href="/a">A</a></h3></article>'
        fake = FakeFetcher({archive_url: archive, "https://example.com/a": ARTICLE_HTML})
        entries = discover_articles(fake, CrawlConfig(archive_url=archive_url, max_pages=1))
        records = collect_records(fake, entries, 1)
        self.assertEqual(records[0].title, "A")
        monthly = monthly_categories(records)
        self.assertEqual(len(monthly), 1)
        total, topics = next(iter(monthly.values()))
        self.assertEqual(total, 1)
        self.assertEqual(topics["سلامت"], 1)
        self.assertEqual(topics["تغذیه"], 1)

    def test_excel_chunks_and_safe_save(self):
        long_html = "x" * (CELL_LIMIT + 10)
        record = ArticleRecord("https://example.com/a", "A", date(2026, 11, 20), ["سلامت"], [], long_html)
        workbook = build_workbook([record])
        self.assertEqual(workbook["مقالات"]["E2"].value, None)
        self.assertEqual(workbook["مقالات"]["G2"].value, "در شیت بخش‌ها")
        chunks = list(workbook["HTMLهای طولانی - بخش‌ها"].values)[1:]
        self.assertEqual("".join(row[3] for row in chunks), long_html)
        workbook.close()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "result.xlsx"
            save_workbook([record], path)
            self.assertTrue(path.exists())
            with self.assertRaises(FileExistsError):
                save_workbook([record], path)

    def test_factory(self):
        self.assertIsInstance(create_fetcher("requests", 30, 2), RequestsFetcher)
        self.assertIsInstance(create_fetcher("selenium", 30, 2), SeleniumFetcher)
        with self.assertRaises(ValueError):
            create_fetcher("unknown", 30, 2)


if __name__ == "__main__":
    unittest.main()
