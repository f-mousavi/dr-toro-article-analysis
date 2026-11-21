"""Coordinates fetching, parsing and export without owning their details."""

from concurrent.futures import ThreadPoolExecutor, as_completed
from time import sleep

from .config import CrawlConfig
from .fetchers import FetchError, PageFetcher
from .models import ArchiveEntry, ArticleRecord
from .parser import parse_archive, parse_article


def discover_articles(fetcher: PageFetcher, config: CrawlConfig) -> list[ArchiveEntry]:
    unique: dict[str, ArchiveEntry] = {}
    for page_number in range(1, config.max_pages + 1):
        page_url = config.archive_url if page_number == 1 else f"{config.archive_url.rstrip('/')}/page/{page_number}/"
        entries = parse_archive(fetcher.fetch(page_url), page_url)
        if not entries:
            raise RuntimeError(f"No article cards found on archive page {page_number}: {page_url}")
        for entry in entries:
            unique[entry.url] = entry
        print(f"آرشیو صفحه {page_number}/{config.max_pages} خوانده شد؛ {len(unique)} مقاله یکتا")
        if page_number < config.max_pages:
            sleep(config.archive_delay)
    return list(unique.values())


def _fetch_article(fetcher: PageFetcher, entry: ArchiveEntry) -> ArticleRecord:
    try:
        return parse_article(fetcher.fetch(entry.url), entry)
    except FetchError as exc:
        return ArticleRecord(url=entry.url, title=entry.title, issues=[f"خطای دریافت: {exc}"])


def collect_records(fetcher: PageFetcher, entries: list[ArchiveEntry], workers: int) -> list[ArticleRecord]:
    records: list[ArticleRecord | None] = [None] * len(entries)
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(_fetch_article, fetcher, entry): index
            for index, entry in enumerate(entries)
        }
        for completed, future in enumerate(as_completed(futures), 1):
            records[futures[future]] = future.result()
            if completed % 100 == 0 or completed == len(entries):
                print(f"{completed}/{len(entries)} مقاله بررسی شد")
    return [record for record in records if record is not None]


def run(fetcher: PageFetcher, config: CrawlConfig, overwrite: bool = False) -> list[ArticleRecord]:
    from .excel_writer import save_workbook

    if config.output.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {config.output}")
    try:
        entries = discover_articles(fetcher, config)
        if not entries:
            raise RuntimeError("هیچ مقاله‌ای از آرشیو پیدا نشد")
        records = collect_records(fetcher, entries, config.workers)
    finally:
        fetcher.close()
    save_workbook(records, config.output, overwrite=overwrite)
    return records
