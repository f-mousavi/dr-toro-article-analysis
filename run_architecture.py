"""Run the modular implementation. The original main.py remains independent."""

import argparse
from pathlib import Path

from doctoreto.config import CrawlConfig
from doctoreto.fetchers import create_fetcher
from doctoreto.pipeline import run


def main() -> None:
    parser = argparse.ArgumentParser(description="Export Doctoreto articles to Excel")
    parser.add_argument("--fetcher", choices=("requests", "selenium"), default="requests")
    parser.add_argument("--output", type=Path, default=Path("articles_architecture_output.xlsx"))
    parser.add_argument("--max-pages", type=int, default=430)
    parser.add_argument("--workers", type=int, default=None)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--overwrite", action="store_true", help="Replace the selected output file")
    args = parser.parse_args()

    workers = args.workers if args.workers is not None else (1 if args.fetcher == "selenium" else 4)
    config = CrawlConfig(
        max_pages=args.max_pages, workers=workers, timeout=args.timeout,
        retries=args.retries, output=args.output,
    )
    fetcher = create_fetcher(args.fetcher, timeout=config.timeout, retries=config.retries)
    records = run(fetcher, config, overwrite=args.overwrite)
    counts = {status: sum(record.status == status for record in records) for status in ("ok", "partial", "error")}
    print(f"فایل ساخته شد: {config.output} | تعداد: {len(records)} | وضعیت‌ها: {counts}")


if __name__ == "__main__":
    main()
