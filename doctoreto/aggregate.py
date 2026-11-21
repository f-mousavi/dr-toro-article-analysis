"""Monthly counts independent of Excel formatting."""

from collections import Counter, defaultdict

import jdatetime

from .models import ArticleRecord


def monthly_categories(records: list[ArticleRecord]) -> dict[str, tuple[int, Counter[str]]]:
    totals: dict[str, int] = defaultdict(int)
    topics: dict[str, Counter[str]] = defaultdict(Counter)
    for record in records:
        if record.published_date is None:
            continue
        month = jdatetime.date.fromgregorian(date=record.published_date)
        key = f"{month.year:04d}-{month.month:02d}"
        totals[key] += 1
        for category in set(record.categories) or {"بدون دسته‌بندی"}:
            topics[key][category] += 1
    return {key: (totals[key], topics[key]) for key in sorted(totals)}
