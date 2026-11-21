"""Pure HTML parsing: no network or Excel dependency."""

from datetime import date
from urllib.parse import urljoin

from bs4 import BeautifulSoup

from .models import ArchiveEntry, ArticleRecord


def parse_archive(html: str, page_url: str) -> list[ArchiveEntry]:
    soup = BeautifulSoup(html, "html.parser")
    entries: list[ArchiveEntry] = []
    for card in soup.select("article.post-card"):
        link = card.select_one("h3.post-card__title a")
        if link and link.get("href"):
            title = link.get_text(" ", strip=True)
            if title:
                entries.append(ArchiveEntry(urljoin(page_url, link["href"]), title))
    return entries


def _chips(soup: BeautifulSoup, label_text: str) -> list[str]:
    values: list[str] = []
    for row in soup.select(".post-info__row"):
        label = row.select_one(".post-info__label")
        if label and label_text in label.get_text(" ", strip=True):
            values.extend(
                item.get_text(" ", strip=True)
                for item in row.select(".post-info__chips a")
            )
    return list(dict.fromkeys(value for value in values if value))


def _article_html(soup: BeautifulSoup) -> str | None:
    content = soup.select_one(".entry-content, .wp-block-post-content")
    if content is None:
        return None

    unwanted_selectors = (
        "script, style, nav, aside, form, audio, video",
        "figure.wp-block-audio, .img-slider, .wp-block-drto-image-slider",
        ".post-actions, .post-tldr, .post-sidebar",
        ".post-info, .related-posts, .latest-posts",
        ".recommended-drs, [class*='recommended-drs']",
        ".wp-block-drto-recommended-drs",
        "[class*='drto-podcast-card']",
        "[class*='drto-banner-block']",
        "[class*='drto-specialty-banner']",
        "[class*='wp-block-drto-service-banner']",
        "[class*='related-post']",
    )
    for selector in unwanted_selectors:
        for element in content.select(selector):
            element.decompose()

    for image in content.select("img"):
        figure = image.find_parent("figure")
        (figure or image).decompose()

    for heading in content.select("h1, h2, h3, h4, h5, h6"):
        text = heading.get_text(" ", strip=True)
        if any(term in text for term in (
            "اطلاعات مقاله", "پزشکان پیشنهادی", "مطالب مرتبط", "خلاصه مقاله"
        )):
            for sibling in list(heading.find_next_siblings()):
                if sibling.name in ("h1", "h2", "h3"):
                    break
                sibling.decompose()
            heading.decompose()

    result = "".join(str(child) for child in content.contents).strip()
    return result or None


def parse_article(html: str, entry: ArchiveEntry) -> ArticleRecord:
    soup = BeautifulSoup(html, "html.parser")
    record = ArticleRecord(url=entry.url, title=entry.title)

    published_meta = soup.select_one('meta[property="article:published_time"]')
    date_text = published_meta.get("content") if published_meta else None
    if not date_text:
        time_tag = soup.select_one("time[datetime]")
        date_text = time_tag.get("datetime") if time_tag else None
    if date_text:
        try:
            record.published_date = date.fromisoformat(date_text[:10])
        except ValueError:
            record.issues.append(f"قالب تاریخ نامعتبر است: {date_text}")
    else:
        record.issues.append("تاریخ انتشار در صفحه پیدا نشد")

    record.html = _article_html(soup)
    record.categories = _chips(soup, "دسته‌بندی")
    record.tags = _chips(soup, "برچسب")
    if not record.html:
        record.issues.append("متن مقاله در صفحه پیدا نشد")
    if not record.categories:
        record.issues.append("دسته‌بندی در صفحه پیدا نشد")
    return record
