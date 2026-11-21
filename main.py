from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter, defaultdict
from datetime import date
from threading import local
from time import sleep

import requests
from bs4 import BeautifulSoup
import jdatetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font


ARCHIVE_URL = "https://doctoreto.com/blog/all-posts/"
LAST_PAGE = 430
HTML_OUTPUT_FILE = "articles_html_categories_tags_histogram.xlsx"
EXCEL_CELL_LIMIT = 32767
MAX_WORKERS = 4
REQUEST_TIMEOUT = 30

thread_state = local()

PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def gregorian_to_persian(value):
    if not value:
        return None
    persian = jdatetime.date.fromgregorian(
        year=value.year,
        month=value.month,
        day=value.day,
    )
    return persian.strftime("%Y/%m/%d").translate(PERSIAN_DIGITS)


def add_monthly_category_sheets(workbook, publication_rows):
    to_ascii_digits = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
    monthly_data = defaultdict(lambda: {"total": 0, "topics": Counter()})

    for published_date, category_text in publication_rows:
        if not published_date:
            continue
        parts = str(published_date).translate(to_ascii_digits).split("/")
        if len(parts) != 3:
            continue

        month_key = f"{parts[0]}-{parts[1]}"
        month = monthly_data[month_key]
        month["total"] += 1
        topics = {
            topic.strip()
            for topic in (category_text or "").replace(",", "،").split("،")
            if topic.strip()
        }
        if not topics:
            topics.add("بدون دسته‌بندی")
        month["topics"].update(topics)

    for month_key in sorted(monthly_data):
        month = monthly_data[month_key]
        total = month["total"]
        persian_month = month_key.translate(PERSIAN_DIGITS)
        sheet = workbook.create_sheet(f"ماه-{persian_month}")
        sheet.sheet_view.rightToLeft = True
        sheet.append(["موضوع", "تعداد مقالات", "درصد از مقالات همان ماه"])
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        for topic, count in month["topics"].most_common():
            sheet.append([topic, count, count / total])
            sheet.cell(sheet.max_row, 3).number_format = "0.0%"

        sheet["E1"] = f"تعداد مقالات این ماه: {total}"
        sheet["E2"] = "مقاله‌ها ممکن است چند دسته‌بندی داشته باشند."
        sheet["E3"] = "بنابراین درصدها الزاماً جمع ۱۰۰٪ نمی‌شوند."
        sheet.column_dimensions["A"].width = 45
        sheet.column_dimensions["B"].width = 18
        sheet.column_dimensions["C"].width = 28
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:C{sheet.max_row}"

        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"
        chart.overlap = 0
        chart.style = 10
        chart.title = f"درصد مقالات بر اساس موضوع - {persian_month}"
        chart.x_axis.title = "موضوع"
        chart.y_axis.title = "درصد مقالات"
        chart.y_axis.numFmt = "0%"
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 1
        chart.height = 15
        chart.width = 28
        chart.legend = None
        chart.add_data(
            Reference(sheet, min_col=3, min_row=1, max_row=sheet.max_row),
            titles_from_data=True,
            from_rows=False,
        )
        chart.set_categories(
            Reference(sheet, min_col=1, min_row=2, max_row=sheet.max_row)
        )
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showCatName = True
        chart.dataLabels.showVal = False
        chart.dataLabels.showSerName = False
        chart.dataLabels.dLblPos = "outEnd"
        sheet.add_chart(chart, "E5")


def get_session():
    if not hasattr(thread_state, "session"):
        thread_state.session = requests.Session()
        thread_state.session.headers.update(
            {"User-Agent": "Mozilla/5.0 (compatible; ArticleMetadataExporter/1.0)"}
        )
    return thread_state.session


def collect_articles():
    articles = {}
    session = requests.Session()
    session.headers.update(
        {"User-Agent": "Mozilla/5.0 (compatible; ArticleMetadataExporter/1.0)"}
    )

    for page_number in range(1, LAST_PAGE + 1):
        page_url = (
            ARCHIVE_URL
            if page_number == 1
            else f"{ARCHIVE_URL}page/{page_number}/"
        )
        response = session.get(page_url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        cards = soup.select("article.post-card")
        if not cards:
            raise RuntimeError(f"No article cards found on archive page {page_number}")

        for card in cards:
            link = card.select_one("h3.post-card__title a")
            if link:
                article_url = link.get("href")
                title = link.get_text(" ", strip=True)
                if article_url and title:
                    articles[article_url] = title

        print(
            f"آرشیو صفحه {page_number}/{LAST_PAGE} خوانده شد؛ "
            f"{len(articles)} مقاله یکتا"
        )
        sleep(0.2)

    return articles


def extract_categories(soup):
    categories = []
    for row in soup.select(".post-info__row"):
        label = row.select_one(".post-info__label")
        if label and "دسته‌بندی" in label.get_text(" ", strip=True):
            categories.extend(
                category.get_text(" ", strip=True)
                for category in row.select(".post-info__chips a")
                if category.get_text(" ", strip=True)
            )
    return "، ".join(dict.fromkeys(categories)) or None


def extract_tags(soup):
    tags = []
    for row in soup.select(".post-info__row"):
        label = row.select_one(".post-info__label")
        if label and "برچسب" in label.get_text(" ", strip=True):
            tags.extend(
                tag.get_text(" ", strip=True)
                for tag in row.select(".post-info__chips a")
                if tag.get_text(" ", strip=True)
            )
    return "، ".join(dict.fromkeys(tags)) or None


def clean_article_html(soup):
    content_element = soup.select_one(".entry-content")
    if content_element is None:
        content_element = soup.select_one(".wp-block-post-content")

    if content_element is None:
        article = soup.select_one("article")
        if article:
            content_element = article.select_one(".entry-content")
    if content_element is None:
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
    for selector_group in unwanted_selectors:
        for unwanted in content_element.select(selector_group):
            unwanted.decompose()

    for image in content_element.select("img"):
        figure = image.find_parent("figure")
        (figure or image).decompose()

    for heading in content_element.select("h1, h2, h3, h4, h5, h6"):
        heading_text = heading.get_text(" ", strip=True)
        if any(
            phrase in heading_text
            for phrase in (
                "اطلاعات مقاله",
                "پزشکان پیشنهادی",
                "مطالب مرتبط",
                "خلاصه مقاله",
            )
        ):
            for sibling in list(heading.find_next_siblings()):
                if sibling.name in ("h1", "h2", "h3"):
                    break
                sibling.decompose()
            heading.decompose()

    html = "".join(str(child) for child in content_element.contents).strip()
    return html or None


def get_article_data(article_url):
    response = get_session().get(article_url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    published_meta = soup.select_one('meta[property="article:published_time"]')
    date_text = published_meta.get("content") if published_meta else None
    if not date_text:
        time_tag = soup.select_one("time[datetime]")
        date_text = time_tag.get("datetime") if time_tag else None

    published_date = None
    date_error = None
    if date_text:
        try:
            published_date = date.fromisoformat(date_text[:10])
        except ValueError:
            date_error = f"قالب تاریخ نامعتبر است: {date_text}"
    else:
        date_error = "تاریخ انتشار در صفحه پیدا نشد"

    content = clean_article_html(soup)
    categories = extract_categories(soup)
    tags = extract_tags(soup)

    content_error = None
    if not content:
        content_error = "متن مقاله در صفحه پیدا نشد"

    errors = []
    if date_error:
        errors.append(date_error)
    if content_error:
        errors.append(content_error)
    if not categories:
        errors.append("دسته‌بندی در صفحه پیدا نشد")

    return (
        published_date,
        content,
        categories,
        tags,
        "; ".join(errors) if errors else None,
    )


def main():
    articles = collect_articles()
    if not articles:
        raise RuntimeError("هیچ مقاله‌ای از آرشیو پیدا نشد")

    article_items = list(articles.items())
    dates = [None] * len(article_items)
    texts = [None] * len(article_items)
    categories = [None] * len(article_items)
    tags = [None] * len(article_items)
    errors = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_article = {
            executor.submit(get_article_data, article_url): (index, article_url)
            for index, (article_url, _) in enumerate(article_items)
        }
        for completed_count, future in enumerate(as_completed(future_to_article), 1):
            index, article_url = future_to_article[future]
            try:
                published_date, text, category, article_tags, error = future.result()
            except requests.RequestException as exc:
                published_date = None
                text = None
                category = None
                article_tags = None
                error = str(exc)

            dates[index] = published_date
            texts[index] = text
            categories[index] = category
            tags[index] = article_tags
            if error:
                errors.append([article_url, error])

            if completed_count % 100 == 0 or completed_count == len(article_items):
                print(f"{completed_count}/{len(article_items)} مقاله بررسی شد")

    html_workbook = Workbook()
    html_sheet = html_workbook.active
    html_sheet.title = "مقالات"
    html_sheet.sheet_view.rightToLeft = True
    html_sheet.append(
        ["تاریخ انتشار", "عنوان", "دسته‌بندی", "برچسب‌ها", "متن مقاله HTML", "لینک"]
    )
    for cell in html_sheet[1]:
        cell.font = Font(bold=True)

    long_html_chunks = []
    for (article_url, title), published_date, text, category, article_tags in zip(
        article_items, dates, texts, categories, tags
    ):
        cell_html = text
        if text and len(text) > EXCEL_CELL_LIMIT:
            cell_html = None
            for chunk_number, start in enumerate(
                range(0, len(text), EXCEL_CELL_LIMIT), start=1
            ):
                long_html_chunks.append(
                    [article_url, title, chunk_number, text[start : start + EXCEL_CELL_LIMIT]]
                )
        html_sheet.append(
            [
                gregorian_to_persian(published_date),
                title,
                category,
                article_tags,
                cell_html,
                article_url,
            ]
        )

    html_sheet.column_dimensions["A"].width = 18
    html_sheet.column_dimensions["B"].width = 60
    html_sheet.column_dimensions["C"].width = 30
    html_sheet.column_dimensions["D"].width = 40
    html_sheet.column_dimensions["E"].width = 100
    html_sheet.column_dimensions["F"].width = 80
    html_sheet.freeze_panes = "A2"
    html_sheet.auto_filter.ref = html_sheet.dimensions

    if long_html_chunks:
        chunk_sheet = html_workbook.create_sheet("HTMLهای طولانی - بخش‌ها")
        chunk_sheet.sheet_view.rightToLeft = True
        chunk_sheet.append(["لینک مقاله", "عنوان", "شماره بخش", "بخش HTML"])
        for cell in chunk_sheet[1]:
            cell.font = Font(bold=True)
        for row in long_html_chunks:
            chunk_sheet.append(row)
        chunk_sheet.column_dimensions["A"].width = 80
        chunk_sheet.column_dimensions["B"].width = 60
        chunk_sheet.column_dimensions["C"].width = 14
        chunk_sheet.column_dimensions["D"].width = 100
        chunk_sheet.freeze_panes = "A2"
        chunk_sheet.auto_filter.ref = chunk_sheet.dimensions

    if errors:
        error_sheet = html_workbook.create_sheet("موارد نیازمند بررسی")
        error_sheet.sheet_view.rightToLeft = True
        error_sheet.append(["لینک مقاله", "خطا"])
        for cell in error_sheet[1]:
            cell.font = Font(bold=True)
        for row in errors:
            error_sheet.append(row)
        error_sheet.column_dimensions["A"].width = 80
        error_sheet.column_dimensions["B"].width = 50
        error_sheet.freeze_panes = "A2"
        error_sheet.auto_filter.ref = error_sheet.dimensions

    add_monthly_category_sheets(
        html_workbook,
        [
            (gregorian_to_persian(published_date), category)
            for published_date, category in zip(dates, categories)
        ],
    )

    html_workbook.save(HTML_OUTPUT_FILE)

    print(
        f"\nفایل خروجی ساخته شد: {HTML_OUTPUT_FILE} "
        f"({len(article_items)} مقاله، {len(errors)} مورد نیازمند بررسی)"
    )


if __name__ == "__main__":
    main()