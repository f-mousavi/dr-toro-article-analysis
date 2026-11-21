# معماری جدید پروژه

`main.py` نسخهٔ قدیمی است و تغییر نکرده است. اجرای معماری جدید از `run_architecture.py` انجام می‌شود. خروجی پیش‌فرض آن `articles_architecture_output.xlsx` است و فایل اکسل قبلی را بازنویسی نمی‌کند. اگر همین نام خروجی از قبل وجود داشته باشد، اجرا با خطا متوقف می‌شود؛ برای جایگزینی آگاهانه از `--overwrite` استفاده کنید.

## جریان کار

```text
run_architecture.py → CrawlConfig → create_fetcher (Factory)
                                      ↓
                    RequestsFetcher یا SeleniumFetcher (Strategy)
                                      ↓ HTML
                    parser → ArticleRecord → aggregate → excel_writer
```

| بخش | مسئولیت |
| --- | --- |
| `doctoreto/config.py` | تنظیمات اجرا و اعتبارسنجی آن‌ها |
| `doctoreto/models.py` | دادهٔ نام‌دار مقاله و ورودی آرشیو |
| `doctoreto/fetchers.py` | دریافت HTML، دو Strategy و Factory انتخاب‌کننده |
| `doctoreto/parser.py` | تبدیل HTML به داده؛ بدون اینترنت و Excel |
| `doctoreto/aggregate.py` | شمارش دسته‌بندی‌ها به تفکیک ماه |
| `doctoreto/excel_writer.py` | ساخت شیت‌ها، نمودارها و ذخیرهٔ امن |
| `doctoreto/pipeline.py` | ترتیب اجرای مراحل و مدیریت هم‌زمانی |

## الگوهای طراحی

**Strategy:** `PageFetcher` قرارداد `fetch(url)` و `close()` را مشخص می‌کند. پیاده‌سازی `RequestsFetcher` پاسخ HTTP را می‌گیرد؛ `SeleniumFetcher` مرورگر Chrome را اجرا می‌کند و بعد از حضور بخش اصلی صفحه، HTML رندرشده را می‌دهد. `parser` با هیچ‌یک از این دو وابسته نیست.

**Factory:** تابع `create_fetcher` بر اساس گزینهٔ `--fetcher` پیاده‌سازی مناسب را می‌سازد. افزودن یک روش دریافت دیگر به این نقطه محدود می‌شود.

**Dependency Injection:** `pipeline` شیء fetcher را به‌عنوان ورودی می‌گیرد. در آزمون‌ها یک `FakeFetcher` جای آن قرار می‌گیرد، بنابراین آزمون به سایت زنده وابسته نیست.

**Pipeline و Separation of Concerns:** کشف لینک، دریافت HTML، پارس، شمارش و نوشتن Excel جدا هستند. تابع پارس می‌تواند روی یک HTML ذخیره‌شده آزموده شود؛ نویسندهٔ Excel نیز رکوردهای ساختگی می‌پذیرد.

**OOP به‌اندازهٔ نیاز:** fetcherها کلاس‌اند چون رفتار و منبع قابل‌جایگزین دارند. `ArticleRecord` یک dataclass است. پارس و تجمیع تابع باقی مانده‌اند چون حالت داخلی نیاز ندارند.

## اجرا

Python 3.10 یا جدیدتر لازم است. اگر در ویندوز دستور `python` شناخته نمی‌شود، پس از نصب Python می‌توانید معادل آن را با `py` اجرا کنید.

```powershell
python -m pip install -r requirements.txt
python run_architecture.py --max-pages 1
```

برای Selenium ابتدا `python -m pip install selenium` و Chrome/ChromeDriver سازگار لازم است؛ سپس:

```powershell
python run_architecture.py --fetcher selenium --max-pages 1
```

روش Selenium به‌طور پیش‌فرض یک worker دارد، چون برای هر صفحه یک مرورگر باز و بسته می‌شود و کندتر است. استفاده از `--workers` بیشتر با Selenium فقط با توجه به منابع سیستم توصیه می‌شود. در حالت عادی، `requests` پیش‌فرض و چهار worker دارد.

اجرای آزمون‌های آفلاین:

```powershell
python -m unittest discover -s tests -v
```

## تفاوت‌های آگاهانه با خروجی قدیمی

- شش ستون اصلی شیت «مقالات» حفظ شده و ستون هفتم «وضعیت HTML» اضافه شده تا متن خالی از متن منتقل‌شده به شیت بخش‌ها قابل تشخیص باشد.
- شیت خطاها ستون «وضعیت» دارد.
- فایل جدید پیش‌فرض جداست و بدون `--overwrite` بازنویسی نمی‌شود.
- خطاهای گذرای HTTP retry محدود دارند. خطای آرشیو همچنان اجرای کامل را متوقف می‌کند، چون خروجی ناقصِ بدون اطلاع مطلوب نیست.
- این بازآرایی به سایت زنده وابسته است؛ selectorهای HTML و رفتار Selenium باید با چند صفحهٔ واقعی بررسی شوند. آزمون‌ها فقط قراردادهای محلی را تأیید می‌کنند.
