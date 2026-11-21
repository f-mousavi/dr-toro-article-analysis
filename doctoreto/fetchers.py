"""Fetching strategies. Selenium is optional and loaded only when selected."""

from threading import local
from typing import Protocol

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class PageFetcher(Protocol):
    def fetch(self, url: str) -> str: ...

    def close(self) -> None: ...


class FetchError(Exception):
    """Expected failure while retrieving a page."""


class RequestsFetcher:
    def __init__(self, timeout: int = 30, retries: int = 2):
        self.timeout = timeout
        self.retries = retries
        self._state = local()
        self._sessions: list[requests.Session] = []

    def _session(self) -> requests.Session:
        if not hasattr(self._state, "session"):
            session = requests.Session()
            session.headers.update(
                {"User-Agent": "Mozilla/5.0 (compatible; ArticleMetadataExporter/2.0)"}
            )
            retry = Retry(
                total=self.retries,
                backoff_factor=0.5,
                status_forcelist=(429, 500, 502, 503, 504),
                allowed_methods=("GET",),
            )
            session.mount("https://", HTTPAdapter(max_retries=retry))
            session.mount("http://", HTTPAdapter(max_retries=retry))
            self._state.session = session
            self._sessions.append(session)
        return self._state.session

    def fetch(self, url: str) -> str:
        try:
            response = self._session().get(url, timeout=self.timeout)
            response.raise_for_status()
            return response.text
        except requests.RequestException as exc:
            raise FetchError(f"{url}: {exc}") from exc

    def close(self) -> None:
        for session in self._sessions:
            session.close()
        self._sessions.clear()


class SeleniumFetcher:
    """Browser strategy for pages whose content requires JavaScript.

    Creates and closes one browser per fetch. Prefer one worker for this strategy.
    """

    def __init__(self, timeout: int = 30):
        self.timeout = timeout

    def fetch(self, url: str) -> str:
        try:
            from selenium import webdriver
            from selenium.common.exceptions import WebDriverException
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait
        except ImportError as exc:
            raise RuntimeError("Install selenium to use --fetcher selenium") from exc

        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        try:
            driver = webdriver.Chrome(options=options)
        except WebDriverException as exc:
            raise FetchError(f"{url}: {exc}") from exc
        try:
            driver.set_page_load_timeout(self.timeout)
            driver.get(url)
            selector = "article.post-card" if "/all-posts/" in url else ".entry-content, .wp-block-post-content"
            WebDriverWait(driver, self.timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            )
            return driver.page_source
        except WebDriverException as exc:
            raise FetchError(f"{url}: {exc}") from exc
        finally:
            driver.quit()

    def close(self) -> None:
        pass


def create_fetcher(kind: str, timeout: int, retries: int) -> PageFetcher:
    """Factory: choose the fetching strategy from configuration."""
    if kind == "requests":
        return RequestsFetcher(timeout=timeout, retries=retries)
    if kind == "selenium":
        return SeleniumFetcher(timeout=timeout)
    raise ValueError(f"Unknown fetcher: {kind}")
