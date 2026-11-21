from dataclasses import dataclass, field
from datetime import date


@dataclass(frozen=True)
class ArchiveEntry:
    url: str
    title: str


@dataclass
class ArticleRecord:
    url: str
    title: str
    published_date: date | None = None
    categories: list[str] = field(default_factory=list)
    tags: list[str] = field(default_factory=list)
    html: str | None = None
    issues: list[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        if not self.issues:
            return "ok"
        if self.published_date is None and self.html is None and not self.categories:
            return "error"
        return "partial"

