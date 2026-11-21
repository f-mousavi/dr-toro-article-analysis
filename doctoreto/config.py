from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CrawlConfig:
    archive_url: str = "https://doctoreto.com/blog/all-posts/"
    max_pages: int = 430
    workers: int = 4
    timeout: int = 30
    archive_delay: float = 0.2
    retries: int = 2
    output: Path = Path("articles_architecture_output.xlsx")

    def __post_init__(self) -> None:
        if self.max_pages < 1 or self.workers < 1 or self.timeout <= 0:
            raise ValueError("max_pages, workers and timeout must be positive")
        if self.archive_delay < 0 or self.retries < 0:
            raise ValueError("archive_delay and retries must be nonnegative")

