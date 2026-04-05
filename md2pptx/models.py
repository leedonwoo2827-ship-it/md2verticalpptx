"""Data classes for md2pptx pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class SlideData:
    """Parsed slide block from extended markdown."""

    index: int = 0
    template: str | None = None
    ref_slide: int | None = None
    reference_pptx: str | None = None
    fields: dict[str, str] = field(default_factory=dict)
    tables: list[dict[str, Any]] = field(default_factory=list)
    note: str | None = None


@dataclass
class SlideResult:
    """Outcome of building a single slide."""

    index: int
    ref_slide: int | None
    template: str | None
    status: str  # "success" | "failed" | "skipped"
    output_path: str = ""
    error: str | None = None
    elapsed: float = 0.0


@dataclass
class BuildSummary:
    """Aggregated build results."""

    results: list[SlideResult] = field(default_factory=list)
    output_path: str = ""
    total_time: float = 0.0

    @property
    def succeeded(self) -> int:
        return sum(1 for r in self.results if r.status == "success")

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if r.status == "failed")

    @property
    def total(self) -> int:
        return len(self.results)
