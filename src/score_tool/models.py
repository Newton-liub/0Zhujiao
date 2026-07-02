from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SourceFileConfig:
    path: Path
    score_name: str | None = None


@dataclass(frozen=True)
class StudentScoreRow:
    student_id: str
    student_name: str
    class_name: str
    score: Any
    source_path: Path


@dataclass
class SourcePreview:
    path: Path
    sheet_name: str
    header_row: int
    student_count: int
    score_name: str
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ReviewFlag:
    level: str
    student_id: str
    student_name: str
    class_name: str
    score_column: str
    score: Any
    reason: str
    suggestion: str


@dataclass
class ProcessResult:
    output_path: Path
    source_previews: list[SourcePreview]
    row_count: int
    score_columns: list[str]
    review_count: int = 0
    focus_review_count: int = 0
    warnings: list[str] = field(default_factory=list)
