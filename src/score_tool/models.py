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


REVIEW_RESULT_OK = "核查无误"
REVIEW_RESULT_MODIFIED = "已修改（登记错误）"
REVIEW_RESULT_NO_SUBMISSION = "未交卷/未给分"


@dataclass(frozen=True)
class ReviewSessionItem:
    key: str
    index: int
    level: str
    student_id: str
    student_name: str
    class_name: str
    score_column: str
    score: Any
    reasons: tuple[str, ...]


@dataclass(frozen=True)
class ReviewDecision:
    item_key: str
    result: str
    corrected_score: Any = None


@dataclass
class MergeData:
    source_previews: list[SourcePreview]
    students: list[dict[str, Any]]
    score_columns: list[str]
    warnings: list[str]
    main_grade: int | None
    review_flags: list[ReviewFlag]


@dataclass(frozen=True)
class ReviewExportResult:
    report_path: Path
    corrected_output_path: Path | None = None
    filtered_report_path: Path | None = None


@dataclass
class ProcessResult:
    output_path: Path
    source_previews: list[SourcePreview]
    row_count: int
    score_columns: list[str]
    review_count: int = 0
    focus_review_count: int = 0
    merge_data: MergeData | None = None
    warnings: list[str] = field(default_factory=list)
