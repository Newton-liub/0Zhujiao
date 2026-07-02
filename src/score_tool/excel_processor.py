from __future__ import annotations

import re
from collections import Counter, OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import ProcessResult, SourceFileConfig, SourcePreview, StudentScoreRow

REQUIRED_HEADERS = {
    "student_id": ("学号/工号", "学号", "工号"),
    "student_name": ("学生姓名", "姓名"),
    "class_name": ("班级",),
    "score": ("总分",),
}

MAX_HEADER_SCAN_ROWS = 20


@dataclass(frozen=True)
class ParsedClass:
    raw: str
    prefix: str
    grade: int | None
    major: int | None
    class_no: int | None
    tail: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\u3000", " ").strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


def normalize_student_id(value: Any) -> str:
    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def is_blank_row(values: Iterable[Any]) -> bool:
    return all(normalize_text(value) == "" for value in values)


def detect_header(row_values: list[Any]) -> dict[str, int] | None:
    headers = [normalize_header(value) for value in row_values]
    found: dict[str, int] = {}
    for key, candidates in REQUIRED_HEADERS.items():
        for index, header in enumerate(headers):
            if header in candidates:
                found[key] = index
                break
        if key not in found:
            return None
    return found


def infer_score_name(path: Path) -> str:
    stem = path.stem.strip()
    patterns = [
        r"(阶段\s*测试\s*\d+)",
        r"(阶段\s*测验\s*\d+)",
        r"(测试\s*\d+)",
        r"(测验\s*\d+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, stem)
        if match:
            return re.sub(r"\s+", "", match.group(1))

    separators = ["-", "_", "—", "–"]
    for separator in separators:
        if separator in stem:
            candidate = stem.split(separator)[-1].strip()
            if candidate:
                return candidate
    return stem or "总分"


def unique_names(names: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    result: list[str] = []
    for name in names:
        base = name or "总分"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return result


def parse_class_name(class_name: str) -> ParsedClass:
    text = normalize_text(class_name)
    match = re.search(r"^(?P<prefix>\D*?)(?P<grade>\d{2})(?P<major>\d{2})-(?P<class_no>\d+)(?P<tail>.*)$", text)
    if match:
        return ParsedClass(
            raw=text,
            prefix=match.group("prefix"),
            grade=int(match.group("grade")),
            major=int(match.group("major")),
            class_no=int(match.group("class_no")),
            tail=match.group("tail"),
        )

    loose = re.search(r"(?P<grade>\d{2})", text)
    return ParsedClass(
        raw=text,
        prefix=re.sub(r"\d.*$", "", text),
        grade=int(loose.group("grade")) if loose else None,
        major=None,
        class_no=None,
        tail="",
    )


def natural_student_id_key(student_id: str) -> tuple[int, Any]:
    text = normalize_student_id(student_id)
    if re.fullmatch(r"\d+", text):
        return (0, int(text))
    return (1, text)


def class_sort_key(class_name: str) -> tuple[Any, ...]:
    parsed = parse_class_name(class_name)
    return (
        parsed.prefix,
        parsed.grade if parsed.grade is not None else 999,
        parsed.major if parsed.major is not None else 999,
        parsed.class_no if parsed.class_no is not None else 999,
        parsed.tail,
        parsed.raw,
    )


def detect_main_grade(class_names: Iterable[str]) -> int | None:
    grades = [parse_class_name(name).grade for name in class_names]
    known_grades = [grade for grade in grades if grade is not None]
    if not known_grades:
        return None
    return Counter(known_grades).most_common(1)[0][0]


def is_retake(class_name: str, main_grade: int | None) -> bool:
    if main_grade is None:
        return False
    grade = parse_class_name(class_name).grade
    return grade is not None and grade != main_grade


def read_source(config: SourceFileConfig) -> tuple[SourcePreview, list[StudentScoreRow]]:
    path = config.path
    warnings: list[str] = []
    workbook = load_workbook(path, data_only=True, read_only=True)

    selected_sheet = None
    header_row_number = 0
    header_map: dict[str, int] | None = None

    for worksheet in workbook.worksheets:
        for row_index, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=min(MAX_HEADER_SCAN_ROWS, worksheet.max_row), values_only=True),
            start=1,
        ):
            detected = detect_header(list(row))
            if detected is not None:
                selected_sheet = worksheet
                header_row_number = row_index
                header_map = detected
                break
        if selected_sheet is not None:
            break

    if selected_sheet is None or header_map is None:
        raise ValueError(f"{path.name} 未找到必需表头：学号/工号、学生姓名、班级、总分")

    rows: list[StudentScoreRow] = []
    seen_ids: set[str] = set()
    max_col = selected_sheet.max_column
    data_start = header_row_number + 1

    for row in selected_sheet.iter_rows(min_row=data_start, max_col=max_col, values_only=True):
        if is_blank_row(row):
            continue

        student_id = normalize_student_id(row[header_map["student_id"]] if header_map["student_id"] < len(row) else None)
        student_name = normalize_text(row[header_map["student_name"]] if header_map["student_name"] < len(row) else None)
        class_name = normalize_text(row[header_map["class_name"]] if header_map["class_name"] < len(row) else None)
        score = row[header_map["score"]] if header_map["score"] < len(row) else None

        if not student_id:
            continue
        if student_id in seen_ids:
            warnings.append(f"{path.name} 中学号/工号 {student_id} 重复，已保留最后一次成绩")
            rows = [item for item in rows if item.student_id != student_id]
        seen_ids.add(student_id)

        if not student_name:
            warnings.append(f"{path.name} 中学号/工号 {student_id} 缺少学生姓名")
        if not class_name:
            warnings.append(f"{path.name} 中学号/工号 {student_id} 缺少班级")
        if score is None or normalize_text(score) == "":
            warnings.append(f"{path.name} 中学号/工号 {student_id} 总分为空")

        rows.append(
            StudentScoreRow(
                student_id=student_id,
                student_name=student_name,
                class_name=class_name,
                score=score,
                source_path=path,
            )
        )

    score_name = normalize_text(config.score_name) or infer_score_name(path)
    preview = SourcePreview(
        path=path,
        sheet_name=selected_sheet.title,
        header_row=header_row_number,
        student_count=len(rows),
        score_name=score_name,
        warnings=warnings,
    )
    return preview, rows


def preview_sources(configs: list[SourceFileConfig]) -> list[SourcePreview]:
    previews: list[SourcePreview] = []
    raw_names: list[str] = []
    for config in configs:
        preview, _ = read_source(config)
        previews.append(preview)
        raw_names.append(preview.score_name)

    for preview, score_name in zip(previews, unique_names(raw_names), strict=True):
        preview.score_name = score_name
    return previews


def merge_sources(configs: list[SourceFileConfig], output_path: Path, include_log_sheet: bool = True) -> ProcessResult:
    if not configs:
        raise ValueError("请至少选择一个 Excel 文件")

    previews: list[SourcePreview] = []
    rows_by_source: list[list[StudentScoreRow]] = []
    raw_score_names: list[str] = []

    for config in configs:
        preview, rows = read_source(config)
        previews.append(preview)
        rows_by_source.append(rows)
        raw_score_names.append(preview.score_name)

    score_columns = unique_names(raw_score_names)
    for preview, score_name in zip(previews, score_columns, strict=True):
        preview.score_name = score_name

    students: OrderedDict[str, dict[str, Any]] = OrderedDict()
    warnings: list[str] = []
    for preview in previews:
        warnings.extend(preview.warnings)

    for source_index, rows in enumerate(rows_by_source):
        score_column = score_columns[source_index]
        for row in rows:
            current = students.setdefault(
                row.student_id,
                {
                    "学号/工号": row.student_id,
                    "学生姓名": row.student_name,
                    "班级": row.class_name,
                    "scores": {},
                },
            )

            if current["学生姓名"] and row.student_name and current["学生姓名"] != row.student_name:
                warnings.append(f"学号/工号 {row.student_id} 姓名不一致：{current['学生姓名']} / {row.student_name}")
            if current["班级"] and row.class_name and current["班级"] != row.class_name:
                warnings.append(f"学号/工号 {row.student_id} 班级不一致：{current['班级']} / {row.class_name}")

            if not current["学生姓名"] and row.student_name:
                current["学生姓名"] = row.student_name
            if not current["班级"] and row.class_name:
                current["班级"] = row.class_name
            current["scores"][score_column] = row.score

    main_grade = detect_main_grade(student["班级"] for student in students.values())
    for student in students.values():
        if parse_class_name(student["班级"]).grade is None:
            warnings.append(f"学号/工号 {student['学号/工号']} 的班级无法识别年级：{student['班级']}")

    ordered_students = sorted(
        students.values(),
        key=lambda student: (
            0 if is_retake(student["班级"], main_grade) else 1,
            class_sort_key(student["班级"]),
            natural_student_id_key(student["学号/工号"]),
        ),
    )

    write_output(output_path, ordered_students, score_columns, previews, warnings, include_log_sheet, main_grade)
    return ProcessResult(
        output_path=output_path,
        source_previews=previews,
        row_count=len(ordered_students),
        score_columns=score_columns,
        warnings=warnings,
    )


def write_output(
    output_path: Path,
    students: list[dict[str, Any]],
    score_columns: list[str],
    previews: list[SourcePreview],
    warnings: list[str],
    include_log_sheet: bool,
    main_grade: int | None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "成绩汇总"

    headers = ["学号/工号", "学生姓名", "班级", *score_columns]
    worksheet.append(headers)
    for student in students:
        worksheet.append(
            [
                student["学号/工号"],
                student["学生姓名"],
                student["班级"],
                *[student["scores"].get(score_column, "") for score_column in score_columns],
            ]
        )

    style_worksheet(worksheet)

    if include_log_sheet:
        log_sheet = workbook.create_sheet("处理日志")
        log_sheet.append(["项目", "内容"])
        log_sheet.append(["主年级", main_grade if main_grade is not None else "未识别"])
        log_sheet.append(["输出人数", len(students)])
        log_sheet.append([])
        log_sheet.append(["源文件", "工作表", "表头行", "人数", "成绩列名"])
        for preview in previews:
            log_sheet.append([preview.path.name, preview.sheet_name, preview.header_row, preview.student_count, preview.score_name])
        log_sheet.append([])
        log_sheet.append(["警告"])
        for warning in warnings:
            log_sheet.append([warning])
        style_worksheet(log_sheet)

    workbook.save(output_path)


def style_worksheet(worksheet: Any) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    worksheet.freeze_panes = "A2"
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = normalize_text(cell.value)
            max_length = max(max_length, len(value))
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 40)


def build_configs(paths: list[str | Path], score_names: dict[str, str] | None = None) -> list[SourceFileConfig]:
    score_names = score_names or {}
    configs: list[SourceFileConfig] = []
    for item in paths:
        path = Path(item)
        configs.append(SourceFileConfig(path=path, score_name=score_names.get(str(path))))
    return configs