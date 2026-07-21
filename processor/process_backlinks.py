"""
合并并去重多个 backlinks Excel 文件。

默认 CLI 行为保持不变：
- 从 processor/原始外链地址 读取 xlsx/csv
- 输出到 processor/域名清洗数据/YYYY-MM-DD.xlsx
- 使用 processed_files.json 跳过已处理文件

Web 后端会调用 process_backlink_files()，传入任务自己的上传目录和输出路径。
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlparse

import pandas as pd
import tldextract


BASE_DIR = Path(__file__).resolve().parent
SOURCE_DIR = BASE_DIR / "原始外链地址"
OUTPUT_DIR = BASE_DIR / "域名清洗数据"
TODAY = date.today().strftime("%Y-%m-%d")
OUTPUT_FILE = OUTPUT_DIR / f"{TODAY}.xlsx"
PROCESSED_LOG = BASE_DIR / "processed_files.json"
SUPPORTED_SUFFIXES = {".csv", ".xlsx"}
EXTRACTOR = tldextract.TLDExtract(suffix_list_urls=(), cache_dir=None)
UNICODE_DASHES = str.maketrans(
    {
        "_": "-",
        "‐": "-",
        "‑": "-",
        "‒": "-",
        "–": "-",
        "—": "-",
        "―": "-",
        "−": "-",
    }
)
SEO_CARTEL_PATTERN = re.compile(r"\bseo[\s-]*cartel\b")
TELEGRAM_MARKER_PATTERN = re.compile(
    r"(?:\btelegram\b|\btg\s*@|(?:https?://)?(?:www\.)?t\.me/)"
)
SEO_SPAM_SIGNAL_PATTERN = re.compile(
    r"(?:"
    r"\bseo\b|"
    r"back[\s-]?links?|"
    r"black[\s-]?(?:hat|links?)|"
    r"bulk\s+link\s+posting|"
    r"mass\s+back[\s-]?link(?:ing)?|"
    r"link\s+indexing|"
    r"hacked\s+sites?|"
    r"homepage\s+links?|"
    r"cross[\s-]?links?|"
    r"traffic\s+boost"
    r")"
)

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float, str], None]


class ProcessorPaused(Exception):
    """Raised when processor progress has stopped at a safe file boundary."""


@dataclass
class ProcessorResult:
    output_file: str
    input_files: list[str]
    failed_files: list[str]
    rows_read: int
    rows_output: int
    duplicate_rows_removed: int
    filtered_rows: int
    filtered_search_yahoo_rows: int
    filtered_telegram_seo_spam_rows: int
    skipped_processed_files: int

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def emit_log(logger: LogCallback | None, message: str) -> None:
    if logger:
        logger(message)


def emit_progress(progress_callback: ProgressCallback | None, percent: float, message: str) -> None:
    if progress_callback:
        progress_callback(max(0.0, min(100.0, percent)), message)


def load_processed(processed_log: Path = PROCESSED_LOG) -> set[str]:
    """加载已处理文件列表，容错空文件或损坏的 JSON。"""
    if processed_log.exists():
        try:
            content = processed_log.read_text(encoding="utf-8").strip()
            if not content:
                return set()
            data = json.loads(content)
            if isinstance(data, list):
                return set(str(item) for item in data)
            try:
                return set(data)
            except Exception:
                return set()
        except json.JSONDecodeError:
            try:
                processed_log.write_text("[]\n", encoding="utf-8")
            except Exception:
                pass
            return set()
        except Exception as exc:
            print(f"加载处理记录失败，忽略并继续：{exc}")
            return set()
    return set()


def save_processed(processed: set[str], processed_log: Path = PROCESSED_LOG) -> None:
    processed_log.parent.mkdir(parents=True, exist_ok=True)
    processed_log.write_text(
        json.dumps(sorted(processed), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def get_registered_domain(url: Any) -> str:
    """提取注册域名（如 example.com、example.co.uk）。"""
    if not isinstance(url, str) or not url.strip():
        return ""
    try:
        extracted = EXTRACTOR(url)
        if extracted.domain and extracted.suffix:
            return f"{extracted.domain}.{extracted.suffix}".lower()
        return url.lower()
    except Exception:
        return url.lower()


def normalize_filter_text(value: Any) -> str:
    """Normalize user-controlled export text for stable spam matching."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize("NFKC", str(value)).casefold()
    text = text.translate(UNICODE_DASHES)
    return re.sub(r"\s+", " ", text).strip()


def get_url_hostname(url: Any) -> str:
    """Return a normalized hostname for absolute or scheme-less URLs."""
    text = str(url).strip() if url is not None else ""
    if not text:
        return ""
    candidate = text if "://" in text else f"//{text}"
    try:
        return (urlparse(candidate).hostname or "").casefold().rstrip(".")
    except ValueError:
        return ""


def is_yahoo_search_url(url: Any) -> bool:
    """Match Yahoo search hosts without blocking unrelated yahoo.com pages."""
    hostname = get_url_hostname(url)
    return hostname == "search.yahoo.com" or hostname.endswith(".search.yahoo.com")


def is_telegram_seo_spam(*values: Any) -> bool:
    """Detect SEO_CARTEL and other high-confidence Telegram SEO promotions."""
    text = " ".join(normalize_filter_text(value) for value in values)
    if not text:
        return False
    if SEO_CARTEL_PATTERN.search(text):
        return True
    return bool(
        TELEGRAM_MARKER_PATTERN.search(text)
        and SEO_SPAM_SIGNAL_PATTERN.search(text)
    )


def classify_prefilter_reason(
    source_url: Any,
    source_title: Any = "",
    anchor: Any = "",
) -> str | None:
    """Return the exclusive reason a row should be removed before URL checks."""
    if is_yahoo_search_url(source_url):
        return "search_yahoo"
    if is_telegram_seo_spam(source_title, anchor, source_url):
        return "telegram_seo_spam"
    return None


def read_file(filepath: str | Path) -> pd.DataFrame:
    """根据文件扩展名读取 Excel 或 CSV 文件，返回 DataFrame。"""
    path = Path(filepath)
    ext = path.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, encoding="utf-8")
    return pd.read_excel(path, engine="openpyxl")


def list_input_files(source_dir: Path) -> list[Path]:
    all_files: list[Path] = []
    for suffix in SUPPORTED_SUFFIXES:
        all_files.extend(Path(item) for item in glob.glob(str(source_dir / f"*{suffix}")))
    return sorted(path for path in all_files if not path.name.startswith(".~"))


def process_backlink_files(
    source_dir: str | Path = SOURCE_DIR,
    output_dir: str | Path = OUTPUT_DIR,
    output_file: str | Path | None = None,
    *,
    processed_log: str | Path = PROCESSED_LOG,
    reprocess: bool = False,
    use_processed_log: bool = True,
    logger: LogCallback | None = print,
    progress_callback: ProgressCallback | None = None,
    read_workers: int = 3,
    should_pause: Callable[[], bool] | None = None,
) -> ProcessorResult:
    """Merge, clean, and de-duplicate backlink export files.

    Web jobs should pass their own source_dir/output_file and set
    use_processed_log=False so each upload batch is independent.
    """
    source_path = Path(source_dir)
    output_path = Path(output_file) if output_file else Path(output_dir) / f"{date.today():%Y-%m-%d}.xlsx"
    processed_log_path = Path(processed_log)

    if reprocess or not use_processed_log:
        if reprocess:
            emit_log(logger, "--reprocess 已启用：忽略已处理记录，所有文件将视为未处理。")
        processed: set[str] = set()
    else:
        processed = load_processed(processed_log_path)

    valid_files = list_input_files(source_path)
    new_files = [
        path
        for path in valid_files
        if reprocess or not use_processed_log or path.name not in processed
    ]

    if not new_files:
        emit_log(logger, "没有新文件需要处理（所有文件均已处理过）")
        emit_progress(progress_callback, 100, "processor 未发现新文件")
        return ProcessorResult(
            output_file="",
            input_files=[],
            failed_files=[],
            rows_read=0,
            rows_output=0,
            duplicate_rows_removed=0,
            filtered_rows=0,
            filtered_search_yahoo_rows=0,
            filtered_telegram_seo_spam_rows=0,
            skipped_processed_files=len(valid_files),
        )

    emit_log(
        logger,
        f"找到 {len(new_files)} 个新文件（已跳过 {len(valid_files) - len(new_files)} 个已处理文件）：",
    )
    for path in new_files:
        emit_log(logger, f"  {path.name}")

    frames_by_path: dict[Path, pd.DataFrame] = {}
    newly_processed: set[str] = set()
    failed_files: list[str] = []
    rows_read = 0
    total_files = len(new_files)

    executor = ThreadPoolExecutor(max_workers=max(1, min(read_workers, total_files)))
    futures = {executor.submit(read_file, path): path for path in new_files}
    completed_files = 0
    try:
        for future in as_completed(futures):
            if should_pause and should_pause():
                for pending in futures:
                    pending.cancel()
                raise ProcessorPaused("任务已在 Excel 文件边界暂停。")
            path = futures[future]
            completed_files += 1
            emit_progress(
                progress_callback,
                completed_files / total_files * 55,
                f"已读取 {completed_files}/{total_files}：{path.name}",
            )
            try:
                df = future.result()
                if df.empty:
                    emit_log(logger, f"  跳过（空文件）：{path.name}")
                    continue
                emit_log(logger, f"  读取 {path.name}：{len(df)} 行")
                rows_read += len(df)
                frames_by_path[path] = df
                newly_processed.add(path.name)
            except Exception as exc:
                failed_files.append(path.name)
                emit_log(logger, f"  读取失败，跳过：{path.name} - {exc}")
    finally:
        executor.shutdown(wait=False, cancel_futures=True)

    frames = [frames_by_path[path] for path in new_files if path in frames_by_path]

    if not frames:
        emit_log(logger, "没有有效数据可处理")
        emit_progress(progress_callback, 100, "processor 没有有效数据")
        return ProcessorResult(
            output_file="",
            input_files=[path.name for path in new_files],
            failed_files=failed_files,
            rows_read=rows_read,
            rows_output=0,
            duplicate_rows_removed=0,
            filtered_rows=0,
            filtered_search_yahoo_rows=0,
            filtered_telegram_seo_spam_rows=0,
            skipped_processed_files=len(valid_files) - len(new_files),
        )

    if should_pause and should_pause():
        raise ProcessorPaused("任务已在合并前暂停。")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        try:
            existing = read_file(output_path)
            if not existing.empty:
                emit_log(logger, f"\n读取已有结果文件：{len(existing)} 行")
                frames.insert(0, existing)
        except Exception as exc:
            emit_log(logger, f"读取已有结果文件失败，将覆盖：{exc}")

    emit_progress(progress_callback, 62, "正在合并文件")
    combined = pd.concat(frames, ignore_index=True)
    emit_log(logger, f"\n合并后总行数：{len(combined)}")

    if "Source url" not in combined.columns:
        emit_log(logger, "错误：找不到 'Source url' 列")
        emit_progress(progress_callback, 100, "processor 缺少 Source url 列")
        raise ValueError("找不到 'Source url' 列")

    columns_by_name = {
        str(column).strip().casefold(): column
        for column in combined.columns
    }
    source_title_column = columns_by_name.get("source title")
    anchor_column = columns_by_name.get("anchor")

    emit_progress(progress_callback, 70, "正在预过滤垃圾外链")
    source_urls = combined["Source url"].tolist()
    source_titles = (
        combined[source_title_column].tolist()
        if source_title_column is not None else [""] * len(combined)
    )
    anchors = (
        combined[anchor_column].tolist()
        if anchor_column is not None else [""] * len(combined)
    )
    reasons: list[str | None] = []
    for start in range(0, len(combined), 5000):
        if should_pause and should_pause():
            raise ProcessorPaused("任务已在预过滤期间暂停。")
        stop = min(start + 5000, len(combined))
        reasons.extend(
            classify_prefilter_reason(source_urls[index], source_titles[index], anchors[index])
            for index in range(start, stop)
        )
        emit_progress(
            progress_callback,
            70 + (stop / max(len(combined), 1)) * 7,
            f"正在预过滤垃圾外链 {stop}/{len(combined)}",
        )
    filter_reasons = pd.Series(reasons, index=combined.index, dtype="object")
    filtered_search_yahoo_rows = int((filter_reasons == "search_yahoo").sum())
    filtered_telegram_seo_spam_rows = int((filter_reasons == "telegram_seo_spam").sum())
    filtered_rows = filtered_search_yahoo_rows + filtered_telegram_seo_spam_rows
    if filtered_rows:
        combined = combined.loc[filter_reasons.isna()].copy()
    emit_log(
        logger,
        "预过滤完成："
        f"共过滤 {filtered_rows} 行"
        f"（Yahoo 搜索页 {filtered_search_yahoo_rows} 行，"
        f"Telegram SEO 垃圾 {filtered_telegram_seo_spam_rows} 行）",
    )
    emit_log(logger, f"预过滤后待去重行数：{len(combined)}")

    if should_pause and should_pause():
        raise ProcessorPaused("任务已在去重前暂停。")

    emit_progress(progress_callback, 78, "正在按注册域名去重")
    combined["_reg_domain"] = combined["Source url"].apply(get_registered_domain)

    if "Page ascore" in combined.columns:
        combined_sorted = combined.sort_values("Page ascore", ascending=False)
    else:
        combined_sorted = combined

    deduped = combined_sorted.drop_duplicates(subset="_reg_domain", keep="first")
    deduped = deduped.drop(columns=["_reg_domain"]).reset_index(drop=True)
    duplicate_rows_removed = len(combined) - len(deduped)

    emit_log(logger, f"去重后行数：{len(deduped)}")
    emit_log(logger, f"共去除重复行：{duplicate_rows_removed}")

    emit_progress(progress_callback, 88, "正在保存 processor 输出")
    deduped.to_excel(output_path, index=False, engine="openpyxl")
    emit_log(logger, f"\n已保存到：{output_path}")

    if use_processed_log:
        save_processed(processed | newly_processed, processed_log_path)
        emit_log(logger, f"已记录处理历史（共 {len(processed | newly_processed)} 个文件）")

    emit_progress(progress_callback, 100, "processor 清洗完成")
    return ProcessorResult(
        output_file=str(output_path),
        input_files=[path.name for path in new_files],
        failed_files=failed_files,
        rows_read=rows_read,
        rows_output=len(deduped),
        duplicate_rows_removed=duplicate_rows_removed,
        filtered_rows=filtered_rows,
        filtered_search_yahoo_rows=filtered_search_yahoo_rows,
        filtered_telegram_seo_spam_rows=filtered_telegram_seo_spam_rows,
        skipped_processed_files=len(valid_files) - len(new_files),
    )


def main(reprocess: bool = False) -> ProcessorResult:
    return process_backlink_files(reprocess=reprocess)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="合并并去重多个 backlinks Excel 文件")
    parser.add_argument("--reprocess", action="store_true", help="忽略已处理记录，重新处理所有文件")
    parser.add_argument("--clear-log", action="store_true", help="删除 processed_files.json 并退出")
    args = parser.parse_args()

    if args.clear_log:
        if PROCESSED_LOG.exists():
            try:
                PROCESSED_LOG.unlink()
                print("已删除 processed_files.json")
            except Exception as exc:
                print(f"删除 processed_files.json 失败：{exc}")
        else:
            print("processed_files.json 不存在，已无需删除")
    else:
        main(reprocess=args.reprocess)
