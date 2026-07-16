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
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Any, Callable

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

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float, str], None]


@dataclass
class ProcessorResult:
    output_file: str
    input_files: list[str]
    failed_files: list[str]
    rows_read: int
    rows_output: int
    duplicate_rows_removed: int
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
            skipped_processed_files=len(valid_files),
        )

    emit_log(
        logger,
        f"找到 {len(new_files)} 个新文件（已跳过 {len(valid_files) - len(new_files)} 个已处理文件）：",
    )
    for path in new_files:
        emit_log(logger, f"  {path.name}")

    frames: list[pd.DataFrame] = []
    newly_processed: set[str] = set()
    failed_files: list[str] = []
    rows_read = 0
    total_files = len(new_files)

    for index, path in enumerate(new_files, start=1):
        emit_progress(
            progress_callback,
            (index - 1) / total_files * 55,
            f"正在读取 {path.name}",
        )
        try:
            df = read_file(path)
            if df.empty:
                emit_log(logger, f"  跳过（空文件）：{path.name}")
                continue
            emit_log(logger, f"  读取 {path.name}：{len(df)} 行")
            rows_read += len(df)
            frames.append(df)
            newly_processed.add(path.name)
        except Exception as exc:
            failed_files.append(path.name)
            emit_log(logger, f"  读取失败，跳过：{path.name} - {exc}")

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
            skipped_processed_files=len(valid_files) - len(new_files),
        )

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

    emit_progress(progress_callback, 74, "正在按注册域名去重")
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
