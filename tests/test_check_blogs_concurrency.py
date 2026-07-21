from __future__ import annotations

import csv
import asyncio
import json
import os
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pandas as pd
from bs4 import BeautifulSoup as RealBeautifulSoup

from check_blogs.check_blogs import (
    CACHE_RULE_VERSION,
    DetectionResult,
    DomainResultCache,
    analyze_html_once,
    detect_url,
    fetch_page_async,
    fetch_page,
    process_excel,
)
from processor.process_backlinks import ProcessorPaused, process_backlink_files
from web.app import AppConfig, JobManager


ORDINARY_HTML = "<html><head><title>Article</title></head><body>ordinary</body></html>"


def write_input(path: Path, count: int) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Page address", "Source title", "Source url"])
    for index in range(count):
        sheet.append([f"page-{index}", f"title-{index}", f"https://domain-{index}.com/post"])
    workbook.save(path)


class NetworkConfigurationTests(unittest.TestCase):
    def test_async_fetch_uses_hard_total_timeout_and_streaming(self) -> None:
        captured: dict[str, object] = {}

        class Content:
            async def iter_chunked(self, size):
                yield ORDINARY_HTML.encode()

        class Response:
            url = "https://example.com/final"
            charset = "utf-8"
            content = Content()
            def raise_for_status(self):
                return None
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                return None

        class Session:
            def get(self, url, **kwargs):
                captured.update(kwargs)
                return Response()

        html, final_url, error = asyncio.run(
            fetch_page_async(Session(), "https://example.com", logger=None)
        )
        self.assertEqual(html, ORDINARY_HTML)
        self.assertEqual(final_url, "https://example.com/final")
        self.assertEqual(error, "")
        self.assertEqual(captured["timeout"].total, 30.0)
        self.assertEqual(captured["max_redirects"], 5)

    def test_fetch_uses_split_timeout(self) -> None:
        response = MagicMock()
        response.text = ORDINARY_HTML
        response.url = "https://example.com/final"
        response.encoding = "utf-8"
        response.iter_content.return_value = [ORDINARY_HTML.encode()]
        session = MagicMock()
        session.get.return_value = response
        with patch("check_blogs.check_blogs._get_session", return_value=session):
            html, final_url, error = fetch_page("https://example.com", logger=None)
        self.assertEqual(html, ORDINARY_HTML)
        self.assertEqual(final_url, "https://example.com/final")
        self.assertEqual(error, "")
        self.assertEqual(session.get.call_args.kwargs["timeout"], (5.0, 20.0))

    def test_same_domain_requests_are_rate_limited(self) -> None:
        starts: list[float] = []
        starts_lock = threading.Lock()

        async def fake_fetch(session, url: str, logger=None):
            with starts_lock:
                starts.append(time.monotonic())
            return ORDINARY_HTML, url, ""

        cache = DomainResultCache(None)
        with patch("check_blogs.check_blogs.fetch_page_async", side_effect=fake_fetch):
            threads = [
                threading.Thread(
                    target=detect_url,
                    args=(f"https://same.example/{index}", "same.example", cache, None, 0.04),
                )
                for index in range(2)
            ]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
        self.assertEqual(len(starts), 2)
        self.assertGreaterEqual(starts[1] - starts[0], 0.035)


class ConcurrentPipelineTests(unittest.TestCase):
    def test_html_is_parsed_only_once(self) -> None:
        calls = 0

        def counting_soup(*args, **kwargs):
            nonlocal calls
            calls += 1
            return RealBeautifulSoup(*args, **kwargs)

        with patch("check_blogs.check_blogs.BeautifulSoup", side_effect=counting_soup):
            result = analyze_html_once(ORDINARY_HTML, "https://example.com", "https://example.com", "example.com")
        self.assertEqual(result.label, "无评论功能")
        self.assertEqual(calls, 1)

    def test_recent_history_is_imported_without_overwriting_newer_cache(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            result_dir = root / "jobs" / "job-1" / "result"
            result_dir.mkdir(parents=True)
            csv_path = result_dir / "history.csv"
            with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle)
                writer.writerow(["顶级域名", "评论表单检测", "特殊标注", "最新评论时间"])
                writer.writerow(["history.example", "博客网站", "", "2026-07-20 10:00"])
            cache = DomainResultCache(str(root / "cache.sqlite3"))
            imported = cache.import_history(root / "jobs")
            self.assertEqual(imported, 1)
            self.assertEqual(cache.get("history.example").label, "博客网站")
            cache.put(DetectionResult("history.example", "评论网站", "", "", "", "success"))
            cache.import_history(root / "jobs")
            self.assertEqual(cache.get("history.example").label, "评论网站")

    def test_slow_first_input_does_not_block_completed_batch_progress(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "input.xlsx"
            output_path = root / "result.xlsx"
            write_input(input_path, 26)
            updates: list[tuple[float, int]] = []
            started = time.monotonic()

            async def fake_fetch(session, url: str, logger=None):
                await asyncio.sleep(0.45 if "domain-0.com" in url else 0.01)
                return ORDINARY_HTML, url, ""

            with patch("check_blogs.check_blogs.fetch_page_async", side_effect=fake_fetch):
                process_excel(
                    str(input_path), output_path=str(output_path), resume=False,
                    logger=None, max_workers=16, checkpoint_batch_size=25,
                    domain_interval=0,
                    stats_callback=lambda stats: updates.append(
                        (time.monotonic() - started, stats["processed_rows"])
                    ),
                )
            self.assertTrue(updates)
            self.assertEqual(updates[0][1], 25)
            self.assertLess(updates[0][0], 0.35)

    def test_cache_ttl_and_rule_version_are_enforced(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cache = DomainResultCache(str(Path(temp_dir) / "cache.sqlite3"))
            cache.put(DetectionResult("fresh.com", "无评论功能", "", "", "", "success"))
            self.assertIsNotNone(cache.get("fresh.com"))
            with cache._connection() as connection:
                connection.execute(
                    "UPDATE domain_results SET checked_at = ?, rule_version = ? WHERE domain = ?",
                    (0, CACHE_RULE_VERSION, "fresh.com"),
                )
            self.assertIsNone(cache.get("fresh.com"))
            cache.put(DetectionResult("versioned.com", "无评论功能", "", "", "", "success"))
            with cache._connection() as connection:
                connection.execute(
                    "UPDATE domain_results SET rule_version = ? WHERE domain = ?",
                    ("old-rule", "versioned.com"),
                )
            self.assertIsNone(cache.get("versioned.com"))

    def test_repeated_cache_reads_do_not_leak_file_descriptors(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cache = DomainResultCache(str(Path(temp_dir) / "cache.sqlite3"))
            cache.put(DetectionResult("cached.com", "无评论功能", "", "", "", "success"))
            fd_root = Path("/dev/fd")
            before = len(os.listdir(fd_root))
            for _ in range(500):
                self.assertIsNotNone(cache.get("cached.com"))
            after = len(os.listdir(fd_root))
            self.assertLessEqual(after - before, 4)

    def test_parallel_fetch_preserves_order_and_reuses_cache(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "input.xlsx"
            output_path = root / "result.xlsx"
            second_output = root / "second.xlsx"
            cache_path = root / "cache.sqlite3"
            write_input(input_path, 32)
            fetched: list[str] = []
            active = 0
            max_active = 0
            lock = threading.Lock()

            async def fake_fetch(session, url: str, logger=None):
                nonlocal active, max_active
                with lock:
                    fetched.append(url)
                    active += 1
                    max_active = max(max_active, active)
                await asyncio.sleep(0.1)
                with lock:
                    active -= 1
                return ORDINARY_HTML, url, ""

            started = time.monotonic()
            with patch("check_blogs.check_blogs.fetch_page_async", side_effect=fake_fetch):
                first = process_excel(
                    str(input_path), output_path=str(output_path), resume=False,
                    logger=None, max_workers=16, checkpoint_batch_size=25,
                    cache_path=str(cache_path), domain_interval=0,
                )
            elapsed = time.monotonic() - started

            self.assertGreaterEqual(max_active, 8)
            self.assertLess(elapsed, 0.4)  # serial baseline is at least 3.2s (8x+ faster)
            self.assertEqual(first.network_checked_rows, 32)
            self.assertEqual(first.cache_hit_rows, 0)
            with open(output_path.with_suffix(".csv"), newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.reader(handle))[1:]
            self.assertEqual([row[0] for row in rows], [f"page-{index}" for index in range(32)])

            with patch("check_blogs.check_blogs.fetch_page_async") as second_fetch:
                second = process_excel(
                    str(input_path), output_path=str(second_output), resume=False,
                    logger=None, max_workers=16, cache_path=str(cache_path), domain_interval=0,
                )
            second_fetch.assert_not_called()
            self.assertEqual(second.cache_hit_rows, 32)
            self.assertEqual(second.network_checked_rows, 0)

    def test_resume_combines_csv_and_checkpoint_without_refetch(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "input.xlsx"
            output_path = root / "result.xlsx"
            write_input(input_path, 3)
            async def ordinary_fetch(session, url: str, logger=None):
                return ORDINARY_HTML, url, ""

            with patch("check_blogs.check_blogs.fetch_page_async", side_effect=ordinary_fetch):
                process_excel(str(input_path), output_path=str(output_path), resume=False, logger=None, domain_interval=0)

            checkpoint = output_path.with_name(f"{output_path.stem}_checkpoint.json")
            checkpoint.write_text(json.dumps({"processed_domains": ["domain-0.com"]}), encoding="utf-8")
            output_path.unlink()
            with patch("check_blogs.check_blogs.fetch_page_async") as resumed_fetch:
                resumed = process_excel(str(input_path), output_path=str(output_path), resume=True, logger=None, domain_interval=0)
            resumed_fetch.assert_not_called()
            self.assertEqual(resumed.resumed_rows, 3)
            self.assertTrue(output_path.exists())


class GoogleLoginTests(unittest.TestCase):
    def test_google_login_is_only_applied_to_comment_pages(self) -> None:
        google_signal = '<script src="https://accounts.google.com/gsi/client"></script>'
        partial_comment = f"<html><body>{google_signal}<p>Leave a comment</p></body></html>"
        ordinary_login = f"<html><body>{google_signal}<p>Account login</p></body></html>"
        complete_form = f"""<html><body>{google_signal}<form>
            <textarea name="comment"></textarea><input name="name"><input name="email">
            <input name="website"><button>Post Comment</button></form></body></html>"""
        cache = DomainResultCache(None)

        for html, expected in (
            (partial_comment, "评论网站 · 谷歌登录"),
            (ordinary_login, "无评论功能"),
            (complete_form, "博客网站 · 谷歌登录"),
        ):
            async def html_fetch(session, url: str, logger=None, html=html):
                return html, "https://example.com", ""

            with patch("check_blogs.check_blogs.fetch_page_async", side_effect=html_fetch):
                result = detect_url("https://example.com", f"{expected}.example", cache, None, 0)
            self.assertEqual(result.label, expected)
            if "谷歌登录" in expected:
                self.assertEqual(result.flags, "需登录（谷歌登录）")


class RecoveryTests(unittest.TestCase):
    def test_non_terminal_job_with_processor_output_is_requeued(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            job_id = "20260717-000000-recovery"
            job_dir = root / "jobs" / job_id
            (job_dir / "processor").mkdir(parents=True)
            (job_dir / "processor" / "processed_backlinks.xlsx").touch()
            metadata = {
                "job_id": job_id, "status": "running", "stage": "check_blogs",
                "progress": 62, "created_at": "2026-07-17T00:00:00",
                "updated_at": "2026-07-17T00:01:00", "files": [], "logs": [],
                "stats": {}, "error": None, "download_path": None,
            }
            (job_dir / "metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
            manager = JobManager(AppConfig(storage_root=root, worker_count=1))
            record = manager.get_job(job_id)
            self.assertEqual(record.status, "queued")
            self.assertIn(job_id, manager.recoverable_job_ids)

    def test_pausing_job_becomes_paused_without_requeueing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            job_id = "20260721-000000-pausing"
            job_dir = root / "jobs" / job_id
            (job_dir / "uploads").mkdir(parents=True)
            (job_dir / "uploads" / "input.xlsx").touch()
            metadata = {
                "job_id": job_id, "status": "pausing", "stage": "processor",
                "progress": 12, "created_at": "2026-07-21T00:00:00",
                "updated_at": "2026-07-21T00:01:00", "files": [], "logs": [],
                "stats": {}, "error": None, "download_path": None,
            }
            (job_dir / "metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
            manager = JobManager(AppConfig(storage_root=root, worker_count=1))
            record = manager.get_job(job_id)
            self.assertEqual(record.status, "paused")
            self.assertNotIn(job_id, manager.recoverable_job_ids)

    def test_failed_job_with_processor_output_can_resume(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            job_id = "20260721-000000-failed"
            job_dir = root / "jobs" / job_id
            (job_dir / "processor").mkdir(parents=True)
            (job_dir / "processor" / "processed_backlinks.xlsx").touch()
            metadata = {
                "job_id": job_id, "status": "failed", "stage": "failed",
                "progress": 52.8, "created_at": "2026-07-21T00:00:00",
                "updated_at": "2026-07-21T00:01:00", "files": [], "logs": [],
                "stats": {}, "error": "Too many open files", "download_path": None,
            }
            (job_dir / "metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
            manager = JobManager(AppConfig(storage_root=root, worker_count=1))
            with patch.object(manager, "enqueue") as enqueue:
                record = manager.resume_job(job_id)
            self.assertEqual(record.status, "queued")
            self.assertIsNone(record.error)
            enqueue.assert_called_once_with(job_id)


class ProcessorConcurrencyTests(unittest.TestCase):
    def test_files_are_read_in_parallel_but_merged_deterministically(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            uploads = root / "uploads"
            uploads.mkdir()
            for name in ("a.xlsx", "b.xlsx", "c.xlsx"):
                (uploads / name).touch()

            def fake_read(path):
                time.sleep(0.1)
                return pd.DataFrame({"Source url": [f"https://{Path(path).stem}.example/post"]})

            started = time.monotonic()
            with patch("processor.process_backlinks.read_file", side_effect=fake_read):
                result = process_backlink_files(
                    source_dir=uploads, output_file=root / "result.xlsx",
                    use_processed_log=False, logger=None, read_workers=3,
                )
            self.assertLess(time.monotonic() - started, 0.25)
            self.assertEqual(result.rows_read, 3)
            output = pd.read_excel(root / "result.xlsx")
            self.assertEqual(output["Source url"].tolist(), [
                "https://a.example/post", "https://b.example/post", "https://c.example/post",
            ])

    def test_processor_pause_stops_at_file_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            uploads = root / "uploads"
            uploads.mkdir()
            for name in ("a.xlsx", "b.xlsx"):
                (uploads / name).touch()
            pause = threading.Event()

            def fake_read(path):
                pause.set()
                return pd.DataFrame({"Source url": [f"https://{Path(path).stem}.example"]})

            with patch("processor.process_backlinks.read_file", side_effect=fake_read):
                with self.assertRaises(ProcessorPaused):
                    process_backlink_files(
                        source_dir=uploads, output_file=root / "result.xlsx",
                        use_processed_log=False, logger=None, read_workers=2,
                        should_pause=pause.is_set,
                    )


if __name__ == "__main__":
    unittest.main()
