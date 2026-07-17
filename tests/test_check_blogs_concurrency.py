from __future__ import annotations

import csv
import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl

from check_blogs.check_blogs import (
    CACHE_RULE_VERSION,
    DetectionResult,
    DomainResultCache,
    detect_url,
    fetch_page,
    process_excel,
)
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
    def test_fetch_uses_split_timeout(self) -> None:
        response = MagicMock()
        response.text = ORDINARY_HTML
        response.url = "https://example.com/final"
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

        def fake_fetch(url: str, logger=None):
            with starts_lock:
                starts.append(time.monotonic())
            return ORDINARY_HTML, url, ""

        cache = DomainResultCache(None)
        with patch("check_blogs.check_blogs.fetch_page", side_effect=fake_fetch):
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
    def test_cache_ttl_and_rule_version_are_enforced(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cache = DomainResultCache(str(Path(temp_dir) / "cache.sqlite3"))
            cache.put(DetectionResult("fresh.com", "无评论功能", "", "", "", "success"))
            self.assertIsNotNone(cache.get("fresh.com"))
            with cache._connect() as connection:
                connection.execute(
                    "UPDATE domain_results SET checked_at = ?, rule_version = ? WHERE domain = ?",
                    (0, CACHE_RULE_VERSION, "fresh.com"),
                )
            self.assertIsNone(cache.get("fresh.com"))
            cache.put(DetectionResult("versioned.com", "无评论功能", "", "", "", "success"))
            with cache._connect() as connection:
                connection.execute(
                    "UPDATE domain_results SET rule_version = ? WHERE domain = ?",
                    ("old-rule", "versioned.com"),
                )
            self.assertIsNone(cache.get("versioned.com"))

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

            def fake_fetch(url: str, logger=None):
                nonlocal active, max_active
                with lock:
                    fetched.append(url)
                    active += 1
                    max_active = max(max_active, active)
                time.sleep(0.1)
                with lock:
                    active -= 1
                return ORDINARY_HTML, url, ""

            started = time.monotonic()
            with patch("check_blogs.check_blogs.fetch_page", side_effect=fake_fetch):
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

            with patch("check_blogs.check_blogs.fetch_page") as second_fetch:
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
            with patch("check_blogs.check_blogs.fetch_page", return_value=(ORDINARY_HTML, "", "")):
                process_excel(str(input_path), output_path=str(output_path), resume=False, logger=None, domain_interval=0)

            checkpoint = output_path.with_name(f"{output_path.stem}_checkpoint.json")
            checkpoint.write_text(json.dumps({"processed_domains": ["domain-0.com"]}), encoding="utf-8")
            output_path.unlink()
            with patch("check_blogs.check_blogs.fetch_page") as resumed_fetch:
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
            with patch("check_blogs.check_blogs.fetch_page", return_value=(html, "https://example.com", "")):
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


if __name__ == "__main__":
    unittest.main()
