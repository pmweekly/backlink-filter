from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.parse import urlparse

import pandas as pd

from check_blogs.check_blogs import process_excel
from processor.process_backlinks import (
    classify_prefilter_reason,
    is_telegram_seo_spam,
    is_yahoo_search_url,
    process_backlink_files,
)


class PrefilterRuleTests(unittest.TestCase):
    def test_yahoo_search_host_matching_is_narrow(self) -> None:
        self.assertTrue(is_yahoo_search_url("https://search.yahoo.com/search?p=test"))
        self.assertTrue(is_yahoo_search_url("https://dk.search.yahoo.com/mobile/s?p=test"))
        self.assertFalse(is_yahoo_search_url("https://yahoo.com/"))
        self.assertFalse(is_yahoo_search_url("https://finance.yahoo.com/"))
        self.assertFalse(is_yahoo_search_url("https://notsearch.yahoo.com/"))

    def test_seo_cartel_is_detected_in_each_relevant_field(self) -> None:
        self.assertTrue(is_telegram_seo_spam("@SEO_CARTEL IN TELEGRAM"))
        self.assertTrue(is_telegram_seo_spam("", "https://t.me/SEO_CARTEL"))
        self.assertTrue(
            is_telegram_seo_spam(
                "",
                "",
                "https://example.com/seo_cartel-in-telegram-seo-backlinks/",
            )
        )

    def test_generic_telegram_seo_spam_requires_both_signal_groups(self) -> None:
        self.assertTrue(
            is_telegram_seo_spam(
                "TG @BHS_LINKS - BEST SEO BACKLINKS - https://t.me/bhs_links"
            )
        )
        self.assertTrue(
            is_telegram_seo_spam(
                "TELEGRAM @SALESOVEN | ACCESS TO HACKED SITES FOR SEO"
            )
        )
        self.assertFalse(is_telegram_seo_spam("Join our Telegram community"))
        self.assertFalse(is_telegram_seo_spam("A practical SEO backlink guide"))


class ProcessorPrefilterIntegrationTests(unittest.TestCase):
    def _write_input(self, source_dir: Path) -> None:
        rows = [
            {
                "Page ascore": 100,
                "Source title": "@SEO_CARTEL IN TELEGRAM – SEO BACKLINKS",
                "Source url": "https://example.com/seo-cartel-post/",
                "Target url": "https://target.example/",
                "Anchor": "Bulk link posting",
            },
            {
                "Page ascore": 10,
                "Source title": "Useful article",
                "Source url": "https://example.com/useful-article/",
                "Target url": "https://target.example/",
                "Anchor": "Useful source",
            },
            {
                "Page ascore": 24,
                "Source title": "Yahoo Search Results",
                "Source url": "https://dk.search.yahoo.com/mobile/s?p=test",
                "Target url": "https://target.example/",
                "Anchor": "Search result",
            },
            {
                "Page ascore": 20,
                "Source title": "Yahoo Finance",
                "Source url": "https://finance.yahoo.com/markets/",
                "Target url": "https://target.example/",
                "Anchor": "Market news",
            },
            {
                "Page ascore": 15,
                "Source title": "Join our Telegram community",
                "Source url": "https://community.example/news/",
                "Target url": "https://target.example/",
                "Anchor": "Community news",
            },
            {
                "Page ascore": 12,
                "Source title": "Normal archive title",
                "Source url": "https://spam.example/archive/",
                "Target url": "https://target.example/",
                "Anchor": "TG @LINKS_DEALER | EFFECTIVE SEO BACKLINKS",
            },
        ]
        pd.DataFrame(rows).to_excel(source_dir / "backlinks.xlsx", index=False)

    def test_prefilter_runs_before_domain_deduplication_and_reports_counts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_dir = root / "uploads"
            output_file = root / "processed.xlsx"
            source_dir.mkdir()
            self._write_input(source_dir)

            result = process_backlink_files(
                source_dir=source_dir,
                output_file=output_file,
                use_processed_log=False,
                logger=None,
            )
            output = pd.read_excel(output_file)
            urls = set(output["Source url"])

            self.assertEqual(result.rows_read, 6)
            self.assertEqual(result.filtered_rows, 3)
            self.assertEqual(result.filtered_search_yahoo_rows, 1)
            self.assertEqual(result.filtered_telegram_seo_spam_rows, 2)
            self.assertEqual(result.duplicate_rows_removed, 0)
            self.assertEqual(result.rows_output, 3)
            self.assertIn("https://example.com/useful-article/", urls)
            self.assertIn("https://finance.yahoo.com/markets/", urls)
            self.assertIn("https://community.example/news/", urls)
            self.assertNotIn("https://example.com/seo-cartel-post/", urls)

    def test_filtered_urls_never_reach_check_blogs_fetch(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_dir = root / "uploads"
            processor_output = root / "processed.xlsx"
            final_output = root / "result.xlsx"
            source_dir.mkdir()
            self._write_input(source_dir)

            process_backlink_files(
                source_dir=source_dir,
                output_file=processor_output,
                use_processed_log=False,
                logger=None,
            )

            fetched_urls: list[str] = []

            def fake_fetch(url: str, logger=None):
                fetched_urls.append(url)
                return "<html><body><p>ordinary page</p></body></html>", url, ""

            with (
                patch("check_blogs.check_blogs.fetch_page", side_effect=fake_fetch),
                patch(
                    "check_blogs.check_blogs.get_top_domain",
                    side_effect=lambda url: urlparse(url).hostname or "",
                ),
            ):
                process_excel(
                    str(processor_output),
                    output_path=str(final_output),
                    resume=False,
                    logger=None,
                    delay_between=0,
                )

            self.assertEqual(
                set(fetched_urls),
                {
                    "https://example.com/useful-article/",
                    "https://finance.yahoo.com/markets/",
                    "https://community.example/news/",
                },
            )
            for url in fetched_urls:
                self.assertIsNone(classify_prefilter_reason(url))


if __name__ == "__main__":
    unittest.main()
