import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from crawler import ScheduleCrawler


class BrowserPageTest(unittest.TestCase):
    def test_launch_reuses_existing_page_in_persistent_context(self):
        existing_page = MagicMock()
        context = MagicMock()
        context.pages = [existing_page]

        playwright = MagicMock()
        playwright.chromium.launch_persistent_context.return_value = context
        playwright_manager = MagicMock()
        playwright_manager.start.return_value = playwright

        crawler = ScheduleCrawler()
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch("crawler.sync_playwright", return_value=playwright_manager):
                with patch("crawler.config.USER_DATA_DIR", Path(temp_dir)):
                    page = crawler.launch()

        self.assertIs(page, existing_page)
        context.new_page.assert_not_called()
        existing_page.goto.assert_called_once()


if __name__ == "__main__":
    unittest.main()
