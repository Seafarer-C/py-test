#!/usr/bin/env python3
"""自动化测试：覆盖文档第 14/15 节中可离线验证的核心规则。"""

from __future__ import annotations

import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

import xlsx2files


class SafeNameTests(unittest.TestCase):
    def test_illegal_chars(self) -> None:
        self.assertEqual(xlsx2files.safe_name('A/B:C*?', "fallback"), "A_B_C__")

    def test_empty_uses_fallback(self) -> None:
        self.assertEqual(xlsx2files.safe_name("   ", "第2行"), "第2行")

    def test_trailing_dot_and_space(self) -> None:
        self.assertEqual(xlsx2files.safe_name("name. ", "x"), "name")


class UniquePathTests(unittest.TestCase):
    def test_compound_suffix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "archive.tar.gz"
            first.write_text("1", encoding="utf-8")
            second = xlsx2files.unique_path(root / "archive.tar.gz")
            self.assertEqual(second.name, "archive_2.tar.gz")


class UrlExtractionTests(unittest.TestCase):
    def test_printerval_design_urls(self) -> None:
        cell = (
            "https://printerval.com/us/folder-design?"
            "product_id=1&design_urls=/a.png,/b.png"
        )
        urls = xlsx2files.extract_urls(cell)
        self.assertEqual(
            urls,
            [
                "https://assets.printerval.com/a.png",
                "https://assets.printerval.com/b.png",
            ],
        )

    def test_multiple_urls_dedup(self) -> None:
        cell = "https://example.com/a.png https://example.com/a.png，https://example.com/b.png"
        self.assertEqual(
            xlsx2files.extract_urls(cell),
            ["https://example.com/a.png", "https://example.com/b.png"],
        )


class PageImageTests(unittest.TestCase):
    def test_discover_page_images_filters_thumbs(self) -> None:
        html_page = """
        <img src="https://cdn.printerval.com/image/960x960/product.jpg">
        <img src="https://cdn.printerval.com/image/product-preview.png">
        <img src="https://cdn.printerval.com/image/other.webp">
        """
        with mock.patch("xlsx2files.urlopen") as urlopen_mock:
            response = mock.MagicMock()
            response.read.return_value = html_page.encode("utf-8")
            response.__enter__.return_value = response
            urlopen_mock.return_value = response
            result = xlsx2files.discover_page_images(
                "https://printerval.com/us/folder-design?product_id=1"
            )
        self.assertEqual(result, ["https://cdn.printerval.com/image/product-preview.png"])


class ArchiveSecurityTests(unittest.TestCase):
    def test_zip_slip_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            archive = root / "evil.zip"
            with zipfile.ZipFile(archive, "w") as package:
                package.writestr("../escape.txt", "bad")
            with self.assertRaises(RuntimeError):
                xlsx2files.safe_extract_archive(archive)

    def test_zip_extract_keeps_archive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            archive = root / "demo.zip"
            with zipfile.ZipFile(archive, "w") as package:
                package.writestr("inside.txt", "ok")
            extracted = xlsx2files.safe_extract_archive(archive)
            self.assertTrue(archive.exists())
            self.assertTrue((extracted / "inside.txt").exists())


class FailureRetryTests(unittest.TestCase):
    def test_finalize_writes_link_and_renames(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            folder = root / "ORDER1"
            folder.mkdir()
            result = xlsx2files.finalize_order_folder(folder, "ORDER1", ["https://fail.example/a"])
            self.assertTrue(result.name.endswith("_下载失败"))
            link = result / "link.txt"
            self.assertTrue(link.exists())
            text = link.read_text(encoding="utf-8-sig")
            self.assertEqual(text.strip(), "https://fail.example/a")

    def test_prepare_restores_failed_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            failed = root / "ORDER1_下载失败"
            failed.mkdir()
            (failed / "link.txt").write_text("https://x\n", encoding="utf-8-sig")
            restored = xlsx2files.prepare_order_folder(root, "ORDER1")
            self.assertEqual(restored.name, "ORDER1")
            self.assertTrue((restored / "link.txt").exists())

    def test_success_removes_link_txt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            folder = root / "ORDER1"
            folder.mkdir()
            (folder / "link.txt").write_text("old\n", encoding="utf-8-sig")
            result = xlsx2files.finalize_order_folder(folder, "ORDER1", [])
            self.assertEqual(result.name, "ORDER1")
            self.assertFalse((result / "link.txt").exists())


class ExcelParserTests(unittest.TestCase):
    def _make_workbook(self, path: Path, order_no: str | None, with_qr: bool = True) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["no", "二维码", "sku", "品名", "通用下载地址", "效果图下载地址"])
        ws.append([order_no, None, "SKU-1", "衬衫", "https://example.com/a.png", None])
        if with_qr:
            image_path = path.with_name("qr.png")
            PILImage.new("RGB", (16, 16), color=(0, 0, 0)).save(image_path)
            img = XLImage(str(image_path))
            img.anchor = "B2"
            ws.add_image(img)
        wb.save(path)

    def test_skip_downloads_creates_info_and_qr(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            xlsx = root / "sample.xlsx"
            output = root / "out"
            self._make_workbook(xlsx, "MEG001")
            args = type(
                "Args",
                (),
                {
                    "xlsx": xlsx,
                    "output": output,
                    "sheet": None,
                    "limit": 1,
                    "skip_downloads": True,
                    "browser_fallback": "off",
                    "browser_profile": root / "profile",
                    "cancel_event": None,
                },
            )()
            code = xlsx2files.process(args)
            self.assertEqual(code, 0)
            folder = output / "MEG001"
            info = (folder / "信息.txt").read_text(encoding="utf-8-sig")
            self.assertIn("no: MEG001", info)
            self.assertIn("sku: SKU-1", info)
            self.assertNotIn("下载地址", info)
            self.assertNotIn("二维码:", info)
            self.assertTrue((folder / "二维码.png").exists())

    def test_empty_first_column_uses_row_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            xlsx = root / "sample.xlsx"
            output = root / "out"
            self._make_workbook(xlsx, None, with_qr=False)
            args = type(
                "Args",
                (),
                {
                    "xlsx": xlsx,
                    "output": output,
                    "sheet": None,
                    "limit": 1,
                    "skip_downloads": True,
                    "browser_fallback": "off",
                    "browser_profile": root / "profile",
                    "cancel_event": None,
                },
            )()
            code = xlsx2files.process(args)
            self.assertEqual(code, 1)  # 缺少二维码计入失败
            self.assertTrue((output / "第2行").exists())

    def test_failed_download_writes_link_txt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            xlsx = root / "sample.xlsx"
            output = root / "out"
            self._make_workbook(xlsx, "MEGFAIL", with_qr=True)
            args = type(
                "Args",
                (),
                {
                    "xlsx": xlsx,
                    "output": output,
                    "sheet": None,
                    "limit": 1,
                    "skip_downloads": False,
                    "browser_fallback": "off",
                    "browser_profile": root / "profile",
                    "cancel_event": None,
                },
            )()

            def boom(*_args, **_kwargs):
                raise RuntimeError("simulated timeout")

            with mock.patch.object(xlsx2files, "download_url", side_effect=boom):
                code = xlsx2files.process(args)
            self.assertEqual(code, 1)
            failed = output / "MEGFAIL_下载失败"
            self.assertTrue(failed.is_dir())
            link = (failed / "link.txt").read_text(encoding="utf-8-sig").strip()
            self.assertEqual(link, "https://example.com/a.png")

    def test_cancel_returns_130(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            xlsx = root / "sample.xlsx"
            output = root / "out"
            self._make_workbook(xlsx, "MEGSTOP", with_qr=True)
            cancel = threading.Event()
            cancel.set()
            args = type(
                "Args",
                (),
                {
                    "xlsx": xlsx,
                    "output": output,
                    "sheet": None,
                    "limit": 1,
                    "skip_downloads": True,
                    "browser_fallback": "off",
                    "browser_profile": root / "profile",
                    "cancel_event": cancel,
                },
            )()
            self.assertEqual(xlsx2files.process(args), 130)

    def test_legacy_dirs_flattened(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            xlsx = root / "sample.xlsx"
            output = root / "out"
            self._make_workbook(xlsx, "MEGFLAT", with_qr=True)
            order = output / "MEGFLAT"
            legacy = order / "通用"
            legacy.mkdir(parents=True)
            (legacy / "old.png").write_bytes(b"png")
            args = type(
                "Args",
                (),
                {
                    "xlsx": xlsx,
                    "output": output,
                    "sheet": None,
                    "limit": 1,
                    "skip_downloads": True,
                    "browser_fallback": "off",
                    "browser_profile": root / "profile",
                    "cancel_event": None,
                },
            )()
            xlsx2files.process(args)
            self.assertTrue((order / "old.png").exists())
            self.assertFalse(legacy.exists())


class FlattenLegacyTests(unittest.TestCase):
    def test_no_overwrite_on_conflict(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "existing.png").write_bytes(b"1")
            legacy = root / "效果图"
            legacy.mkdir()
            (legacy / "existing.png").write_bytes(b"2")
            moved = xlsx2files.flatten_legacy_download_dirs(root, {"效果图"})
            self.assertEqual(moved, 1)
            self.assertTrue((root / "existing_2.png").exists())


if __name__ == "__main__":
    unittest.main()
