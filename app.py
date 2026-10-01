"""Vercel / WSGI kirish nuqtasi — Flask ilovasini ochadi."""
import os

if os.environ.get("VERCEL"):
    os.environ.setdefault("SHEETAPP_DB", "/tmp/web_workbook.db")

from sheets_app.web import app  # noqa: E402,F401

__all__ = ["app"]
