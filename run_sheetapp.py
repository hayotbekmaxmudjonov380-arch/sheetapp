"""PyInstaller uchun kirish nuqtasi."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from sheets_app.__main__ import main

if __name__ == "__main__":
    main()
