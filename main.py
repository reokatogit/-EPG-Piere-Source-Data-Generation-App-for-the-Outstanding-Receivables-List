"""
main.py - エントリポイント
"""
import sys
import os

# PyInstaller でコンパイルされた場合のパス解決
if getattr(sys, "frozen", False):
    _base = sys._MEIPASS
else:
    _base = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, _base)


def main() -> None:
    from gui import Application
    app = Application()
    app.mainloop()


if __name__ == "__main__":
    main()
