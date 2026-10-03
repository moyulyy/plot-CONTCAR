# -*- coding: utf-8 -*-
"""无控制台启动器（用 pythonw.exe 运行）。

pyw / pythonw 启动时没有控制台窗口，因此这里把任何启动异常
（例如缺少 PySide6）用 Windows 消息框弹出来，避免“点了没反应”。
"""
import sys
import traceback
from pathlib import Path

BASE = Path(__file__).resolve().parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))


def _show_error(text: str) -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, "CONTCAR → PPT 启动失败", 0x10)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        import vesta_gui
        vesta_gui.main()
    except Exception:
        _show_error(traceback.format_exc())
