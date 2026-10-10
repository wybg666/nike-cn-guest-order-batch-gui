import os
import sys
import traceback
from pathlib import Path


def _show_error(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("订单查询 GUI 启动失败", message)
        root.destroy()
    except Exception:
        pass


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    os.chdir(base_dir)
    if str(base_dir) not in sys.path:
        sys.path.insert(0, str(base_dir))

    try:
        from nike_guest_batch_gui import main as gui_main

        gui_main()
        return 0
    except Exception:
        _show_error(traceback.format_exc())
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
