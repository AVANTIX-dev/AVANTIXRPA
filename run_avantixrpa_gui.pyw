from avantixrpa.ui.main_window import main
import traceback
import tkinter as tk
from tkinter import messagebox

if __name__ == "__main__":
    try:
        main()
    except Exception:
        # ログに出てる前提だけど、一応画面でも知らせる
        err = traceback.format_exc()
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "AVANTIXRPA 起動エラー",
            "AVANTIXRPA の起動中にエラーが発生しました。\n"
            "詳しくは logs/avantixrpa.log を確認してください。\n\n"
            f"{err}",
        )
        root.destroy()
