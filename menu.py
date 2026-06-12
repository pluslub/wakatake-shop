import tkinter as tk
from tkinter import font as tkfont
import csv_to_excel
import delivery_note
import order_form

BG = "#f5f5f5"
ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
TEXT = "#1e293b"
TEXT_SUB = "#64748b"
CARD_BG = "#ffffff"
CARD_BORDER = "#e2e8f0"


TOOLS = [
    {
        "label": "部署金額まとめ表を作成",
        "desc":  "CSVを読み込んで部署別Excelを生成します",
        "func":  csv_to_excel.run,
    },
    {
        "label": "納品書を作成",
        "desc":  "納品書用CSVを読み込んでダイキン用納品書Excelを生成します",
        "func":  delivery_note.run,
    },
    {
        "label": "注文書を作成",
        "desc":  "納品書用CSVを読み込んで各注文書テンプレートに個数を自動入力します",
        "func":  order_form.run,
    },
    # ここに追加していく
]


class MenuApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("わかたけショップ 管理ツール")
        self.configure(bg=BG)
        self.resizable(False, False)
        self._build_ui()
        self._center()

    def _build_ui(self):
        # ヘッダー
        header = tk.Frame(self, bg=ACCENT, padx=30, pady=18)
        header.pack(fill="x")
        tk.Label(
            header,
            text="わかたけショップ",
            bg=ACCENT, fg="#ffffff",
            font=("游ゴシック", 11),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="管理ツール",
            bg=ACCENT, fg="#ffffff",
            font=("游ゴシック", 20, "bold"),
        ).pack(anchor="w")

        # ツール一覧
        body = tk.Frame(self, bg=BG, padx=24, pady=20)
        body.pack(fill="both", expand=True)

        for tool in TOOLS:
            self._add_tool_card(body, tool["label"], tool["desc"], tool["func"])

    def _add_tool_card(self, parent, label, desc, func):
        card = tk.Frame(parent, bg=CARD_BG, relief="flat", bd=0,
                        highlightthickness=1, highlightbackground=CARD_BORDER)
        card.pack(fill="x", pady=6)

        inner = tk.Frame(card, bg=CARD_BG, padx=16, pady=14)
        inner.pack(fill="x")

        left = tk.Frame(inner, bg=CARD_BG)
        left.pack(side="left", fill="both", expand=True)

        tk.Label(left, text=label, bg=CARD_BG, fg=TEXT,
                 font=("游ゴシック", 13, "bold"), anchor="w").pack(fill="x")
        tk.Label(left, text=desc, bg=CARD_BG, fg=TEXT_SUB,
                 font=("游ゴシック", 9), anchor="w").pack(fill="x")

        btn = tk.Button(
            inner, text="実行",
            bg=ACCENT, fg="#ffffff", relief="flat",
            activebackground=ACCENT_HOVER, activeforeground="#ffffff",
            font=("游ゴシック", 10, "bold"),
            padx=18, pady=6, cursor="hand2",
            command=func,
        )
        btn.pack(side="right", padx=(12, 0))

    def _center(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


if __name__ == "__main__":
    app = MenuApp()
    app.mainloop()
