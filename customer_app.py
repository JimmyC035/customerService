import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import os
from datetime import datetime

# 主資料庫檔案
MASTER_FILE = 'master_data.xlsx'

# 色彩主題
COLORS = {
    "bg": "#eef2f7",          # 主背景 淡藍灰
    "card": "#ffffff",        # 卡片背景
    "primary": "#4a6fa5",     # 主色 藍
    "primary_hover": "#3b5d8c",
    "success": "#2e7d32",     # 綠
    "text": "#2c3e50",        # 主文字
    "text_light": "#7f8c8d",  # 次要文字
    "border": "#dce3eb",      # 邊框
    "input_bg": "#f8fafc",    # 輸入框背景
    "table_header": "#4a6fa5",
    "table_row_alt": "#f5f8fb",
    "table_selected": "#4a6fa5",
}


class CustomerSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("客戶帳務管理系統")
        self.root.geometry("1400x850")
        self.root.configure(bg=COLORS["bg"])

        # 定義 Excel 的欄位順序 (對應帳表)
        self.cols = [
            "日期", "訂購人", "電話", "地址", "贈品", "品項", "備註", "尺寸",
            "數量", "單價", "總價", "折扣", "特殊折扣", "實拿",
            "成本(品項+贈品)", "利潤", "付款方式", "是否已付"
        ]

        self.load_master_data()
        self.setup_styles()
        self.build_ui()
        self.display(self.df)

    # ─── 風格設定 ───
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Treeview 表格
        style.configure("Custom.Treeview",
                        background="white",
                        fieldbackground="white",
                        foreground=COLORS["text"],
                        rowheight=32,
                        font=("Arial", 11))
        style.configure("Custom.Treeview.Heading",
                        background=COLORS["table_header"],
                        foreground="white",
                        font=("Arial", 11, "bold"),
                        relief="flat")
        style.map("Custom.Treeview.Heading",
                  background=[('active', COLORS["primary_hover"])])
        style.map("Custom.Treeview",
                  background=[('selected', COLORS["table_selected"])],
                  foreground=[('selected', 'white')])

        # 滾動條
        style.configure("Custom.Vertical.TScrollbar", troughcolor=COLORS["bg"], borderwidth=0)
        style.configure("Custom.Horizontal.TScrollbar", troughcolor=COLORS["bg"], borderwidth=0)

    # ─── 建構 UI ───
    def build_ui(self):
        # === 標題列 ===
        header = tk.Frame(self.root, bg=COLORS["primary"], height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="客戶帳務管理系統",
                 bg=COLORS["primary"], fg="white",
                 font=("Arial", 18, "bold")).pack(side="left", padx=24)

        # === 搜尋 + 功能列 ===
        toolbar = tk.Frame(self.root, bg=COLORS["bg"])
        toolbar.pack(fill="x", padx=24, pady=(16, 8))

        # 搜尋框
        search_box = tk.Frame(toolbar, bg=COLORS["card"],
                              highlightbackground=COLORS["border"], highlightthickness=1)
        search_box.pack(side="left", fill="x", expand=True, ipady=4)

        tk.Label(search_box, text=" 🔍", bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 13)).pack(side="left", padx=(8, 0))
        self.search_ent = tk.Entry(search_box, width=40,
                                   bg=COLORS["card"], fg=COLORS["text"],
                                   insertbackground=COLORS["text"],
                                   font=("Arial", 12), relief="flat",
                                   highlightthickness=0)
        self.search_ent.pack(side="left", padx=8, pady=6, fill="x", expand=True)
        self.search_ent.insert(0, "搜尋姓名、電話或地址...")
        self.search_ent.config(fg=COLORS["text_light"])
        self.search_ent.bind("<FocusIn>", self._search_focus_in)
        self.search_ent.bind("<FocusOut>", self._search_focus_out)
        self.search_ent.bind("<KeyRelease>", lambda e: self.search())

        # 按鈕區
        btn_frame = tk.Frame(toolbar, bg=COLORS["bg"])
        btn_frame.pack(side="right", padx=(16, 0))

        self._make_btn(btn_frame, "顯示全部", COLORS["card"], COLORS["text"],
                       lambda: self.display(self.df)).pack(side="left", padx=4)
        self._make_btn(btn_frame, "📥 匯入帳表", "#e3f2fd", COLORS["primary"],
                       self.import_data).pack(side="left", padx=4)

        # === 新增紀錄區 ===
        self.add_visible = tk.BooleanVar(value=True)
        toggle_frame = tk.Frame(self.root, bg=COLORS["bg"])
        toggle_frame.pack(fill="x", padx=24)
        self.toggle_btn = tk.Label(toggle_frame, text="▼ 快速新增紀錄",
                                   bg=COLORS["bg"], fg=COLORS["primary"],
                                   font=("Arial", 12, "bold"), cursor="hand2")
        self.toggle_btn.pack(anchor="w")
        self.toggle_btn.bind("<Button-1>", self._toggle_add_panel)

        self.add_card = tk.Frame(self.root, bg=COLORS["card"],
                                 highlightbackground=COLORS["border"], highlightthickness=1)
        self.add_card.pack(fill="x", padx=24, pady=(4, 8))

        self.inputs = {}
        input_rows = [
            ["日期", "訂購人", "電話", "地址", "贈品", "品項"],
            ["備註", "尺寸", "數量", "單價", "總價", "折扣"],
            ["特殊折扣", "實拿", "成本(品項+贈品)", "利潤", "付款方式", "是否已付"],
        ]

        form = tk.Frame(self.add_card, bg=COLORS["card"])
        form.pack(fill="x", padx=20, pady=16)

        for row_idx, row_cols in enumerate(input_rows):
            for col_idx, field in enumerate(row_cols):
                cell = tk.Frame(form, bg=COLORS["card"])
                cell.grid(row=row_idx, column=col_idx, padx=6, pady=6, sticky="ew")

                tk.Label(cell, text=field, bg=COLORS["card"], fg=COLORS["text_light"],
                         font=("Arial", 10)).pack(anchor="w")
                ent = tk.Entry(cell, width=14,
                               bg=COLORS["input_bg"], fg=COLORS["text"],
                               insertbackground=COLORS["text"],
                               font=("Arial", 11), relief="solid",
                               highlightthickness=0, bd=1)
                ent.pack(fill="x", ipady=4)
                if field == "日期":
                    ent.insert(0, datetime.now().strftime("%Y-%m-%d"))
                self.inputs[field] = ent

            form.columnconfigure(col_idx, weight=1)

        # 儲存按鈕
        btn_row = tk.Frame(self.add_card, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=20, pady=(0, 16))
        save_btn = tk.Button(btn_row, text="💾  儲存紀錄",
                             command=self.add_one,
                             bg=COLORS["success"], fg="white",
                             activebackground="#256d29", activeforeground="white",
                             font=("Arial", 12, "bold"),
                             relief="flat", padx=24, pady=8, cursor="hand2")
        save_btn.pack(side="right")

        # === 表格區 ===
        table_outer = tk.Frame(self.root, bg=COLORS["card"],
                               highlightbackground=COLORS["border"], highlightthickness=1)
        table_outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        table_frame = tk.Frame(table_outer, bg=COLORS["card"])
        table_frame.pack(fill="both", expand=True, padx=2, pady=2)

        self.tree = ttk.Treeview(table_frame, columns=self.cols,
                                 show='headings', style="Custom.Treeview")

        col_widths = {
            "日期": 90, "訂購人": 70, "電話": 100, "地址": 180,
            "贈品": 60, "品項": 70, "備註": 80, "尺寸": 55,
            "數量": 50, "單價": 70, "總價": 70, "折扣": 50,
            "特殊折扣": 65, "實拿": 70, "成本(品項+贈品)": 110,
            "利潤": 65, "付款方式": 70, "是否已付": 65,
        }
        for col in self.cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 65), anchor="center")

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical",
                                    command=self.tree.yview, style="Custom.Vertical.TScrollbar")
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal",
                                    command=self.tree.xview, style="Custom.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        # 斑馬紋
        self.tree.tag_configure('odd', background=COLORS["table_row_alt"])

        # === 狀態列 ===
        self.status_var = tk.StringVar(value="就緒")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                              bg=COLORS["border"], fg=COLORS["text_light"],
                              font=("Arial", 10), anchor="w", padx=12)
        status_bar.pack(fill="x", side="bottom")

    # ─── UI 輔助 ───
    def _make_btn(self, parent, text, bg, fg, cmd):
        btn = tk.Button(parent, text=text, command=cmd,
                        bg=bg, fg=fg, activebackground=bg,
                        font=("Arial", 11), relief="flat",
                        padx=14, pady=6, cursor="hand2")
        return btn

    def _search_focus_in(self, event):
        if self.search_ent.get() == "搜尋姓名、電話或地址...":
            self.search_ent.delete(0, tk.END)
            self.search_ent.config(fg=COLORS["text"])

    def _search_focus_out(self, event):
        if not self.search_ent.get():
            self.search_ent.insert(0, "搜尋姓名、電話或地址...")
            self.search_ent.config(fg=COLORS["text_light"])

    def _toggle_add_panel(self, event=None):
        if self.add_visible.get():
            self.add_card.pack_forget()
            self.toggle_btn.config(text="▶ 快速新增紀錄")
        else:
            # 重新插入到正確位置（toggle_btn 的 frame 後面、table 前面）
            self.add_card.pack(fill="x", padx=24, pady=(4, 8),
                               after=self.toggle_btn.master)
        self.toggle_btn.config(text="▼ 快速新增紀錄")
        self.add_visible.set(not self.add_visible.get())

    # ─── 資料操作 ───
    def load_master_data(self):
        if os.path.exists(MASTER_FILE):
            try:
                self.df = pd.read_excel(MASTER_FILE).fillna("")
                # 統一日期欄位為字串
                self.df["日期"] = self.df["日期"].astype(str).str.split(" ").str[0]
            except Exception:
                self.df = pd.DataFrame(columns=self.cols)
        else:
            self.df = pd.DataFrame(columns=self.cols)

    def display(self, target_df):
        for r in self.tree.get_children():
            self.tree.delete(r)

        if not target_df.empty:
            show_df = target_df.copy()
            show_df["日期"] = show_df["日期"].astype(str).str.split(" ").str[0]
            sorted_df = show_df.sort_values(by="日期", ascending=False)
        else:
            sorted_df = target_df

        for idx, (_, row) in enumerate(sorted_df.fillna("").iterrows()):
            values = [row.get(c, "") for c in self.cols]
            tag = ('odd',) if idx % 2 == 1 else ()
            self.tree.insert('', 'end', values=values, tags=tag)

        count = len(sorted_df)
        self.status_var.set(f"共 {count} 筆紀錄")

    def add_one(self):
        new_row = {f: self.inputs[f].get().strip() for f in self.inputs}
        if not new_row["訂購人"] or not new_row["電話"]:
            messagebox.showwarning("提示", "「訂購人」與「電話」為必填欄位")
            return

        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.save_and_refresh()

        for f, e in self.inputs.items():
            if f != "日期":
                e.delete(0, tk.END)
        messagebox.showinfo("成功", "已儲存一筆新紀錄")

    def import_data(self):
        files = filedialog.askopenfilenames(title="選擇帳表 Excel",
                                            filetypes=[("Excel", "*.xlsx *.xls")])
        if not files:
            return

        count = 0
        for f in files:
            try:
                temp = pd.read_excel(f)
                rename_dict = {"姓名": "訂購人", "客戶姓名": "訂購人"}
                temp = temp.rename(columns=rename_dict)

                # 統一日期為字串
                if "日期" in temp.columns:
                    temp["日期"] = temp["日期"].astype(str).str.split(" ").str[0]

                valid_cols = [c for c in self.cols if c in temp.columns]
                self.df = pd.concat([self.df, temp[valid_cols]], ignore_index=True)
                count += 1
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取 {f} 失敗：{e}")

        self.df = self.df.drop_duplicates(subset=["日期", "訂購人", "電話", "品項"], keep='last')
        self.save_and_refresh()
        messagebox.showinfo("成功", f"已成功合併 {count} 個檔案")

    def save_and_refresh(self):
        try:
            # 確保日期欄位統一為字串
            self.df["日期"] = self.df["日期"].astype(str).str.split(" ").str[0]
            self.df.to_excel(MASTER_FILE, index=False)
            self.display(self.df)
        except Exception:
            messagebox.showerror("存檔失敗", "請關閉 master_data.xlsx 檔案後再試一次。")

    def search(self):
        q = self.search_ent.get().strip()
        if not q or q == "搜尋姓名、電話或地址...":
            self.display(self.df)
            return

        q = q.lower()
        mask = (self.df["訂購人"].astype(str).str.contains(q, case=False) |
                self.df["電話"].astype(str).str.contains(q, case=False) |
                self.df["地址"].astype(str).str.contains(q, case=False))
        self.display(self.df[mask])


if __name__ == "__main__":
    root = tk.Tk()
    window_width = 1400
    window_height = 850
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    app = CustomerSystem(root)
    root.mainloop()
