import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import os

from app.constants import COLORS, COLS, MASTER_FILE
from app.product_manager import ProductDatabase, ProductManagerUI
from app.form_panel import FormPanel
from app.table_panel import TablePanel
from app.analytics import AnalyticsPanel


class CustomerSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("客戶帳務管理系統")
        self.root.geometry("1400x850")
        self.root.configure(bg=COLORS["bg"])

        self.product_db = ProductDatabase()
        self.load_master_data()
        self.setup_styles()
        self.build_ui()
        self.table.display(self.df)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

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

        style.configure("Custom.Vertical.TScrollbar",
                        troughcolor=COLORS["bg"], borderwidth=0)
        style.configure("Custom.Horizontal.TScrollbar",
                        troughcolor=COLORS["bg"], borderwidth=0)

        style.configure("Custom.TNotebook", background=COLORS["bg"], borderwidth=0)
        style.configure("Custom.TNotebook.Tab",
                        background=COLORS["border"],
                        foreground=COLORS["text"],
                        padding=[16, 8],
                        font=("Arial", 12))
        style.map("Custom.TNotebook.Tab",
                  background=[('selected', COLORS["primary"])],
                  foreground=[('selected', 'white')])

    def build_ui(self):
        # 標題列
        header = tk.Frame(self.root, bg=COLORS["primary"], height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="客戶帳務管理系統",
                 bg=COLORS["primary"], fg="white",
                 font=("Arial", 18, "bold")).pack(side="left", padx=24)

        # Notebook 分頁
        self.notebook = ttk.Notebook(self.root, style="Custom.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        # ── Tab 1：訂單管理 ──
        order_tab = tk.Frame(self.notebook, bg=COLORS["bg"])
        self.notebook.add(order_tab, text="  訂單管理  ")

        self.form = FormPanel(order_tab, on_save=self._on_add_record,
                              product_db=self.product_db)

        self.table = TablePanel(order_tab, get_df=lambda: self.df,
                                on_delete=self._on_delete_record,
                                on_edit=self._on_edit_record)
        self._make_btn(self.table.btn_frame, "📥 匯入帳表", "#e3f2fd",
                       COLORS["primary"], self.import_data).pack(side="left", padx=4)

        # ── Tab 2：銷售分析 ──
        analytics_tab = tk.Frame(self.notebook, bg=COLORS["bg"])
        self.notebook.add(analytics_tab, text="  銷售分析  ")
        self.analytics = AnalyticsPanel(analytics_tab, get_df=lambda: self.df)

        # ── Tab 3：品項管理 ──
        product_tab = tk.Frame(self.notebook, bg=COLORS["bg"])
        self.notebook.add(product_tab, text="  品項管理  ")
        self.product_manager_ui = ProductManagerUI(product_tab, self.product_db)

        # 切換分頁時刷新品項下拉
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _on_tab_changed(self, event=None):
        current = self.notebook.index(self.notebook.select())
        if current == 0:
            self.form.refresh_products()
        elif current == 1:
            self.analytics.refresh()

    # ─── 資料操作 ───

    def load_master_data(self):
        if os.path.exists(MASTER_FILE):
            try:
                self.df = pd.read_excel(MASTER_FILE, dtype={"電話": str}).fillna("")
                self.df["日期"] = self.df["日期"].astype(str).str.split(" ").str[0]
                # 過濾掉合計行等非資料行
                self.df = self.df[self.df["日期"].str.match(r'^\d{4}-', na=False)].reset_index(drop=True)
            except Exception:
                self.df = pd.DataFrame(columns=COLS)
        else:
            self.df = pd.DataFrame(columns=COLS)

    def _on_add_record(self, new_row):
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.save_and_refresh()

    def _on_delete_record(self, df_idx):
        self.df = self.df.drop(df_idx).reset_index(drop=True)
        self.save_and_refresh()

    def _on_edit_record(self, df_idx, new_data):
        for col, val in new_data.items():
            if col in self.df.columns:
                self.df.at[df_idx, col] = val
        self.save_and_refresh()

    def import_data(self):
        files = filedialog.askopenfilenames(title="選擇帳表 Excel",
                                            filetypes=[("Excel", "*.xlsx *.xls")])
        if not files:
            return

        count = 0
        for f in files:
            try:
                temp = pd.read_excel(f, dtype={"電話": str})
                rename_dict = {"姓名": "訂購人", "客戶姓名": "訂購人"}
                temp = temp.rename(columns=rename_dict)

                if "日期" in temp.columns:
                    temp["日期"] = temp["日期"].astype(str).str.split(" ").str[0]

                valid_cols = [c for c in COLS if c in temp.columns]
                self.df = pd.concat([self.df, temp[valid_cols]], ignore_index=True)
                count += 1
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取 {f} 失敗：{e}")

        self.df = self.df.drop_duplicates(
            subset=["日期", "訂購人", "電話", "品項"], keep='last')
        self.save_and_refresh()
        messagebox.showinfo("成功", f"已成功合併 {count} 個檔案")

    def save_and_refresh(self):
        try:
            self.df["日期"] = self.df["日期"].astype(str).str.split(" ").str[0]
            self.df.to_excel(MASTER_FILE, index=False)
            self.table.display(self.df)
        except Exception:
            messagebox.showerror("存檔失敗", "請關閉 master_data.xlsx 檔案後再試一次。")

    def _make_btn(self, parent, text, bg, fg, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg=fg, activebackground=bg,
                         font=("Arial", 11), relief="flat",
                         padx=14, pady=6, cursor="hand2")
