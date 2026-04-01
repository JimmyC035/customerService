import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

from app.constants import COLORS, PRODUCT_DB_FILE

PRODUCT_COLS = ["品項", "尺寸", "成本", "單價"]


class ProductDatabase:
    """品項資料庫：讀寫 product_db.xlsx，提供查詢方法"""

    def __init__(self):
        self.reload()

    def reload(self):
        if os.path.exists(PRODUCT_DB_FILE):
            try:
                self.df = pd.read_excel(PRODUCT_DB_FILE).fillna("")
            except Exception:
                self.df = pd.DataFrame(columns=PRODUCT_COLS)
        else:
            self.df = pd.DataFrame(columns=PRODUCT_COLS)

    def save(self):
        self.df.to_excel(PRODUCT_DB_FILE, index=False)

    def get_products(self):
        return sorted(self.df["品項"].astype(str).unique().tolist())

    def get_sizes(self, product):
        sizes = self.df[self.df["品項"].astype(str) == str(product)]["尺寸"].tolist()
        return [str(s) for s in sizes if s != ""]

    def get_price(self, product, size):
        df = self.df.copy()
        df["品項"] = df["品項"].astype(str)
        df["尺寸"] = df["尺寸"].astype(str)
        row = df[(df["品項"] == str(product)) & (df["尺寸"] == str(size))]
        if not row.empty:
            return row.iloc[0]["單價"]
        row = df[(df["品項"] == str(product)) & (df["尺寸"] == "")]
        if not row.empty:
            return row.iloc[0]["單價"]
        return 0

    def get_cost(self, product, size):
        df = self.df.copy()
        df["品項"] = df["品項"].astype(str)
        df["尺寸"] = df["尺寸"].astype(str)
        row = df[(df["品項"] == str(product)) & (df["尺寸"] == str(size))]
        if not row.empty:
            return row.iloc[0]["成本"]
        row = df[(df["品項"] == str(product)) & (df["尺寸"] == "")]
        if not row.empty:
            return row.iloc[0]["成本"]
        return 0

    def has_product(self, product, size=""):
        df = self.df.copy()
        df["品項"] = df["品項"].astype(str)
        df["尺寸"] = df["尺寸"].astype(str)
        if size:
            return not df[(df["品項"] == str(product)) & (df["尺寸"] == str(size))].empty
        return not df[df["品項"] == str(product)].empty

    def add(self, product, size, price, cost):
        new_row = {"品項": product, "尺寸": size, "單價": price, "成本": cost}
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.save()

    def update(self, idx, product, size, price, cost):
        self.df.at[idx, "品項"] = product
        self.df.at[idx, "尺寸"] = size
        self.df.at[idx, "單價"] = price
        self.df.at[idx, "成本"] = cost
        self.save()

    def delete(self, idx):
        self.df = self.df.drop(idx).reset_index(drop=True)
        self.save()

    def delete_product(self, product):
        self.df = self.df[self.df["品項"] != product].reset_index(drop=True)
        self.save()


class ProductManagerUI:
    """品項管理分頁 — 樹狀結構（品項 → 尺寸/單價/成本）"""

    def __init__(self, parent, product_db):
        self.parent = parent
        self.db = product_db
        self._build_ui()
        self.refresh_table()

    def _build_ui(self):
        # 上方新增區
        add_card = tk.Frame(self.parent, bg=COLORS["card"],
                            highlightbackground=COLORS["border"], highlightthickness=1)
        add_card.pack(fill="x", padx=24, pady=(16, 8))

        title = tk.Label(add_card, text="新增品項 / 尺寸",
                         bg=COLORS["card"], fg=COLORS["text"],
                         font=("Arial", 13, "bold"))
        title.pack(anchor="w", padx=20, pady=(16, 8))

        form = tk.Frame(add_card, bg=COLORS["card"])
        form.pack(fill="x", padx=20, pady=(0, 16))

        self.pm_inputs = {}
        for i, field in enumerate(PRODUCT_COLS):
            cell = tk.Frame(form, bg=COLORS["card"])
            cell.grid(row=0, column=i, padx=8, pady=4, sticky="ew")
            tk.Label(cell, text=field, bg=COLORS["card"], fg=COLORS["text_light"],
                     font=("Arial", 10)).pack(anchor="w")
            ent = tk.Entry(cell, width=16,
                           bg=COLORS["input_bg"], fg=COLORS["text"],
                           insertbackground=COLORS["text"],
                           font=("Arial", 11), relief="solid",
                           highlightthickness=0, bd=1)
            ent.pack(fill="x", ipady=4)
            self.pm_inputs[field] = ent
            form.columnconfigure(i, weight=1)

        btn_cell = tk.Frame(form, bg=COLORS["card"])
        btn_cell.grid(row=0, column=len(PRODUCT_COLS), padx=8, pady=4)

        tk.Button(btn_cell, text="新增", command=self._add,
                  bg=COLORS["success"], fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=16, pady=6, cursor="hand2").pack(side="left", padx=4, pady=(18, 0))

        # 搜尋列
        search_frame = tk.Frame(self.parent, bg=COLORS["card"],
                                highlightbackground=COLORS["border"], highlightthickness=1)
        search_frame.pack(fill="x", padx=24, pady=(0, 8))

        search_inner = tk.Frame(search_frame, bg=COLORS["card"])
        search_inner.pack(fill="x", padx=20, pady=8)

        tk.Label(search_inner, text=" 🔍", bg=COLORS["card"],
                 fg=COLORS["text_light"],
                 font=("Arial", 13)).pack(side="left", padx=(0, 4))
        self.search_ent = tk.Entry(search_inner, width=30,
                                    bg=COLORS["card"], fg=COLORS["text_light"],
                                    insertbackground=COLORS["text"],
                                    font=("Arial", 12), relief="flat",
                                    highlightthickness=0)
        self.search_ent.pack(side="left", fill="x", expand=True, padx=4)
        self.search_ent.insert(0, "搜尋品項...")
        self.search_ent.bind("<FocusIn>", self._search_focus_in)
        self.search_ent.bind("<FocusOut>", self._search_focus_out)
        self.search_ent.bind("<KeyRelease>", lambda e: self._search())

        # 下方樹狀表格
        table_outer = tk.Frame(self.parent, bg=COLORS["card"],
                               highlightbackground=COLORS["border"], highlightthickness=1)
        table_outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        table_frame = tk.Frame(table_outer, bg=COLORS["card"])
        table_frame.pack(fill="both", expand=True, padx=2, pady=2)

        tree_cols = ("尺寸", "成本", "單價")
        self.tree = ttk.Treeview(table_frame, columns=tree_cols,
                                 show='tree headings', style="Custom.Treeview")

        self.tree.heading("#0", text="品項", anchor="w")
        self.tree.column("#0", width=250, minwidth=200)
        self.tree.heading("尺寸", text="尺寸")
        self.tree.column("尺寸", width=120, anchor="center")
        self.tree.heading("成本", text="成本")
        self.tree.column("成本", width=150, anchor="center")
        self.tree.heading("單價", text="單價")
        self.tree.column("單價", width=150, anchor="center")

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical",
                                    command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure('odd', background=COLORS["table_row_alt"])
        self.tree.tag_configure('product', font=("Arial", 11, "bold"))

        # 右鍵選單
        self.ctx_menu = tk.Menu(self.tree, tearoff=0)
        self.ctx_menu.add_command(label="✏️  修改", command=self._edit_selected)
        self.ctx_menu.add_command(label="🗑  刪除", command=self._delete_selected)

        self.tree.bind("<Button-2>", self._show_context_menu)
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Control-Button-1>", self._show_context_menu)

    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _search_focus_in(self, event):
        if self.search_ent.get() == "搜尋品項...":
            self.search_ent.delete(0, tk.END)
            self.search_ent.config(fg=COLORS["text"])

    def _search_focus_out(self, event):
        if not self.search_ent.get():
            self.search_ent.insert(0, "搜尋品項...")
            self.search_ent.config(fg=COLORS["text_light"])

    def _search(self):
        q = self.search_ent.get().strip().lower()
        if not q or q == "搜尋品項...":
            self.refresh_table()
        else:
            self.refresh_table(filter_text=q)

    def refresh_table(self, filter_text=None):
        self.db.reload()
        for r in self.tree.get_children():
            self.tree.delete(r)

        # 按品項分組（統一轉字串避免混合型別）
        df = self.db.df.copy()
        df["品項"] = df["品項"].astype(str)
        df["_orig_idx"] = df.index

        if filter_text:
            df = df[df["品項"].str.lower().str.contains(filter_text, na=False)]

        products = df.groupby("品項", sort=True)
        row_count = 0
        pid = 0
        for product_name, group in products:
            product_name = str(product_name)

            if len(group) == 1 and group.iloc[0]["尺寸"] == "":
                # 無尺寸的品項，直接顯示為一行
                r = group.iloc[0]
                tag = ('odd',) if row_count % 2 == 1 else ()
                self.tree.insert('', 'end', iid=f"row_{r['_orig_idx']}",
                                 text=product_name,
                                 values=(r["尺寸"], r["成本"], r["單價"]),
                                 tags=tag)
                row_count += 1
            else:
                # 有多個尺寸，用樹狀展開
                parent_id = f"p_{pid}"
                pid += 1
                tag = ('product',) + (('odd',) if row_count % 2 == 1 else ())
                self.tree.insert('', 'end', iid=parent_id,
                                 text=f"📦 {product_name}  ({len(group)} 個尺寸)",
                                 values=("", "", ""),
                                 tags=tag, open=True)
                row_count += 1

                for _, r in group.iterrows():
                    child_tag = ('odd',) if row_count % 2 == 1 else ()
                    self.tree.insert(parent_id, 'end', iid=f"row_{r['_orig_idx']}",
                                     text="",
                                     values=(r["尺寸"], r["成本"], r["單價"]),
                                     tags=child_tag)
                    row_count += 1

    def _get_selected_info(self):
        """回傳選取項目的 (value, is_product_group)
        - product group: value = product_name (str)
        - single row: value = df index (int)
        """
        sel = self.tree.selection()
        if not sel:
            return None, False

        item_id = sel[0]
        if item_id.startswith("p_"):
            # 從 tree item text 取得品項名稱
            text = self.tree.item(item_id, 'text')
            # text format: "📦 品項名  (N 個尺寸)"
            product_name = text.split("📦 ")[1].split("  (")[0] if "📦" in text else text
            return product_name, True
        elif item_id.startswith("row_"):
            return int(item_id.replace("row_", "")), False
        return None, False

    def _edit_selected(self):
        item_id, is_group = self._get_selected_info()
        if item_id is None:
            return

        if is_group:
            messagebox.showinfo("提示", "請展開品項後，選擇要修改的尺寸")
            return

        idx = item_id
        row = self.db.df.iloc[idx]
        EditProductDialog(self.tree.winfo_toplevel(), row, idx, self.db, self.refresh_table)

    def _delete_selected(self):
        item_id, is_group = self._get_selected_info()
        if item_id is None:
            return

        if is_group:
            product_name = item_id
            if messagebox.askyesno("確認", f"確定要刪除品項「{product_name}」的所有尺寸？"):
                self.db.delete_product(product_name)
                self.refresh_table()
        else:
            idx = item_id
            row = self.db.df.iloc[idx]
            label = f"{row['品項']} {row['尺寸']}" if row['尺寸'] else row['品項']
            if messagebox.askyesno("確認", f"確定要刪除「{label}」？"):
                self.db.delete(idx)
                self.refresh_table()

    def _get_form_values(self):
        product = self.pm_inputs["品項"].get().strip()
        size = self.pm_inputs["尺寸"].get().strip()
        try:
            price = float(self.pm_inputs["單價"].get().strip() or 0)
        except ValueError:
            messagebox.showwarning("提示", "單價必須是數字")
            return None
        try:
            cost = float(self.pm_inputs["成本"].get().strip() or 0)
        except ValueError:
            messagebox.showwarning("提示", "成本必須是數字")
            return None
        if not product:
            messagebox.showwarning("提示", "品項名稱不能為空")
            return None
        return product, size, price, cost

    def _add(self):
        vals = self._get_form_values()
        if not vals:
            return
        self.db.add(*vals)
        self.refresh_table()
        self._clear_form()

    def _clear_form(self):
        for ent in self.pm_inputs.values():
            ent.delete(0, tk.END)


class EditProductDialog:
    """修改品項的彈窗"""

    def __init__(self, parent, row, idx, db, on_done):
        self.db = db
        self.idx = idx
        self.on_done = on_done

        self.win = tk.Toplevel(parent)
        self.win.title("修改品項")
        self.win.configure(bg=COLORS["card"])
        self.win.geometry("400x280")
        self.win.resizable(False, False)
        self.win.grab_set()

        self.win.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 280) // 2
        self.win.geometry(f"+{x}+{y}")

        tk.Label(self.win, text="修改品項", bg=COLORS["card"], fg=COLORS["text"],
                 font=("Arial", 14, "bold")).pack(pady=(16, 12))

        self.entries = {}
        form = tk.Frame(self.win, bg=COLORS["card"])
        form.pack(fill="x", padx=24)

        for field in PRODUCT_COLS:
            row_frame = tk.Frame(form, bg=COLORS["card"])
            row_frame.pack(fill="x", pady=4)
            tk.Label(row_frame, text=field, width=8, anchor="e",
                     bg=COLORS["card"], fg=COLORS["text_light"],
                     font=("Arial", 11)).pack(side="left", padx=(0, 8))
            ent = tk.Entry(row_frame,
                           bg=COLORS["input_bg"], fg=COLORS["text"],
                           insertbackground=COLORS["text"],
                           font=("Arial", 11), relief="solid",
                           highlightthickness=0, bd=1)
            ent.pack(side="left", fill="x", expand=True, ipady=4)
            ent.insert(0, str(row[field]))
            self.entries[field] = ent

        btn_frame = tk.Frame(self.win, bg=COLORS["card"])
        btn_frame.pack(fill="x", padx=24, pady=16)

        tk.Button(btn_frame, text="取消", command=self.win.destroy,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Arial", 11), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)
        tk.Button(btn_frame, text="儲存", command=self._save,
                  bg=COLORS["success"], fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)

    def _save(self):
        product = self.entries["品項"].get().strip()
        size = self.entries["尺寸"].get().strip()
        try:
            price = float(self.entries["單價"].get().strip() or 0)
        except ValueError:
            messagebox.showwarning("提示", "單價必須是數字")
            return
        try:
            cost = float(self.entries["成本"].get().strip() or 0)
        except ValueError:
            messagebox.showwarning("提示", "成本必須是數字")
            return
        if not product:
            messagebox.showwarning("提示", "品項名稱不能為空")
            return

        self.db.update(self.idx, product, size, price, cost)
        self.on_done()
        self.win.destroy()
