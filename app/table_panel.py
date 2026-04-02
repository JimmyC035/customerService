import tkinter as tk
from tkinter import ttk, messagebox

from app.constants import COLORS, COLS, COL_WIDTHS, PAYMENT_METHODS, PAYMENT_STATUS


class TablePanel:
    """Treeview 表格顯示 + 搜尋 + 右鍵選單"""

    def __init__(self, parent, get_df, on_delete=None, on_edit=None, product_db=None):
        """
        parent: 父容器
        get_df: callable，回傳目前的 DataFrame
        on_delete: callback(df_index) 刪除紀錄
        on_edit: callback(df_index, new_row_dict) 修改紀錄
        product_db: ProductDatabase 實例
        """
        self.parent = parent
        self.get_df = get_df
        self.on_delete = on_delete
        self.on_edit = on_edit
        self.product_db = product_db

        # 儲存顯示順序對應的 df index
        self._display_indices = []

        self._build_toolbar()
        self._build_table()
        self._build_status_bar()
        self._build_context_menu()

    def _build_toolbar(self):
        toolbar = tk.Frame(self.parent, bg=COLORS["bg"])
        toolbar.pack(fill="x", padx=24, pady=(16, 8))

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

        btn_frame = tk.Frame(toolbar, bg=COLORS["bg"])
        btn_frame.pack(side="right", padx=(16, 0))

        self._make_btn(btn_frame, "顯示全部", COLORS["card"], COLORS["text"],
                       lambda: self.display(self.get_df())).pack(side="left", padx=4)

        self.btn_frame = btn_frame

    def _build_table(self):
        table_outer = tk.Frame(self.parent, bg=COLORS["card"],
                               highlightbackground=COLORS["border"], highlightthickness=1)
        table_outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        table_frame = tk.Frame(table_outer, bg=COLORS["card"])
        table_frame.pack(fill="both", expand=True, padx=2, pady=2)

        self.tree = ttk.Treeview(table_frame, columns=COLS,
                                 show='headings', style="Custom.Treeview")

        for col in COLS:
            self.tree.heading(col, text=col)
            w = COL_WIDTHS.get(col, 75)
            self.tree.column(col, width=w, minwidth=w, stretch=False, anchor="center")

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical",
                                    command=self.tree.yview,
                                    style="Custom.Vertical.TScrollbar")
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal",
                                    command=self.tree.xview,
                                    style="Custom.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure('odd', background=COLORS["table_row_alt"])
        self.tree.tag_configure('summary', background="#e3f0fa", font=("Arial", 11, "bold"))

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="就緒")
        status_bar = tk.Label(self.parent, textvariable=self.status_var,
                              bg=COLORS["border"], fg=COLORS["text_light"],
                              font=("Arial", 10), anchor="w", padx=12)
        status_bar.pack(fill="x", side="bottom")

    def _build_context_menu(self):
        self.ctx_menu = tk.Menu(self.tree, tearoff=0)

        # 複製子選單
        copy_menu = tk.Menu(self.ctx_menu, tearoff=0)
        copy_menu.add_command(label="全部欄位", command=self._copy_selected)
        copy_menu.add_command(label="地址", command=lambda: self._copy_column("地址"))
        copy_menu.add_command(label="電話", command=lambda: self._copy_column("電話"))
        self.ctx_menu.add_cascade(label="📋  複製", menu=copy_menu)

        self.ctx_menu.add_command(label="✏️  修改此筆", command=self._edit_selected)
        self.ctx_menu.add_command(label="🗑  刪除此筆", command=self._delete_selected)

        # macOS 用 Button-2，Windows/Linux 用 Button-3
        self.tree.bind("<Button-2>", self._show_context_menu)
        self.tree.bind("<Button-3>", self._show_context_menu)
        # macOS Control-Click
        self.tree.bind("<Control-Button-1>", self._show_context_menu)

        # Cmd+C / Ctrl+C 複製選取行
        self.tree.bind("<Command-c>", lambda e: self._copy_selected())
        self.tree.bind("<Control-c>", lambda e: self._copy_selected())

    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item or item == '_summary':
            return
        self.tree.selection_set(item)
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _get_selected_df_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        tree_idx = self.tree.index(sel[0])
        if tree_idx < len(self._display_indices):
            return self._display_indices[tree_idx]
        return None

    def _copy_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        lines = []
        lines.append("\t".join(COLS))
        for item in sel:
            values = self.tree.item(item, 'values')
            lines.append("\t".join(str(v) for v in values))
        text = "\n".join(lines)
        self.tree.clipboard_clear()
        self.tree.clipboard_append(text)

    def _copy_column(self, col_name):
        sel = self.tree.selection()
        if not sel:
            return
        col_idx = COLS.index(col_name)
        values = []
        for item in sel:
            v = self.tree.item(item, 'values')[col_idx]
            if str(v).strip():
                values.append(str(v))
        text = "\n".join(values)
        self.tree.clipboard_clear()
        self.tree.clipboard_append(text)

    def _delete_selected(self):
        df_idx = self._get_selected_df_index()
        if df_idx is None:
            return
        if not messagebox.askyesno("確認", "確定要刪除這筆紀錄？"):
            return
        if self.on_delete:
            self.on_delete(df_idx)

    def _edit_selected(self):
        df_idx = self._get_selected_df_index()
        if df_idx is None:
            return

        sel = self.tree.selection()
        values = self.tree.item(sel[0], 'values')
        row_data = dict(zip(COLS, values))

        EditDialog(self.tree.winfo_toplevel(), row_data,
                   on_save=lambda new_data: self._on_edit_save(df_idx, new_data),
                   product_db=self.product_db)

    def _on_edit_save(self, df_idx, new_data):
        if self.on_edit:
            self.on_edit(df_idx, new_data)

    # ─── 公開方法 ───

    def display(self, target_df):
        import pandas as pd

        for r in self.tree.get_children():
            self.tree.delete(r)

        self._display_indices = []

        if not target_df.empty:
            show_df = target_df.copy()
            show_df["日期"] = show_df["日期"].astype(str).str.split(" ").str[0]
            sorted_df = show_df.sort_values(by="日期", ascending=False)
        else:
            sorted_df = target_df

        int_cols = {"數量", "單價", "總價", "實拿", "成本(品項+贈品)", "利潤"}
        for idx, (df_idx, row) in enumerate(sorted_df.fillna("").iterrows()):
            values = []
            for c in COLS:
                v = row.get(c, "")
                if c in int_cols and v != "":
                    try:
                        v = int(float(v))
                    except (ValueError, TypeError):
                        pass
                values.append(v)
            tag = ('odd',) if idx % 2 == 1 else ()
            self.tree.insert('', 'end', values=values, tags=tag)
            self._display_indices.append(df_idx)

        # 加總行
        if not sorted_df.empty:
            sum_cols = {"總價", "實拿", "成本(品項+贈品)", "利潤"}
            summary = []
            for c in COLS:
                if c in sum_cols:
                    val = pd.to_numeric(sorted_df[c], errors='coerce').sum()
                    summary.append(f"{val:,.0f}" if val else "")
                elif c == "日期":
                    summary.append("【合計】")
                elif c == "付款方式":
                    # 利潤率
                    total_rev = pd.to_numeric(sorted_df["總價"], errors='coerce').sum()
                    total_profit = pd.to_numeric(sorted_df["利潤"], errors='coerce').sum()
                    rate = (total_profit / total_rev * 100) if total_rev > 0 else 0
                    summary.append(f"利潤率 {rate:.1f}%")
                else:
                    summary.append("")
            self.tree.insert('', 'end', iid='_summary', values=summary, tags=('summary',))

        self.status_var.set(f"共 {len(sorted_df)} 筆紀錄")

    def search(self):
        df = self.get_df()
        q = self.search_ent.get().strip()
        if not q or q == "搜尋姓名、電話或地址...":
            self.display(df)
            return

        q = q.lower()
        mask = (df["訂購人"].astype(str).str.contains(q, case=False) |
                df["電話"].astype(str).str.contains(q, case=False) |
                df["地址"].astype(str).str.contains(q, case=False))
        self.display(df[mask])

    # ─── 輔助 ───

    def _make_btn(self, parent, text, bg, fg, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg=fg, activebackground=bg,
                         font=("Arial", 11), relief="flat",
                         padx=14, pady=6, cursor="hand2")

    def _search_focus_in(self, event):
        if self.search_ent.get() == "搜尋姓名、電話或地址...":
            self.search_ent.delete(0, tk.END)
            self.search_ent.config(fg=COLORS["text"])

    def _search_focus_out(self, event):
        if not self.search_ent.get():
            self.search_ent.insert(0, "搜尋姓名、電話或地址...")
            self.search_ent.config(fg=COLORS["text_light"])


class EditDialog:
    """修改紀錄的彈窗"""

    def __init__(self, parent, row_data, on_save, product_db=None):
        self.on_save = on_save
        self.product_db = product_db

        self.win = tk.Toplevel(parent)
        self.win.title("修改紀錄")
        self.win.configure(bg=COLORS["card"])
        self.win.geometry("600x700")
        self.win.resizable(False, True)
        self.win.grab_set()

        # 置中
        self.win.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 700) // 2
        self.win.geometry(f"+{x}+{y}")

        tk.Label(self.win, text="修改紀錄", bg=COLORS["card"], fg=COLORS["text"],
                 font=("Arial", 14, "bold")).pack(pady=(16, 8))

        # 按鈕（先 pack 固定在底部，確保不被擠掉）
        btn_frame = tk.Frame(self.win, bg=COLORS["card"])
        btn_frame.pack(side="bottom", fill="x", padx=20, pady=16)

        tk.Button(btn_frame, text="取消", command=self.win.destroy,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Arial", 11), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)
        tk.Button(btn_frame, text="儲存", command=self._save,
                  bg=COLORS["success"], fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)

        # 可捲動區域
        canvas_container = tk.Frame(self.win, bg=COLORS["card"])
        canvas_container.pack(fill="both", expand=True, padx=20)

        canvas = tk.Canvas(canvas_container, bg=COLORS["card"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=COLORS["card"])

        scroll_frame.bind("<Configure>",
                          lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_win = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # scroll_frame 寬度跟隨 canvas
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_win, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        # 滑鼠滾輪支援
        def _on_mousewheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120 or (-1 if event.delta < 0 else 1)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # 需要自動計算的欄位
        self._calc_fields = {"數量", "單價", "折扣", "特殊折扣"}
        # Combobox 欄位
        self._combobox_cols = {
            "品項": self.product_db.get_products() if self.product_db else [],
            "尺寸": [],
            "付款方式": PAYMENT_METHODS,
            "是否已付": PAYMENT_STATUS,
        }

        self.entries = {}
        self._base_cost = 0
        self._cost_manual = False
        self._manual_cost_base = 0

        for i, col in enumerate(COLS):
            row_frame = tk.Frame(scroll_frame, bg=COLORS["card"])
            row_frame.pack(fill="x", pady=3, padx=4)

            tk.Label(row_frame, text=col, width=14, anchor="e",
                     bg=COLORS["card"], fg=COLORS["text_light"],
                     font=("Arial", 10)).pack(side="left", padx=(0, 8))

            if col in self._combobox_cols:
                widget = ttk.Combobox(row_frame, font=("Arial", 11),
                                      values=self._combobox_cols[col])
                widget.pack(side="left", fill="x", expand=True, ipady=3)
                widget.set(str(row_data.get(col, "")))
                if col == "品項":
                    widget.bind("<<ComboboxSelected>>", self._on_product_changed)
                    widget.bind("<FocusOut>", lambda e: self._on_price_lookup())
                elif col == "尺寸":
                    widget.bind("<<ComboboxSelected>>", lambda e: self._on_price_lookup())
                    widget.bind("<FocusOut>", lambda e: self._on_price_lookup())
                    # 初始化尺寸選項
                    product = str(row_data.get("品項", ""))
                    if product and self.product_db:
                        sizes = self.product_db.get_sizes(product)
                        widget['values'] = sizes
            else:
                widget = tk.Entry(row_frame,
                                  bg=COLORS["input_bg"], fg=COLORS["text"],
                                  insertbackground=COLORS["text"],
                                  font=("Arial", 11), relief="solid",
                                  highlightthickness=0, bd=1)
                widget.pack(side="left", fill="x", expand=True, ipady=3)
                widget.insert(0, str(row_data.get(col, "")))
                if col == "成本(品項+贈品)":
                    widget.bind("<KeyRelease>", lambda e: self._on_cost_manual_edit())
                elif col in self._calc_fields:
                    widget.bind("<KeyRelease>", lambda e: self._recalculate())
            self.entries[col] = widget

        # ── 附加費用區域 ──
        SHIPPING_RATE = 0.04
        CARD_RATE = 0.03
        INVOICE_RATE = 0.05
        self._rates = (SHIPPING_RATE, CARD_RATE, INVOICE_RATE)

        sep = tk.Frame(scroll_frame, bg=COLORS["border"], height=1)
        sep.pack(fill="x", padx=4, pady=8)

        extra_label = tk.Label(scroll_frame, text="附加費用（影響成本計算）",
                               bg=COLORS["card"], fg=COLORS["primary"],
                               font=("Arial", 11, "bold"))
        extra_label.pack(anchor="w", padx=4, pady=(0, 4))

        # 其他成本
        other_row = tk.Frame(scroll_frame, bg=COLORS["card"])
        other_row.pack(fill="x", pady=3, padx=4)
        tk.Label(other_row, text="其他成本", width=14, anchor="e",
                 bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 10)).pack(side="left", padx=(0, 8))
        self.other_cost_ent = tk.Entry(other_row,
                                        bg=COLORS["input_bg"], fg=COLORS["text"],
                                        insertbackground=COLORS["text"],
                                        font=("Arial", 11), relief="solid",
                                        highlightthickness=0, bd=1)
        self.other_cost_ent.pack(side="left", fill="x", expand=True, ipady=3)
        self.other_cost_ent.bind("<KeyRelease>", lambda e: self._recalculate())

        # 勾選框
        cb_row = tk.Frame(scroll_frame, bg=COLORS["card"])
        cb_row.pack(fill="x", pady=6, padx=4)

        self.shipping_var = tk.BooleanVar(value=False)
        self.card_fee_var = tk.BooleanVar(value=False)
        self.invoice_var = tk.BooleanVar(value=False)

        for text, var in [("運費 (4%)", self.shipping_var),
                          ("刷卡 (3%)", self.card_fee_var),
                          ("開發票 (5%)", self.invoice_var)]:
            tk.Checkbutton(cb_row, text=text, variable=var,
                           bg=COLORS["card"], fg=COLORS["text"],
                           activebackground=COLORS["card"],
                           font=("Arial", 10), command=self._recalculate
                           ).pack(side="left", padx=(0, 12))

        # 付款方式連動勾選
        if "付款方式" in self.entries:
            self.entries["付款方式"].bind("<<ComboboxSelected>>",
                                          self._on_payment_selected)

        # 初始化 base_cost
        if self.product_db:
            product = str(row_data.get("品項", ""))
            size = str(row_data.get("尺寸", ""))
            self._base_cost = float(self.product_db.get_cost(product, size) or 0)

    def _on_payment_selected(self, event=None):
        method = self.entries["付款方式"].get()
        if method == "現金":
            self.shipping_var.set(True)
            self.card_fee_var.set(False)
            self.invoice_var.set(False)
        elif method == "刷卡":
            self.shipping_var.set(True)
            self.card_fee_var.set(True)
            self.invoice_var.set(True)
        elif method == "匯款":
            self.shipping_var.set(True)
            self.card_fee_var.set(False)
            self.invoice_var.set(False)
        elif method == "第三方支付":
            self.shipping_var.set(True)
            self.card_fee_var.set(True)
            self.invoice_var.set(True)
        self._recalculate()

    def _on_product_changed(self, event=None):
        if not self.product_db:
            return
        product = self.entries["品項"].get()
        sizes = self.product_db.get_sizes(product)
        self.entries["尺寸"]['values'] = sizes
        if sizes and len(sizes) == 1:
            self.entries["尺寸"].set(sizes[0])
        else:
            self.entries["尺寸"].set('')
        self._on_price_lookup()

    def _on_price_lookup(self):
        """品項/尺寸變動時，從 DB 帶入單價和成本"""
        if not self.product_db:
            return
        product = self.entries["品項"].get().strip()
        size = self.entries["尺寸"].get().strip()
        if not product:
            return
        price = self.product_db.get_price(product, size)
        cost = self.product_db.get_cost(product, size)
        if price:
            self._set_entry(self.entries["單價"], price)
        self._base_cost = float(cost) if cost else 0
        self._recalculate()

    def _on_cost_manual_edit(self):
        """使用者手動編輯成本欄位時，記錄手動基底並重算"""
        self._cost_manual = True
        actual = self._get_num("實拿")
        shipping_rate, card_rate, invoice_rate = self._rates
        extra = 0
        if self.shipping_var.get():
            extra += round(actual * shipping_rate)
        if self.card_fee_var.get():
            extra += round(actual * card_rate)
        if self.invoice_var.get():
            extra += round(actual * invoice_rate)
        self._manual_cost_base = self._get_num("成本(品項+贈品)") - extra
        cost = self._get_num("成本(品項+贈品)")
        profit = actual - cost
        self._set_if_not_focused("利潤", profit)

    def _recalculate(self):
        """自動計算 總價、實拿、成本、利潤"""
        price = self._get_num("單價")
        qty = self._get_num("數量") or 1
        discount = self._get_num("折扣")

        total = price * qty
        if discount > 0:
            total = total * discount
        self._set_if_not_focused("總價", total)

        special = self._get_num("特殊折扣")
        actual = total - special
        self._set_if_not_focused("實拿", actual)

        if self._cost_manual:
            # 手動模式：基底 + 附加費用（強制更新，不受 focus 影響）
            shipping_rate, card_rate, invoice_rate = self._rates
            extra = 0
            if self.shipping_var.get():
                extra += round(actual * shipping_rate)
            if self.card_fee_var.get():
                extra += round(actual * card_rate)
            if self.invoice_var.get():
                extra += round(actual * invoice_rate)
            cost = self._manual_cost_base + extra
            w = self.entries["成本(品項+贈品)"]
            w.delete(0, tk.END)
            w.insert(0, str(int(cost) if cost == int(cost) else round(cost, 1)))
        else:
            # 自動模式：計算成本
            try:
                other_cost = float(self.other_cost_ent.get() or 0)
            except (ValueError, TypeError):
                other_cost = 0

            shipping_rate, card_rate, invoice_rate = self._rates
            extra = 0
            if self.shipping_var.get():
                extra += round(actual * shipping_rate)
            if self.card_fee_var.get():
                extra += round(actual * card_rate)
            if self.invoice_var.get():
                extra += round(actual * invoice_rate)

            cost = (self._base_cost * qty) + other_cost + extra
            self._set_if_not_focused("成本(品項+贈品)", cost)

        profit = actual - cost
        self._set_if_not_focused("利潤", profit)

    def _get_num(self, col):
        try:
            return float(self.entries[col].get())
        except (ValueError, TypeError):
            return 0

    def _set_entry(self, widget, value):
        widget.delete(0, tk.END)
        if value is not None and value != "" and value != 0:
            try:
                v = float(value)
                v = int(v) if v == int(v) else round(v, 1)
                widget.insert(0, str(v))
            except (ValueError, TypeError):
                widget.insert(0, str(value))

    def _set_if_not_focused(self, col, value):
        """設定欄位值，但如果正在被編輯則跳過"""
        widget = self.entries[col]
        try:
            if widget == widget.focus_get():
                return
        except (KeyError, RuntimeError):
            pass
        self._set_entry(widget, value)

    def _save(self):
        new_data = {col: self.entries[col].get().strip() for col in COLS}
        self.on_save(new_data)
        self.win.destroy()
