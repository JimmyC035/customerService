import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

from app.constants import COLORS, PAYMENT_METHODS, PAYMENT_STATUS


# Combobox 欄位
COMBOBOX_FIELDS = {"品項", "尺寸", "付款方式", "是否已付"}
# 自動計算的 readonly 欄位
READONLY_FIELDS = {"單價", "總價", "實拿", "成本(品項+贈品)", "利潤"}


class FormPanel:
    """快速新增紀錄的表單面板"""

    def __init__(self, parent, on_save, product_db=None):
        self.parent = parent
        self.on_save = on_save
        self.product_db = product_db

        self.add_visible = tk.BooleanVar(value=True)
        self._build_toggle()
        self._build_form()

    def _build_toggle(self):
        toggle_frame = tk.Frame(self.parent, bg=COLORS["bg"])
        toggle_frame.pack(fill="x", padx=24)
        self.toggle_btn = tk.Label(toggle_frame, text="▼ 快速新增紀錄",
                                   bg=COLORS["bg"], fg=COLORS["primary"],
                                   font=("Arial", 12, "bold"), cursor="hand2")
        self.toggle_btn.pack(anchor="w")
        self.toggle_btn.bind("<Button-1>", self._toggle)
        self.toggle_frame = toggle_frame

    def _build_form(self):
        self.add_card = tk.Frame(self.parent, bg=COLORS["card"],
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

                if field in COMBOBOX_FIELDS:
                    widget = ttk.Combobox(cell, width=12, font=("Arial", 11))
                    widget.pack(fill="x", ipady=4)
                    self._setup_combobox(field, widget)
                elif field in READONLY_FIELDS:
                    widget = tk.Entry(cell, width=14,
                                      bg="#e8edf2", fg=COLORS["text"],
                                      insertbackground=COLORS["text"],
                                      font=("Arial", 11), relief="solid",
                                      highlightthickness=0, bd=1,
                                      state="readonly",
                                      readonlybackground="#e8edf2")
                    widget.pack(fill="x", ipady=4)
                else:
                    widget = tk.Entry(cell, width=14,
                                      bg=COLORS["input_bg"], fg=COLORS["text"],
                                      insertbackground=COLORS["text"],
                                      font=("Arial", 11), relief="solid",
                                      highlightthickness=0, bd=1)
                    widget.pack(fill="x", ipady=4)

                if field == "日期":
                    widget.insert(0, datetime.now().strftime("%Y-%m-%d"))

                # 綁定自動計算事件
                if field == "數量":
                    widget.bind("<KeyRelease>", lambda e: self._recalculate())
                elif field == "折扣":
                    widget.bind("<KeyRelease>", lambda e: self._recalculate())
                elif field == "特殊折扣":
                    widget.bind("<KeyRelease>", lambda e: self._recalculate())

                self.inputs[field] = widget

            form.columnconfigure(col_idx, weight=1)

        # ── Row 4：附加費用 ──
        extra_row = tk.Frame(self.add_card, bg=COLORS["card"])
        extra_row.pack(fill="x", padx=20, pady=(0, 8))

        self.card_fee_var = tk.BooleanVar(value=False)
        self.invoice_fee_var = tk.BooleanVar(value=False)

        cb_card = tk.Checkbutton(extra_row, text="刷卡手續費 (4%)",
                                 variable=self.card_fee_var,
                                 bg=COLORS["card"], fg=COLORS["text"],
                                 activebackground=COLORS["card"],
                                 font=("Arial", 11),
                                 command=self._recalculate)
        cb_card.pack(side="left", padx=(0, 16))

        cb_invoice = tk.Checkbutton(extra_row, text="開發票 (11%)",
                                    variable=self.invoice_fee_var,
                                    bg=COLORS["card"], fg=COLORS["text"],
                                    activebackground=COLORS["card"],
                                    font=("Arial", 11),
                                    command=self._recalculate)
        cb_invoice.pack(side="left", padx=(0, 16))

        tk.Label(extra_row, text="運費:", bg=COLORS["card"], fg=COLORS["text"],
                 font=("Arial", 11)).pack(side="left", padx=(0, 4))
        self.shipping_ent = tk.Entry(extra_row, width=10,
                                     bg=COLORS["input_bg"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Arial", 11), relief="solid",
                                     highlightthickness=0, bd=1)
        self.shipping_ent.pack(side="left", ipady=4)
        self.shipping_ent.bind("<KeyRelease>", lambda e: self._recalculate())

        # 儲存按鈕
        btn_row = tk.Frame(self.add_card, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=20, pady=(0, 16))
        save_btn = tk.Button(btn_row, text="💾  儲存紀錄",
                             command=self._save,
                             bg=COLORS["success"], fg="white",
                             activebackground="#256d29", activeforeground="white",
                             font=("Arial", 12, "bold"),
                             relief="flat", padx=24, pady=8, cursor="hand2")
        save_btn.pack(side="right")

    # ─── Combobox 設定 ───

    def _setup_combobox(self, field, widget):
        if field == "品項":
            if self.product_db:
                widget['values'] = self.product_db.get_products()
            widget.bind("<<ComboboxSelected>>", self._on_product_selected)
        elif field == "尺寸":
            widget.bind("<<ComboboxSelected>>", self._on_size_selected)
        elif field == "付款方式":
            widget['values'] = PAYMENT_METHODS
            widget.bind("<<ComboboxSelected>>", self._on_payment_changed)
        elif field == "是否已付":
            widget['values'] = PAYMENT_STATUS

    def _on_product_selected(self, event=None):
        product = self.inputs["品項"].get()
        if not self.product_db:
            return
        sizes = self.product_db.get_sizes(product)
        size_widget = self.inputs["尺寸"]
        size_widget['values'] = sizes
        size_widget.set('')
        if sizes and len(sizes) == 1:
            size_widget.set(sizes[0])
            self._on_size_selected()
        else:
            # 無尺寸的品項也帶入單價和成本
            self._on_size_selected()

    def _on_size_selected(self, event=None):
        if not self.product_db:
            return
        product = self.inputs["品項"].get()
        size = self.inputs["尺寸"].get()
        price = self.product_db.get_price(product, size)
        cost = self.product_db.get_cost(product, size)

        self._set_readonly("單價", price)
        self._base_cost = float(cost) if cost else 0
        self._recalculate()

    def _on_payment_changed(self, event=None):
        method = self.inputs["付款方式"].get()
        if method == "刷卡":
            self.card_fee_var.set(True)
        else:
            self.card_fee_var.set(False)
        self._recalculate()

    # ─── 自動計算 ───

    def _recalculate(self):
        price = self._get_num("單價")
        qty = self._get_num("數量")
        discount = self._get_num("折扣")
        special_discount = self._get_num("特殊折扣")

        # 總價 = 單價 × 數量
        total = price * qty
        self._set_readonly("總價", total)

        # 實拿 = 總價 × 折扣 - 特殊折扣
        if discount > 0:
            actual = total * discount - special_discount
        else:
            actual = total - special_discount
        self._set_readonly("實拿", actual)

        # 成本 = 基礎成本 + 附加費用
        base_cost = getattr(self, '_base_cost', 0)
        extra = 0
        if self.card_fee_var.get():
            extra += actual * 0.04
        if self.invoice_fee_var.get():
            extra += actual * 0.11
        shipping = self._get_num_from_entry(self.shipping_ent)
        final_cost = base_cost + extra + shipping
        self._set_readonly("成本(品項+贈品)", final_cost)

        # 利潤 = 實拿 - 成本
        profit = actual - final_cost
        self._set_readonly("利潤", profit)

    def _get_num(self, field):
        try:
            return float(self.inputs[field].get())
        except (ValueError, TypeError):
            return 0

    def _get_num_from_entry(self, entry):
        try:
            return float(entry.get())
        except (ValueError, TypeError):
            return 0

    def _set_readonly(self, field, value):
        widget = self.inputs.get(field)
        if not widget:
            return
        widget.config(state="normal")
        widget.delete(0, tk.END)
        if value:
            # 整數就不顯示小數點
            val = int(value) if value == int(value) else round(value, 1)
            widget.insert(0, str(val))
        widget.config(state="readonly")

    # ─── 公開方法 ───

    def refresh_products(self):
        if self.product_db:
            self.product_db.reload()
            self.inputs["品項"]['values'] = self.product_db.get_products()

    # ─── 儲存 ───

    def _save(self):
        new_row = {f: self.inputs[f].get().strip() for f in self.inputs}
        if not new_row["訂購人"] or not new_row["電話"]:
            messagebox.showwarning("提示", "「訂購人」與「電話」為必填欄位")
            return

        self.on_save(new_row)

        # 清空輸入（保留日期）
        for f, widget in self.inputs.items():
            if f == "日期":
                continue
            if f in READONLY_FIELDS:
                widget.config(state="normal")
                widget.delete(0, tk.END)
                widget.config(state="readonly")
            elif isinstance(widget, ttk.Combobox):
                widget.set('')
            else:
                widget.delete(0, tk.END)

        self.shipping_ent.delete(0, tk.END)
        self.card_fee_var.set(False)
        self.invoice_fee_var.set(False)
        self._base_cost = 0
        messagebox.showinfo("成功", "已儲存一筆新紀錄")

    def _toggle(self, event=None):
        if self.add_visible.get():
            self.add_card.pack_forget()
            self.toggle_btn.config(text="▶ 快速新增紀錄")
        else:
            self.add_card.pack(fill="x", padx=24, pady=(4, 8),
                               after=self.toggle_frame)
            self.toggle_btn.config(text="▼ 快速新增紀錄")
        self.add_visible.set(not self.add_visible.get())
