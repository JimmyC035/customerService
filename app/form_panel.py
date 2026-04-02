import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

from app.constants import COLORS, PAYMENT_METHODS, PAYMENT_STATUS

# 費率
SHIPPING_RATE = 0.04
CARD_RATE = 0.02
INVOICE_RATE = 0.05

MAX_PRODUCT_ROWS = 20


class FormPanel:
    """快速新增紀錄的表單面板（支援多品項）"""

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

        main_form = tk.Frame(self.add_card, bg=COLORS["card"])
        main_form.pack(fill="x", padx=20, pady=(16, 0))

        # ── Row 1: 客戶資訊 ──
        self.shared_inputs = {}
        row1_fields = ["日期", "訂購人", "電話", "地址", "贈品"]
        for i, field in enumerate(row1_fields):
            cell = tk.Frame(main_form, bg=COLORS["card"])
            cell.grid(row=0, column=i, padx=6, pady=6, sticky="ew")
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
            self.shared_inputs[field] = ent
            main_form.columnconfigure(i, weight=1)

        # ── Row 2+: 多品項區域 ──
        product_section = tk.Frame(self.add_card, bg=COLORS["card"])
        product_section.pack(fill="x", padx=20, pady=(8, 8))

        label_row = tk.Frame(product_section, bg=COLORS["card"])
        label_row.pack(fill="x", pady=(0, 4))

        tk.Label(label_row, text="品項明細", bg=COLORS["card"],
                 fg=COLORS["primary"], font=("Arial", 11, "bold")).pack(side="left")
        tk.Button(label_row, text="＋ 新增品項", command=self._add_product_row,
                  bg="#e3f2fd", fg=COLORS["primary"],
                  font=("Arial", 10), relief="flat", padx=10, cursor="hand2"
                  ).pack(side="right")

        # 品項表格（表頭 + 品項行在同一個 grid 框架中，確保對齊）
        self.product_grid = tk.Frame(product_section, bg=COLORS["card"])
        self.product_grid.pack(fill="x")

        headers = ["品項", "尺寸", "數量", "單價", "折扣", "總價", "成本", ""]
        col_weights = [3, 2, 1, 2, 1, 2, 2, 0]
        for i, (h, w) in enumerate(zip(headers, col_weights)):
            tk.Label(self.product_grid, text=h, bg=COLORS["card"],
                     fg=COLORS["text_light"],
                     font=("Arial", 10)).grid(row=0, column=i, padx=6, sticky="w")
            self.product_grid.columnconfigure(i, weight=w)

        self.product_rows = []
        self._product_grid_row = 1
        self._add_product_row()  # 預設一行

        # ── Row 3: 財務 ──
        fin_form = tk.Frame(self.add_card, bg=COLORS["card"])
        fin_form.pack(fill="x", padx=20, pady=(4, 0))

        self.fin_inputs = {}
        fin_fields = ["特殊折扣", "實拿", "成本(品項+贈品)", "利潤"]
        for i, field in enumerate(fin_fields):
            cell = tk.Frame(fin_form, bg=COLORS["card"])
            cell.grid(row=0, column=i, padx=6, pady=6, sticky="ew")
            tk.Label(cell, text=field, bg=COLORS["card"], fg=COLORS["text_light"],
                     font=("Arial", 10)).pack(anchor="w")
            ent = tk.Entry(cell, width=14,
                           bg=COLORS["input_bg"], fg=COLORS["text"],
                           insertbackground=COLORS["text"],
                           font=("Arial", 11), relief="solid",
                           highlightthickness=0, bd=1)
            if field not in ("成本(品項+贈品)", "利潤"):
                ent.bind("<KeyRelease>", lambda e: self._recalculate())
            ent.pack(fill="x", ipady=4)
            self.fin_inputs[field] = ent
            fin_form.columnconfigure(i, weight=1)

        # ── Row 4: 付款 / 附加費用 ──
        pay_form = tk.Frame(self.add_card, bg=COLORS["card"])
        pay_form.pack(fill="x", padx=20, pady=(4, 0))

        # 付款方式
        cell = tk.Frame(pay_form, bg=COLORS["card"])
        cell.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        tk.Label(cell, text="付款方式", bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 10)).pack(anchor="w")
        self.payment_cb = ttk.Combobox(cell, width=12, font=("Arial", 11),
                                        values=PAYMENT_METHODS)
        self.payment_cb.pack(fill="x", ipady=4)
        self.payment_cb.bind("<<ComboboxSelected>>", self._on_payment_selected)

        # 是否已付
        cell2 = tk.Frame(pay_form, bg=COLORS["card"])
        cell2.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
        tk.Label(cell2, text="是否已付", bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 10)).pack(anchor="w")
        self.paid_cb = ttk.Combobox(cell2, width=12, font=("Arial", 11),
                                     values=PAYMENT_STATUS)
        self.paid_cb.pack(fill="x", ipady=4)

        # 其他成本
        cell3 = tk.Frame(pay_form, bg=COLORS["card"])
        cell3.grid(row=0, column=2, padx=6, pady=6, sticky="ew")
        tk.Label(cell3, text="其他成本", bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 10)).pack(anchor="w")
        self.other_cost_ent = tk.Entry(cell3, width=14,
                                        bg=COLORS["input_bg"], fg=COLORS["text"],
                                        insertbackground=COLORS["text"],
                                        font=("Arial", 11), relief="solid",
                                        highlightthickness=0, bd=1)
        self.other_cost_ent.pack(fill="x", ipady=4)
        self.other_cost_ent.bind("<KeyRelease>", lambda e: self._recalculate())

        # 勾選框
        cb_frame = tk.Frame(pay_form, bg=COLORS["card"])
        cb_frame.grid(row=0, column=3, padx=6, pady=6, sticky="sw")

        self.shipping_var = tk.BooleanVar(value=False)
        self.card_fee_var = tk.BooleanVar(value=False)
        self.invoice_var = tk.BooleanVar(value=False)

        for text, var in [("運費 (4%)", self.shipping_var),
                          ("刷卡 (2%)", self.card_fee_var),
                          ("開發票 (5%)", self.invoice_var)]:
            tk.Checkbutton(cb_frame, text=text, variable=var,
                           bg=COLORS["card"], fg=COLORS["text"],
                           activebackground=COLORS["card"],
                           font=("Arial", 10), command=self._recalculate
                           ).pack(side="left", padx=(0, 12))

        for i in range(4):
            pay_form.columnconfigure(i, weight=1)

        # ── 備註 + 儲存 ──
        bottom = tk.Frame(self.add_card, bg=COLORS["card"])
        bottom.pack(fill="x", padx=20, pady=(4, 16))

        remark_cell = tk.Frame(bottom, bg=COLORS["card"])
        remark_cell.pack(side="left", fill="x", expand=True)
        tk.Label(remark_cell, text="備註", bg=COLORS["card"], fg=COLORS["text_light"],
                 font=("Arial", 10)).pack(anchor="w")
        self.remark_ent = tk.Entry(remark_cell, width=30,
                                    bg=COLORS["input_bg"], fg=COLORS["text"],
                                    insertbackground=COLORS["text"],
                                    font=("Arial", 11), relief="solid",
                                    highlightthickness=0, bd=1)
        self.remark_ent.pack(fill="x", ipady=4)

        tk.Button(bottom, text="💾  儲存紀錄", command=self._save,
                  bg=COLORS["success"], fg="white",
                  activebackground="#256d29", activeforeground="white",
                  font=("Arial", 12, "bold"),
                  relief="flat", padx=24, pady=8, cursor="hand2"
                  ).pack(side="right", padx=(16, 0))

    # ─── 多品項行管理 ───

    def _add_product_row(self):
        if len(self.product_rows) >= MAX_PRODUCT_ROWS:
            return

        r = self._product_grid_row
        row_data = {}
        widgets = []

        # 品項 Combobox（支援打字搜尋）
        cb_product = ttk.Combobox(self.product_grid, width=12, font=("Arial", 11))
        all_products = self.product_db.get_products() if self.product_db else []
        cb_product['values'] = all_products
        cb_product.grid(row=r, column=0, padx=6, pady=2, sticky="ew")
        cb_product.bind("<<ComboboxSelected>>",
                        lambda e, rd=row_data: self._on_product_selected(rd))
        cb_product.bind("<KeyRelease>",
                        lambda e, cb=cb_product, full=all_products, rd=row_data:
                        self._on_product_key(e, cb, full, rd))
        cb_product.bind("<FocusOut>",
                        lambda e, rd=row_data: self._check_manual_entry(rd))
        row_data["品項"] = cb_product
        row_data["_all_products"] = all_products
        widgets.append(cb_product)

        # 尺寸 Combobox
        cb_size = ttk.Combobox(self.product_grid, width=10, font=("Arial", 11))
        cb_size.grid(row=r, column=1, padx=6, pady=2, sticky="ew")
        cb_size.bind("<<ComboboxSelected>>",
                     lambda e, rd=row_data: self._on_size_selected(rd))
        cb_size.bind("<FocusOut>",
                     lambda e, rd=row_data: self._check_manual_entry(rd))
        row_data["尺寸"] = cb_size
        widgets.append(cb_size)

        # 數量
        ent_qty = tk.Entry(self.product_grid, width=8,
                           bg=COLORS["input_bg"], fg=COLORS["text"],
                           insertbackground=COLORS["text"],
                           font=("Arial", 11), relief="solid",
                           highlightthickness=0, bd=1)
        ent_qty.grid(row=r, column=2, padx=6, pady=2, sticky="ew")
        ent_qty.bind("<KeyRelease>",
                     lambda e, rd=row_data: self._on_qty_changed(rd))
        row_data["數量"] = ent_qty
        widgets.append(ent_qty)

        # 單價
        ent_price = tk.Entry(self.product_grid, width=10,
                             bg=COLORS["input_bg"], fg=COLORS["text"],
                             insertbackground=COLORS["text"],
                             font=("Arial", 11), relief="solid",
                             highlightthickness=0, bd=1)
        ent_price.grid(row=r, column=3, padx=6, pady=2, sticky="ew")
        ent_price.bind("<KeyRelease>",
                       lambda e, rd=row_data: self._on_qty_changed(rd))
        row_data["單價"] = ent_price
        widgets.append(ent_price)

        # 折扣
        ent_discount = tk.Entry(self.product_grid, width=6,
                                bg=COLORS["input_bg"], fg=COLORS["text"],
                                insertbackground=COLORS["text"],
                                font=("Arial", 11), relief="solid",
                                highlightthickness=0, bd=1)
        ent_discount.grid(row=r, column=4, padx=6, pady=2, sticky="ew")
        ent_discount.bind("<KeyRelease>",
                          lambda e, rd=row_data: self._on_qty_changed(rd))
        row_data["折扣"] = ent_discount
        widgets.append(ent_discount)

        # 總價
        ent_total = tk.Entry(self.product_grid, width=10,
                             bg=COLORS["input_bg"], fg=COLORS["text"],
                             insertbackground=COLORS["text"],
                             font=("Arial", 11), relief="solid",
                             highlightthickness=0, bd=1)
        ent_total.grid(row=r, column=5, padx=6, pady=2, sticky="ew")
        ent_total.bind("<KeyRelease>",
                       lambda e, rd=row_data: self._recalculate())
        row_data["總價"] = ent_total
        widgets.append(ent_total)

        # 成本（可手動編輯，DB 會自動帶入）
        ent_cost = tk.Entry(self.product_grid, width=10,
                            bg=COLORS["input_bg"], fg=COLORS["text"],
                            insertbackground=COLORS["text"],
                            font=("Arial", 11), relief="solid",
                            highlightthickness=0, bd=1)
        ent_cost.grid(row=r, column=6, padx=6, pady=2, sticky="ew")
        ent_cost.bind("<KeyRelease>",
                      lambda e, rd=row_data: self._on_row_cost_edited(rd))
        row_data["成本"] = ent_cost
        row_data["_base_cost"] = 0
        row_data["_cost_manual"] = False
        widgets.append(ent_cost)

        # 刪除按鈕（第一行不顯示）
        if len(self.product_rows) > 0:
            btn_del = tk.Button(self.product_grid, text="✕",
                                command=lambda rd=row_data: self._remove_product_row(rd),
                                bg=COLORS["card"], fg="#c62828",
                                font=("Arial", 10), relief="flat", cursor="hand2")
            btn_del.grid(row=r, column=7, padx=2, pady=2)
            widgets.append(btn_del)

        row_data["_widgets"] = widgets
        self.product_rows.append(row_data)
        self._product_grid_row += 1

    def _remove_product_row(self, row_data):
        for w in row_data["_widgets"]:
            w.destroy()
        self.product_rows.remove(row_data)
        self._recalculate()

    # ─── 品項搜尋 ───

    def _on_product_key(self, event, cb, full_list, row_data):
        # 延遲篩選，等輸入穩定後才更新下拉選單
        if hasattr(cb, '_filter_after_id'):
            cb.after_cancel(cb._filter_after_id)
        cb._filter_after_id = cb.after(300, lambda: self._filter_products(cb, full_list))

    def _filter_products(self, cb, full_list):
        typed = cb.get().strip().lower()
        if not typed:
            cb['values'] = full_list
        else:
            filtered = [p for p in full_list if typed in p.lower()]
            cb['values'] = filtered if filtered else full_list

    # ─── 品項連動 ───

    def _on_product_selected(self, row_data):
        product = row_data["品項"].get()
        if not self.product_db:
            return
        sizes = self.product_db.get_sizes(product)
        row_data["尺寸"]['values'] = sizes
        row_data["尺寸"].set('')
        if sizes and len(sizes) == 1:
            row_data["尺寸"].set(sizes[0])
        self._on_size_selected(row_data)

    def _on_size_selected(self, row_data):
        """從下拉選單選擇尺寸時觸發"""
        if not self.product_db:
            return
        product = row_data["品項"].get()
        size = row_data["尺寸"].get()
        price = self.product_db.get_price(product, size)
        cost = self.product_db.get_cost(product, size)
        self._set_entry(row_data["單價"], price)
        row_data["_base_cost"] = float(cost) if cost else 0
        row_data["_cost_manual"] = False
        self._on_qty_changed(row_data)

    def _check_manual_entry(self, row_data):
        """FocusOut 時檢查：若品項/尺寸在 DB 中則自動帶入單價和成本"""
        if not self.product_db:
            return
        product = row_data["品項"].get().strip()
        size = row_data["尺寸"].get().strip()
        if not product:
            return

        if size:
            found = self.product_db.has_product(product, size)
        else:
            found = self.product_db.has_product(product)

        if found:
            price = self.product_db.get_price(product, size)
            cost = self.product_db.get_cost(product, size)
            self._set_entry(row_data["單價"], price)
            row_data["_base_cost"] = float(cost) if cost else 0
            row_data["_cost_manual"] = False
        else:
            row_data["_base_cost"] = 0
        self._on_qty_changed(row_data)

    def _on_qty_changed(self, row_data):
        price = self._get_num_widget(row_data["單價"])
        qty = self._get_num_widget(row_data["數量"])
        discount = self._get_num_widget(row_data["折扣"])
        total = price * qty
        if discount > 0:
            total = total * discount
        self._set_entry(row_data["總價"], total)
        # 自動模式下更新該行成本
        if not row_data["_cost_manual"]:
            row_cost = row_data["_base_cost"] * (qty or 1)
            self._set_entry(row_data["成本"], row_cost)
        self._recalculate()

    def _on_row_cost_edited(self, row_data):
        """使用者手動編輯某行成本"""
        row_data["_cost_manual"] = True
        self._recalculate()

    # ─── 付款方式連動 ───

    def _on_payment_selected(self, event=None):
        method = self.payment_cb.get()
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

    # ─── 自動計算 ───

    def _recalculate(self):
        grand_total = 0
        total_row_cost = 0
        for rd in self.product_rows:
            price = self._get_num_widget(rd["單價"])
            qty = self._get_num_widget(rd["數量"])
            discount = self._get_num_widget(rd["折扣"])
            line_total = price * qty
            if discount > 0:
                line_total = line_total * discount
            self._set_entry(rd["總價"], line_total)
            grand_total += line_total

            # 每行成本：手動模式讀欄位值，自動模式用 DB 成本
            if rd["_cost_manual"]:
                total_row_cost += self._get_num_widget(rd["成本"])
            else:
                row_cost = rd["_base_cost"] * (qty or 1)
                self._set_entry(rd["成本"], row_cost)
                total_row_cost += row_cost

        special_discount = self._get_num_widget(self.fin_inputs["特殊折扣"])

        # 實拿
        actual = grand_total - special_discount
        self._set_entry(self.fin_inputs["實拿"], actual)

        # 成本 = 各品項成本加總 + 其他成本 + 附加費用
        other_cost = self._get_num_widget(self.other_cost_ent)
        extra = 0
        if self.shipping_var.get():
            extra += round(actual * SHIPPING_RATE)
        if self.card_fee_var.get():
            extra += round(actual * CARD_RATE)
        if self.invoice_var.get():
            extra += round(actual * INVOICE_RATE)

        final_cost = total_row_cost + other_cost + extra
        self._set_entry(self.fin_inputs["成本(品項+贈品)"], final_cost)

        # 利潤
        profit = actual - final_cost
        self._set_entry(self.fin_inputs["利潤"], profit)

    # ─── 儲存 ───

    def _save(self):
        customer = self.shared_inputs["訂購人"].get().strip()
        phone = self.shared_inputs["電話"].get().strip()
        if not customer or not phone:
            messagebox.showwarning("提示", "「訂購人」與「電話」為必填欄位")
            return

        # 收集共用欄位
        shared = {
            "日期": self.shared_inputs["日期"].get().strip(),
            "訂購人": customer,
            "電話": phone,
            "地址": self.shared_inputs["地址"].get().strip(),
            "贈品": self.shared_inputs["贈品"].get().strip(),
            "備註": self.remark_ent.get().strip(),
            "特殊折扣": self.fin_inputs["特殊折扣"].get().strip(),
            "付款方式": self.payment_cb.get().strip(),
            "是否已付": self.paid_cb.get().strip(),
        }

        # 收集每個品項行
        valid_rows = [rd for rd in self.product_rows if rd["品項"].get().strip()]
        if not valid_rows:
            messagebox.showwarning("提示", "至少要有一個品項")
            return

        grand_total = 0
        total_row_cost = 0
        lines = []
        for rd in valid_rows:
            qty = self._get_num_widget(rd["數量"]) or 1
            price = self._get_num_widget(rd["單價"])
            discount = self._get_num_widget(rd["折扣"])
            line_total = price * qty
            if discount > 0:
                line_total = line_total * discount
            row_cost = self._get_num_widget(rd["成本"])
            grand_total += line_total
            total_row_cost += row_cost
            lines.append({
                "品項": rd["品項"].get().strip(),
                "尺寸": rd["尺寸"].get().strip(),
                "數量": str(int(qty)),
                "單價": str(int(price)) if price else "",
                "折扣": str(discount) if discount > 0 else "",
                "總價": str(int(line_total)) if line_total else "",
                "_row_cost": row_cost,
                "_line_total": line_total,
            })

        # 計算財務
        special_discount = self._get_num_widget(self.fin_inputs["特殊折扣"])
        actual = grand_total - special_discount

        other_cost = self._get_num_widget(self.other_cost_ent)
        extra = 0
        if self.shipping_var.get():
            extra += round(actual * SHIPPING_RATE)
        if self.card_fee_var.get():
            extra += round(actual * CARD_RATE)
        if self.invoice_var.get():
            extra += round(actual * INVOICE_RATE)

        final_cost = total_row_cost + other_cost + extra
        profit = actual - final_cost

        # 每行獨立儲存，按比例分配財務數據
        for i, line in enumerate(lines):
            ratio = line["_line_total"] / grand_total if grand_total > 0 else 1 / len(lines)
            row = {**shared, **line}
            row["實拿"] = str(round(actual * ratio))
            row["成本(品項+贈品)"] = str(round(line["_row_cost"] + (other_cost + extra) * ratio))
            row["利潤"] = str(round((actual * ratio) - (line["_row_cost"] + (other_cost + extra) * ratio)))
            # 特殊折扣只記在第一行
            if i > 0:
                row["特殊折扣"] = ""
            # 清理內部欄位
            row.pop("_row_cost", None)
            row.pop("_line_total", None)
            self.on_save(row)

        self._clear_form()
        messagebox.showinfo("成功", f"已儲存 {len(lines)} 筆紀錄")

    def _clear_form(self):
        # 清共用欄位（保留日期）
        for f, e in self.shared_inputs.items():
            if f != "日期":
                e.delete(0, tk.END)

        # 清品項行（保留第一行，刪除其餘）
        while len(self.product_rows) > 1:
            rd = self.product_rows[-1]
            for w in rd["_widgets"]:
                w.destroy()
            self.product_rows.pop()

        # 清第一行
        rd = self.product_rows[0]
        rd["品項"].set('')
        rd["尺寸"].set('')
        rd["數量"].delete(0, tk.END)
        rd["折扣"].delete(0, tk.END)
        rd["單價"].delete(0, tk.END)
        rd["總價"].delete(0, tk.END)
        rd["成本"].delete(0, tk.END)
        rd["_base_cost"] = 0
        rd["_cost_manual"] = False

        # 清財務
        for f, e in self.fin_inputs.items():
            e.delete(0, tk.END)

        self.payment_cb.set('')
        self.paid_cb.set('')
        self.remark_ent.delete(0, tk.END)
        self.other_cost_ent.delete(0, tk.END)
        self.shipping_var.set(False)
        self.card_fee_var.set(False)
        self.invoice_var.set(False)

    # ─── 公開方法 ───

    def refresh_products(self):
        if self.product_db:
            self.product_db.reload()
            products = self.product_db.get_products()
            for rd in self.product_rows:
                rd["品項"]['values'] = products
                rd["_all_products"] = products

    # ─── 輔助 ───

    def _get_num_widget(self, widget):
        try:
            return float(widget.get())
        except (ValueError, TypeError):
            return 0

    def _set_entry(self, widget, value):
        """設定可編輯 Entry 的值（不鎖定）。若該欄位正在被使用者編輯則跳過，避免覆蓋。"""
        try:
            if widget == widget.focus_get():
                return
        except (KeyError, RuntimeError):
            pass
        widget.delete(0, tk.END)
        if value != "" and value is not None:
            try:
                v = float(value)
                v = int(v) if v == int(v) else round(v, 1)
                widget.insert(0, str(v))
            except (ValueError, TypeError):
                widget.insert(0, str(value))

    def _set_readonly(self, widget, value):
        widget.config(state="normal")
        widget.delete(0, tk.END)
        if value != "" and value is not None:
            try:
                v = float(value)
                v = int(v) if v == int(v) else round(v, 1)
                widget.insert(0, str(v))
            except (ValueError, TypeError):
                widget.insert(0, str(value))
        widget.config(state="readonly")

    def _toggle(self, event=None):
        if self.add_visible.get():
            self.add_card.pack_forget()
            self.toggle_btn.config(text="▶ 快速新增紀錄")
        else:
            self.add_card.pack(fill="x", padx=24, pady=(4, 8),
                               after=self.toggle_frame)
            self.toggle_btn.config(text="▼ 快速新增紀錄")
        self.add_visible.set(not self.add_visible.get())
