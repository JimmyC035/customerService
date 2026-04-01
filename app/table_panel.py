import tkinter as tk
from tkinter import ttk, messagebox

from app.constants import COLORS, COLS, COL_WIDTHS


class TablePanel:
    """Treeview 表格顯示 + 搜尋 + 右鍵選單"""

    def __init__(self, parent, get_df, on_delete=None, on_edit=None):
        """
        parent: 父容器
        get_df: callable，回傳目前的 DataFrame
        on_delete: callback(df_index) 刪除紀錄
        on_edit: callback(df_index, new_row_dict) 修改紀錄
        """
        self.parent = parent
        self.get_df = get_df
        self.on_delete = on_delete
        self.on_edit = on_edit

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

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="就緒")
        status_bar = tk.Label(self.parent, textvariable=self.status_var,
                              bg=COLORS["border"], fg=COLORS["text_light"],
                              font=("Arial", 10), anchor="w", padx=12)
        status_bar.pack(fill="x", side="bottom")

    def _build_context_menu(self):
        self.ctx_menu = tk.Menu(self.tree, tearoff=0)
        self.ctx_menu.add_command(label="✏️  修改此筆", command=self._edit_selected)
        self.ctx_menu.add_command(label="🗑  刪除此筆", command=self._delete_selected)

        # macOS 用 Button-2，Windows/Linux 用 Button-3
        self.tree.bind("<Button-2>", self._show_context_menu)
        self.tree.bind("<Button-3>", self._show_context_menu)
        # macOS Control-Click
        self.tree.bind("<Control-Button-1>", self._show_context_menu)

    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
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
                   on_save=lambda new_data: self._on_edit_save(df_idx, new_data))

    def _on_edit_save(self, df_idx, new_data):
        if self.on_edit:
            self.on_edit(df_idx, new_data)

    # ─── 公開方法 ───

    def display(self, target_df):
        for r in self.tree.get_children():
            self.tree.delete(r)

        self._display_indices = []

        if not target_df.empty:
            show_df = target_df.copy()
            show_df["日期"] = show_df["日期"].astype(str).str.split(" ").str[0]
            sorted_df = show_df.sort_values(by="日期", ascending=False)
        else:
            sorted_df = target_df

        for idx, (df_idx, row) in enumerate(sorted_df.fillna("").iterrows()):
            values = [row.get(c, "") for c in COLS]
            tag = ('odd',) if idx % 2 == 1 else ()
            self.tree.insert('', 'end', values=values, tags=tag)
            self._display_indices.append(df_idx)

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

    def __init__(self, parent, row_data, on_save):
        self.on_save = on_save

        self.win = tk.Toplevel(parent)
        self.win.title("修改紀錄")
        self.win.configure(bg=COLORS["card"])
        self.win.geometry("600x500")
        self.win.resizable(False, False)
        self.win.grab_set()

        # 置中
        self.win.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
        self.win.geometry(f"+{x}+{y}")

        tk.Label(self.win, text="修改紀錄", bg=COLORS["card"], fg=COLORS["text"],
                 font=("Arial", 14, "bold")).pack(pady=(16, 8))

        # 用 canvas + scrollbar 以防欄位太多
        canvas_frame = tk.Frame(self.win, bg=COLORS["card"])
        canvas_frame.pack(fill="both", expand=True, padx=20)

        self.entries = {}
        for i, col in enumerate(COLS):
            row_frame = tk.Frame(canvas_frame, bg=COLORS["card"])
            row_frame.pack(fill="x", pady=3)

            tk.Label(row_frame, text=col, width=14, anchor="e",
                     bg=COLORS["card"], fg=COLORS["text_light"],
                     font=("Arial", 10)).pack(side="left", padx=(0, 8))
            ent = tk.Entry(row_frame,
                           bg=COLORS["input_bg"], fg=COLORS["text"],
                           insertbackground=COLORS["text"],
                           font=("Arial", 11), relief="solid",
                           highlightthickness=0, bd=1)
            ent.pack(side="left", fill="x", expand=True, ipady=3)
            ent.insert(0, str(row_data.get(col, "")))
            self.entries[col] = ent

        # 按鈕
        btn_frame = tk.Frame(self.win, bg=COLORS["card"])
        btn_frame.pack(fill="x", padx=20, pady=16)

        tk.Button(btn_frame, text="取消", command=self.win.destroy,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Arial", 11), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)
        tk.Button(btn_frame, text="儲存", command=self._save,
                  bg=COLORS["success"], fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  padx=20, pady=6).pack(side="right", padx=4)

    def _save(self):
        new_data = {col: self.entries[col].get().strip() for col in COLS}
        self.on_save(new_data)
        self.win.destroy()
