import csv
import sqlite3
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox, filedialog

DB_PATH = Path(r"C:\meylis\car-parts-script\car-parts-db.db")
PAGE_SIZE = 500

COLUMNS = (
    "id",
    "article",
    "brand",
    "weight_kg",
    "product_name",
    "source_file",
    "inserted_at",
)

COLUMN_TITLES = {
    "id": "ID",
    "article": "Артикул",
    "brand": "Марка",
    "weight_kg": "Вес, кг",
    "product_name": "Наименование",
    "source_file": "Имя файла",
    "inserted_at": "Дата и время",
}

COLUMN_WIDTHS = {
    "id": 70,
    "article": 150,
    "brand": 140,
    "weight_kg": 90,
    "product_name": 420,
    "source_file": 220,
    "inserted_at": 190,
}

DATASETS = {
    "Последние записи (latest_parts)": "latest_parts",
    "Вся история (parts)": "parts",
}


def open_readonly_connection():
    if not DB_PATH.is_file():
        raise FileNotFoundError(f"База данных не найдена:\n{DB_PATH}")

    # SQLite URI mode=ro гарантирует, что viewer ничего не изменит в базе.
    uri = DB_PATH.resolve().as_uri() + "?mode=ro"
    con = sqlite3.connect(uri, uri=True, timeout=10)
    con.row_factory = sqlite3.Row
    return con


class PartsDBViewer(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Parts DB Viewer")
        self.geometry("1450x820")
        self.minsize(1000, 600)

        self.current_page = 0
        self.sort_column = "id"
        self.sort_desc = True
        self.total_rows = 0
        self.last_rows = []

        self._build_ui()
        self.after(100, self.initial_load)

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.data_tab = ttk.Frame(notebook)
        self.sql_tab = ttk.Frame(notebook)

        notebook.add(self.data_tab, text="Просмотр базы")
        notebook.add(self.sql_tab, text="SQL SELECT")

        self._build_data_tab()
        self._build_sql_tab()

    def _build_data_tab(self):
        top = ttk.Frame(self.data_tab, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Раздел:").grid(row=0, column=0, padx=(0, 5), pady=3, sticky="w")

        self.dataset_var = tk.StringVar(value="Последние записи (latest_parts)")
        dataset_box = ttk.Combobox(
            top,
            textvariable=self.dataset_var,
            values=list(DATASETS.keys()),
            state="readonly",
            width=31,
        )
        dataset_box.grid(row=0, column=1, padx=(0, 15), pady=3, sticky="w")
        dataset_box.bind("<<ComboboxSelected>>", lambda event: self.reset_and_load())

        ttk.Label(top, text="Артикул:").grid(row=0, column=2, padx=(0, 5), pady=3)
        self.article_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.article_var, width=20).grid(
            row=0, column=3, padx=(0, 15), pady=3
        )

        ttk.Label(top, text="Марка:").grid(row=0, column=4, padx=(0, 5), pady=3)
        self.brand_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.brand_var, width=18).grid(
            row=0, column=5, padx=(0, 15), pady=3
        )

        ttk.Label(top, text="Наименование:").grid(row=0, column=6, padx=(0, 5), pady=3)
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var, width=28).grid(
            row=0, column=7, padx=(0, 15), pady=3
        )

        ttk.Button(top, text="Найти", command=self.reset_and_load).grid(
            row=0, column=8, padx=4, pady=3
        )
        ttk.Button(top, text="Сбросить", command=self.clear_filters).grid(
            row=0, column=9, padx=4, pady=3
        )
        ttk.Button(top, text="Обновить", command=self.load_page).grid(
            row=0, column=10, padx=4, pady=3
        )
        ttk.Button(top, text="Экспорт CSV", command=self.export_current_query).grid(
            row=0, column=11, padx=4, pady=3
        )

        # Enter запускает поиск.
        for var in (self.article_var, self.brand_var, self.name_var):
            var.trace_add("write", lambda *_: None)

        table_frame = ttk.Frame(self.data_tab, padding=(8, 0, 8, 0))
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            table_frame,
            columns=COLUMNS,
            show="headings",
            selectmode="browse",
        )

        for col in COLUMNS:
            self.tree.heading(
                col,
                text=COLUMN_TITLES[col],
                command=lambda c=col: self.change_sort(c),
            )
            anchor = "e" if col in ("id", "weight_kg") else "w"
            self.tree.column(
                col,
                width=COLUMN_WIDTHS[col],
                minwidth=50,
                anchor=anchor,
                stretch=True,
            )

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.show_selected_record)

        bottom = ttk.Frame(self.data_tab, padding=8)
        bottom.pack(fill="x")

        self.prev_button = ttk.Button(bottom, text="← Назад", command=self.prev_page)
        self.prev_button.pack(side="left")

        self.next_button = ttk.Button(bottom, text="Вперед →", command=self.next_page)
        self.next_button.pack(side="left", padx=(5, 15))

        self.page_label = ttk.Label(bottom, text="")
        self.page_label.pack(side="left")

        self.status_label = ttk.Label(bottom, text="")
        self.status_label.pack(side="right")

    def _build_sql_tab(self):
        top = ttk.Frame(self.sql_tab, padding=8)
        top.pack(fill="x")

        ttk.Label(
            top,
            text=(
                "Здесь можно выполнять только SELECT / WITH. "
                "Соединение открывается только для чтения."
            ),
        ).pack(side="left")

        ttk.Button(top, text="Выполнить", command=self.run_sql).pack(side="right")
        ttk.Button(top, text="Пример", command=self.put_sql_example).pack(
            side="right", padx=(0, 5)
        )

        sql_frame = ttk.Frame(self.sql_tab, padding=(8, 0, 8, 8))
        sql_frame.pack(fill="x")

        self.sql_text = tk.Text(sql_frame, height=8, wrap="none")
        self.sql_text.pack(fill="x")
        self.sql_text.insert(
            "1.0",
            "SELECT *\n"
            "FROM latest_parts\n"
            "ORDER BY id DESC\n"
            "LIMIT 100;"
        )

        result_frame = ttk.Frame(self.sql_tab, padding=(8, 0, 8, 8))
        result_frame.pack(fill="both", expand=True)

        self.sql_tree = ttk.Treeview(result_frame, show="headings")
        sql_y = ttk.Scrollbar(result_frame, orient="vertical", command=self.sql_tree.yview)
        sql_x = ttk.Scrollbar(result_frame, orient="horizontal", command=self.sql_tree.xview)

        self.sql_tree.configure(yscrollcommand=sql_y.set, xscrollcommand=sql_x.set)

        self.sql_tree.grid(row=0, column=0, sticky="nsew")
        sql_y.grid(row=0, column=1, sticky="ns")
        sql_x.grid(row=1, column=0, sticky="ew")

        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        self.sql_status = ttk.Label(self.sql_tab, padding=(8, 0, 8, 8), text="")
        self.sql_status.pack(fill="x")

    def initial_load(self):
        try:
            self.load_page()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def clear_filters(self):
        self.article_var.set("")
        self.brand_var.set("")
        self.name_var.set("")
        self.reset_and_load()

    def reset_and_load(self):
        self.current_page = 0
        self.load_page()

    def get_dataset(self):
        return DATASETS[self.dataset_var.get()]

    def build_where(self):
        clauses = []
        params = []

        article = self.article_var.get().strip()
        brand = self.brand_var.get().strip()
        product_name = self.name_var.get().strip()

        if article:
            clauses.append("article LIKE ?")
            params.append(f"%{article}%")

        if brand:
            clauses.append("brand LIKE ?")
            params.append(f"%{brand}%")

        if product_name:
            clauses.append("product_name LIKE ?")
            params.append(f"%{product_name}%")

        where_sql = ""
        if clauses:
            where_sql = " WHERE " + " AND ".join(clauses)

        return where_sql, params

    def load_page(self):
        dataset = self.get_dataset()
        where_sql, params = self.build_where()

        # Имя таблицы и имя сортируемого столбца берутся только из белых списков.
        if dataset not in DATASETS.values():
            raise ValueError("Недопустимое имя таблицы.")
        if self.sort_column not in COLUMNS:
            raise ValueError("Недопустимый столбец сортировки.")

        order = "DESC" if self.sort_desc else "ASC"
        offset = self.current_page * PAGE_SIZE

        with open_readonly_connection() as con:
            count_sql = f"SELECT COUNT(*) FROM {dataset}{where_sql}"
            self.total_rows = con.execute(count_sql, params).fetchone()[0]

            sql = (
                f"SELECT {', '.join(COLUMNS)} "
                f"FROM {dataset}"
                f"{where_sql} "
                f"ORDER BY {self.sort_column} {order} "
                f"LIMIT ? OFFSET ?"
            )

            rows = con.execute(sql, [*params, PAGE_SIZE, offset]).fetchall()

        self.last_rows = [dict(row) for row in rows]

        self.tree.delete(*self.tree.get_children())

        for row in rows:
            values = []
            for col in COLUMNS:
                value = row[col]
                if value is None:
                    value = ""
                values.append(value)

            self.tree.insert("", "end", values=values)

        max_page = max(1, (self.total_rows + PAGE_SIZE - 1) // PAGE_SIZE)
        shown_page = min(self.current_page + 1, max_page)

        self.page_label.config(
            text=f"Страница {shown_page} из {max_page}  |  по {PAGE_SIZE} строк"
        )
        self.status_label.config(
            text=f"Найдено записей: {self.total_rows}"
        )

        self.prev_button.config(state="normal" if self.current_page > 0 else "disabled")
        self.next_button.config(
            state="normal"
            if (self.current_page + 1) * PAGE_SIZE < self.total_rows
            else "disabled"
        )

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.load_page()

    def next_page(self):
        if (self.current_page + 1) * PAGE_SIZE < self.total_rows:
            self.current_page += 1
            self.load_page()

    def change_sort(self, column):
        if column not in COLUMNS:
            return

        if self.sort_column == column:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_column = column
            self.sort_desc = False

        self.current_page = 0
        self.load_page()

    def show_selected_record(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return

        values = self.tree.item(selected[0], "values")
        if not values:
            return

        win = tk.Toplevel(self)
        win.title("Запись")
        win.geometry("800x430")

        frame = ttk.Frame(win, padding=12)
        frame.pack(fill="both", expand=True)

        for idx, col in enumerate(COLUMNS):
            ttk.Label(
                frame,
                text=COLUMN_TITLES[col] + ":",
                width=16,
            ).grid(row=idx, column=0, sticky="nw", padx=(0, 10), pady=4)

            value = values[idx] if idx < len(values) else ""

            if col == "product_name":
                box = tk.Text(frame, height=5, wrap="word")
                box.insert("1.0", value)
                box.config(state="disabled")
                box.grid(row=idx, column=1, sticky="nsew", pady=4)
            else:
                entry = ttk.Entry(frame)
                entry.insert(0, value)
                entry.config(state="readonly")
                entry.grid(row=idx, column=1, sticky="ew", pady=4)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(COLUMNS.index("product_name"), weight=1)

    def export_current_query(self):
        dataset = self.get_dataset()
        where_sql, params = self.build_where()

        if dataset not in DATASETS.values():
            return

        order = "DESC" if self.sort_desc else "ASC"

        path = filedialog.asksaveasfilename(
            title="Сохранить CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        with open_readonly_connection() as con:
            sql = (
                f"SELECT {', '.join(COLUMNS)} "
                f"FROM {dataset}"
                f"{where_sql} "
                f"ORDER BY {self.sort_column} {order}"
            )
            rows = con.execute(sql, params).fetchall()

        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow([COLUMN_TITLES[c] for c in COLUMNS])

            for row in rows:
                writer.writerow([
                    "" if row[c] is None else row[c]
                    for c in COLUMNS
                ])

        messagebox.showinfo(
            "Экспорт",
            f"Сохранено строк: {len(rows)}\n\n{path}"
        )

    def put_sql_example(self):
        example = (
            "SELECT article, brand, weight_kg, product_name, source_file, inserted_at\n"
            "FROM parts\n"
            "WHERE article = '3075801'\n"
            "ORDER BY id DESC;"
        )
        self.sql_text.delete("1.0", "end")
        self.sql_text.insert("1.0", example)

    def run_sql(self):
        sql = self.sql_text.get("1.0", "end").strip()

        if not sql:
            return

        # Убираем ведущие SQL-комментарии для простой проверки команды.
        check_sql = sql.lstrip()

        # Разрешаем только чтение.
        upper = check_sql.upper()
        if not (upper.startswith("SELECT") or upper.startswith("WITH")):
            messagebox.showwarning(
                "Только чтение",
                "Viewer разрешает выполнять только SELECT или WITH."
            )
            return

        try:
            with open_readonly_connection() as con:
                cur = con.execute(sql)
                rows = cur.fetchall()
                description = cur.description or []

            columns = [item[0] for item in description]

            self.sql_tree.delete(*self.sql_tree.get_children())
            self.sql_tree["columns"] = columns

            for col in columns:
                self.sql_tree.heading(col, text=col)
                self.sql_tree.column(col, width=150, anchor="w")

            # Ограничиваем только отображение, но сам SELECT выполняется целиком.
            display_rows = rows[:5000]

            for row in display_rows:
                self.sql_tree.insert(
                    "",
                    "end",
                    values=[
                        "" if value is None else value
                        for value in row
                    ],
                )

            suffix = ""
            if len(rows) > len(display_rows):
                suffix = f" (показаны первые {len(display_rows)})"

            self.sql_status.config(
                text=f"Получено строк: {len(rows)}{suffix}"
            )

        except Exception as exc:
            messagebox.showerror("SQL ошибка", str(exc))


def main():
    try:
        app = PartsDBViewer()
        app.mainloop()
    except Exception as exc:
        try:
            messagebox.showerror("Ошибка запуска", str(exc))
        except Exception:
            print(f"Ошибка запуска: {exc}")


if __name__ == "__main__":
    main()
