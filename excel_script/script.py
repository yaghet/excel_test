import tkinter as tk

from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook


class ExcelFilterApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel Filter")

        self.btn_open = tk.Button(master, text="Выбрать исходный Excel файл", command=self.open_file)
        self.btn_open.pack(padx=10, pady=5)

        self.label_filter_col = tk.Label(master, text="Столбец для фильтра (название):")
        self.label_filter_col.pack(padx=10, pady=5)

        self.entry_filter_col = tk.Entry(master)
        self.entry_filter_col.pack(padx=10, pady=5)

        self.label_filter_val = tk.Label(master, text="Значение фильтра:")
        self.label_filter_val.pack(padx=10, pady=5)

        self.entry_filter_val = tk.Entry(master)
        self.entry_filter_val.pack(padx=10, pady=5)

        self.btn_save = tk.Button(master, text="Сохранить в новый Excel файл", command=self.save_file)
        self.btn_save.pack(padx=10, pady=10)

        self.source_path = None
        self.data = None
        self.headers = None

        self.header_row_number = 11

    def open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Выберите Excel файл"
        )
        if not path:
            return

        try:
            wb = load_workbook(path, read_only=True)
            sheet = wb.active

            rows = sheet.iter_rows(values_only=True)

            for _ in range(8):
                next(rows)

            header_row_1 = next(rows)
            header_row_2 = next(rows)

            self.headers = []
            max_len = max(len(header_row_1), len(header_row_2))
            for i in range(max_len):
                part1 = header_row_1[i] if i < len(header_row_1) and header_row_1[i] is not None else ""
                part2 = header_row_2[i] if i < len(header_row_2) and header_row_2[i] is not None else ""
                combined = (str(part1).strip() + " " + str(part2).strip()).strip()
                self.headers.append(combined)

            wb.close()
            self.source_path = path
            messagebox.showinfo("Файл открыт", f"Файл успешно открыт. В нем столбцы:\n{', '.join(self.headers)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def save_file(self):
        if not self.source_path:
            messagebox.showwarning("Нет файла", "Сначала выберите исходный файл.")
            return

        filter_col = self.entry_filter_col.get().strip()
        filter_val = self.entry_filter_val.get().strip()

        if not filter_col or not filter_val:
            messagebox.showwarning("Ввод", "Введите имя столбца и значение для фильтрации.")
            return

        if filter_col not in self.headers:
            matched_cols = [h for h in self.headers if h.lower() == filter_col.lower()]
            if not matched_cols:
                messagebox.showerror("Ошибка", f"Столбец '{filter_col}' не найден в файле.")
                return
            filter_col = matched_cols[0]

        try:
            wb = load_workbook(self.source_path, read_only=True)
            sheet = wb.active

            rows = sheet.iter_rows(values_only=True)
            for _ in range(10):
                next(rows)

            needed_cols = ["ФИО", "Должность", "Отдел", "Дата найма", "Зарплата"]
            missing = [c for c in needed_cols if
                       c not in self.headers and c.lower() not in [h.lower() for h in self.headers]]
            if missing:
                messagebox.showerror("Ошибка", f"В исходном файле отсутствуют столбцы:\n{', '.join(missing)}")
                return

            def col_index(col_name):
                for i, h in enumerate(self.headers):
                    if h == col_name:
                        return i
                for i, h in enumerate(self.headers):
                    if h.lower() == col_name.lower():
                        return i
                return -1

            filter_col_idx = col_index(filter_col)
            needed_col_idxs = [col_index(c) for c in needed_cols]

            wb_out = Workbook()
            ws_out = wb_out.active
            ws_out.title = "Отфильтрованные данные"
            ws_out.append(needed_cols)

            rows_iter = sheet.iter_rows(values_only=True)
            next(rows_iter)

            for row in rows_iter:
                cell_val = row[filter_col_idx]
                if cell_val is None:
                    continue
                if str(cell_val).strip().lower() == filter_val.lower():
                    filtered_row = [row[i] if 0 <= i < len(row) else None for i in needed_col_idxs]
                    ws_out.append(filtered_row)

            wb.close()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обработке файла:\n{e}")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить отфильтрованный файл"
        )
        if not save_path:
            return

        try:
            wb_out.save(save_path)
            messagebox.showinfo("Успех", f"Файл успешно сохранён: {save_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
