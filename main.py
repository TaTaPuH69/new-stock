import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import pandas as pd
import difflib
import re

def numeric_clean(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
              .str.replace(r"\s+", "", regex=True)
              .str.replace(",", ".", regex=False)
              .pipe(pd.to_numeric, errors="coerce")
              .fillna(0)
    )

PRODUCT_COL = "Товар"
QTY_COL = "Количество"

class StockMatcherApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        master.title("Codex Stock Matcher")
        master.geometry("800x500")

        # Данные
        self.stock_df: pd.DataFrame | None = None
        self.invoice_df: pd.DataFrame | None = None
        self.result_df: pd.DataFrame | None = None
        self.invoice_path: Path | None = None

        # --- UI ---
        frm = ttk.Frame(master, padding=20)
        frm.pack(fill="both", expand=True)

        ttk.Button(frm, text="Загрузить остатки", command=self.load_stock).grid(row=0, column=0, sticky="ew", pady=5)
        ttk.Button(frm, text="Загрузить счёт",   command=self.load_invoice).grid(row=1, column=0, sticky="ew", pady=5)
        ttk.Button(frm, text="Сохранить новый счёт", command=self.save_result).grid(row=2, column=0, sticky="ew", pady=5)

        # Text‑area для логов
        self.log = tk.Text(frm, height=20)
        self.log.grid(row=0, column=1, rowspan=10, sticky="nsew", padx=(15,0))
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(9, weight=1)

    # ---------- Загрузка файлов ----------
    def load_stock(self):
        path = filedialog.askopenfilename(
            title="Оборотно-сальдовая ведомость",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path: return
        try:
            self.stock_df = pd.read_excel(path, header=7)
            col_name = self.stock_df.columns[0]
            col_count = self.stock_df.columns[1]

            clean_df = pd.DataFrame({
                "Товар": self.stock_df[col_name],
                "Количество": numeric_clean(self.stock_df[col_count]),
            })

            self.stock_df = clean_df
            self.log_write(f"✅ Остатки загружены: {Path(path).name} | {len(self.stock_df)} строк\n")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def load_invoice(self):
        path = filedialog.askopenfilename(
            title="Счёт",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path: return
        try:
            self.invoice_path = Path(path)
            self.invoice_df = pd.read_excel(path, header=16, nrows=10)

            self.invoice_df[QTY_COL] = numeric_clean(self.invoice_df[QTY_COL])

            self.log_write(f"✅ Счёт загружен: {self.invoice_path.name} | {len(self.invoice_df)} строк\n")
            self.process()
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    # ---------- Логика подбора ----------
    def process(self):
        if self.stock_df is None or self.invoice_df is None:
            messagebox.showwarning("Нет данных", "Сначала загрузи и остатки, и счёт.")
            return

        stock = self.stock_df.copy()
        invoice = self.invoice_df.copy()

        for df in (stock, invoice):
            if QTY_COL not in df.columns:
                df[QTY_COL] = 1

        taken_rows = []  # Сюда будем добавлять все взятые позиции

        for _, row in invoice.iterrows():
            product = str(row[PRODUCT_COL]).strip()
            need_qty = row[QTY_COL]

            self.log_write(f"{product}: требуется {need_qty}\n")

            # --- сначала пробуем точное совпадение ---
            mask_exact = stock[PRODUCT_COL] == product
            available = stock.loc[mask_exact, QTY_COL].sum()
            take_qty = min(available, need_qty)

            if take_qty > 0:
                taken_rows.append({PRODUCT_COL: product, QTY_COL: take_qty})
                stock.loc[mask_exact, QTY_COL] -= take_qty
                need_qty -= take_qty
                self.log_write(f"  - взяли {take_qty} с точным совпадением\n")

            # --- ищем похожие позиции, если не хватило ---
            if need_qty > 0:
                candidates = stock[stock[QTY_COL] > 0]
                mask = candidates[PRODUCT_COL].str.contains(product.split()[0], case=False, na=False)
                alt_candidates = candidates[mask]

                if alt_candidates.empty:
                    ratios = candidates[PRODUCT_COL].apply(
                        lambda x: difflib.SequenceMatcher(None, product, str(x)).ratio())
                    top_matches = ratios.nlargest(5).index
                    alt_candidates = candidates.loc[top_matches]

                for idx, alt in alt_candidates.iterrows():
                    if need_qty <= 0:
                        break
                    alt_available = alt[QTY_COL]
                    take_alt = min(alt_available, need_qty)
                    taken_rows.append({PRODUCT_COL: alt[PRODUCT_COL], QTY_COL: take_alt})
                    stock.loc[idx, QTY_COL] -= take_alt
                    need_qty -= take_alt
                    self.log_write(f"  - взяли {take_alt} из '{alt[PRODUCT_COL]}'\n")

            if need_qty > 0:
                self.log_write(f"  - не удалось закрыть {need_qty} единиц\n")

        # Группируем одинаковые товары для финального счёта
        if taken_rows:
            result = pd.DataFrame(taken_rows).groupby(PRODUCT_COL, as_index=False)[QTY_COL].sum()
        else:
            result = pd.DataFrame(columns=[PRODUCT_COL, QTY_COL])

        self.result_df = result
        self.stock_df = stock  # обновляем остатки после отбора
        self.log_write("=== Подбор завершён ===\n")

    # ---------- Сохранение ----------
    def save_result(self):
        if self.result_df is None:
            messagebox.showinfo("Нет данных", "Сначала загрузи и обработай счёт.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Сохранить новый счёт",
            initialfile=f"{self.invoice_path.stem}_new.xlsx" if self.invoice_path else "result.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not save_path: return
        try:
            self.result_df.to_excel(save_path, index=False)
            messagebox.showinfo("Сохранено", f"Новый счёт сохранён: {Path(save_path).name}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

    # ---------- helper ----------
    def log_write(self, msg: str):
        self.log.insert(tk.END, msg)
        self.log.see(tk.END)

# --- Точка входа ---
if __name__ == "__main__":
    root = tk.Tk()
    app = StockMatcherApp(root)
    root.mainloop()
