import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import pandas as pd
import difflib

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
                "Количество": self.stock_df[col_count]
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

            name_account = self.invoice_df["Товар"]
            quantity_account = self.invoice_df["Количество"]

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
        inv = self.invoice_df.copy()
        for df in (stock, inv):
            if QTY_COL not in df.columns:
                df[QTY_COL] = 1

        result_rows = []

        for _, row in inv.iterrows():
            product = str(row[PRODUCT_COL]).strip()
            req_qty = float(row[QTY_COL])
            available = stock.loc[stock[PRODUCT_COL] == product, QTY_COL].sum()

            take_qty = min(available, req_qty)
            shortfall = req_qty - take_qty
            alt_list = []

            if take_qty > 0:
                alt_list.append({"Product": product, "Taken": take_qty})
                idx_exact = stock[stock[PRODUCT_COL] == product].index
                stock.loc[idx_exact, QTY_COL] -= take_qty

            if shortfall > 0:
                candidates = stock[stock[QTY_COL] > 0]
                mask = candidates[PRODUCT_COL].str.contains(product.split()[0], case=False, na=False)
                alt_candidates = candidates[mask]

                if alt_candidates.empty:
                    ratios = candidates[PRODUCT_COL].apply(
                        lambda x: difflib.SequenceMatcher(None, product, str(x)).ratio())
                    top_matches = ratios.nlargest(5).index
                    alt_candidates = candidates.loc[top_matches]

                for _, alt in alt_candidates.iterrows():
                    if shortfall <= 0:
                        break
                    alt_available = alt[QTY_COL]
                    take_alt = min(alt_available, shortfall)
                    alt_list.append({"Product": alt[PRODUCT_COL], "Taken": take_alt})
                    shortfall -= take_alt
                    stock.loc[alt.name, QTY_COL] -= take_alt

            result_rows.append({
                "Исходный товар": product,
                "Запрошено": req_qty,
                "Подобранное позиционирование": "; ".join(f"{a['Product']} x{a['Taken']}" for a in alt_list),
                "Итого отгружено": sum(a["Taken"] for a in alt_list),
                "Не закрыто": max(0, req_qty - sum(a["Taken"] for a in alt_list))
            })

            self.log_write(
                f"{product}: нужно {req_qty} | отдали {req_qty - shortfall} | осталось закрыть {shortfall}\n")

        self.result_df = pd.DataFrame(result_rows)
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
