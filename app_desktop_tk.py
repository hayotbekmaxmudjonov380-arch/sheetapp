import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from datetime import datetime

DB_PATH = "restaurant.db"

class RestaurantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Restoran Hisob Dasturi (MVP)")
        self.root.geometry("1000 = 650" if False else "1050x700")
        self.root.minsize(900, 600)
        
        # Style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook.Tab", font=("Arial", 11, "bold"), padding=[12, 8])
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 10), rowheight=25)
        
        # Main layout - Notebook (Tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Tabs
        self.tab_employees = ttk.Frame(self.notebook)
        self.tab_salary = ttk.Frame(self.notebook)
        self.tab_transactions = ttk.Frame(self.notebook)
        self.tab_balance = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_employees, text="👥 Xodimlar")
        self.notebook.add(self.tab_salary, text="💰 Ish haqi")
        self.notebook.add(self.tab_transactions, text="📥 Kirim / Chiqim")
        self.notebook.add(self.tab_balance, text="📊 Balans va DDS")
        
        self.init_employees_tab()
        self.init_salary_tab()
        self.init_transactions_tab()
        self.init_balance_tab()
        
        # Bottom status / export bar
        bottom_frame = ttk.Frame(root)
        bottom_frame.pack(fill="x", padx=10, pady=5)
        
        btn_export = ttk.Button(bottom_frame, text="📥 Excel'ga Export qilish", command=self.export_to_excel)
        btn_export.pack(side="left", padx=5)
        
        btn_refresh = ttk.Button(bottom_frame, text="🔄 Ma'lumotlarni yangilash", command=self.refresh_all)
        btn_refresh.pack(side="left", padx=5)

    def get_db_connection(self):
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    # --- 1. XODIMLAR TAB ---
    def init_employees_tab(self):
        frame = self.tab_employees
        
        top_frame = ttk.Frame(frame, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="Qidirish (Ism bo'yicha):", font=("Arial", 10)).pack(side="left", padx=5)
        self.emp_search_var = tk.StringVar()
        search_entry = ttk.Entry(top_frame, textvariable=self.emp_search_var, width=25, font=("Arial", 10))
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", self.filter_employees)
        
        btn_add = ttk.Button(top_frame, text="➕ Yangi xodim qo'shish", command=self.add_employee_dialog)
        btn_add.pack(side="right", padx=5)
        
        # Treeview
        tree_frame = ttk.Frame(frame, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("ID", "Ism familiya", "Lavozimi", "Telefon", "Tug'ilgan kun", "Status")
        self.emp_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            self.emp_tree.heading(col, text=col)
            self.emp_tree.column(col, width=150, anchor="w")
        self.emp_tree.column("ID", width=50, anchor="center")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.emp_tree.yview)
        self.emp_tree.configure(yscrollcommand=scrollbar.set)
        
        self.emp_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.load_employees()

    def load_employees(self, search=""):
        for row in self.emp_tree.get_children():
            self.emp_tree.delete(row)
            
        conn = self.get_db_connection()
        cursor = conn.cursor()
        if search:
            cursor.execute("SELECT * FROM employees WHERE name LIKE ?", ('%' + search + '%',))
        else:
            cursor.execute("SELECT * FROM employees")
        
        for row in cursor.fetchall():
            self.emp_tree.insert("", "end", values=(row["id"], row["name"], row["position"], row["phone"], row["birth_date"], row["status"]))
        conn.close()

    def filter_employees(self, event):
        query = self.emp_search_var.get()
        self.load_employees(query)

    def add_employee_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Yangi xodim qo'shish")
        dialog.geometry("400x300")
        dialog.grab_set()
        
        ttk.Label(dialog, text="Ism familiya:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_name = ttk.Entry(dialog, width=35)
        e_name.pack(padx=20)
        
        ttk.Label(dialog, text="Lavozimi:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_pos = ttk.Entry(dialog, width=35)
        e_pos.pack(padx=20)
        
        ttk.Label(dialog, text="Telefon:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_phone = ttk.Entry(dialog, width=35)
        e_phone.pack(padx=20)
        
        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showerror("Xatolik", "Ism kiritilishi shart!", parent=dialog)
                return
            try:
                conn = self.get_db_connection()
                conn.execute("INSERT INTO employees (name, position, phone, status) VALUES (?, ?, ?, 'ishda')",
                             (name, e_pos.get().strip(), e_phone.get().strip()))
                conn.commit()
                conn.close()
                messagebox.SUCCESS = messagebox.showinfo("Muvaffaqiyat", "Xodim qo'shildi!", parent=dialog)
                dialog.destroy()
                self.load_employees()
            except Exception as e:
                messagebox.showerror("Xatolik", f"Xatolik yuz berdi: {e}", parent=dialog)
                
        ttk.Button(dialog, text="Saqlash", command=save).pack(pady=20)

    # --- 2. ISH HAQI TAB ---
    def init_salary_tab(self):
        frame = self.tab_salary
        
        top_frame = ttk.Frame(frame, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="Xodimning ish haqi yozuvlari", font=("Arial", 12, "bold")).pack(side="left", padx=5)
        
        tree_frame = ttk.Frame(frame, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("ID", "Xodim", "Sana", "Kirish", "Chiqish", "Kunbay summa", "Ishlagan vaqti")
        self.salary_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            self.salary_tree.heading(col, text=col)
            self.salary_tree.column(col, width=130, anchor="w")
        self.salary_tree.column("ID", width=50, anchor="center")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.salary_tree.yview)
        self.salary_tree.configure(yscrollcommand=scrollbar.set)
        
        self.salary_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.load_salary()

    def load_salary(self):
        for row in self.salary_tree.get_children():
            self.salary_tree.delete(row)
            
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM work_records ORDER BY work_date DESC")
        for row in cursor.fetchall():
            self.salary_tree.insert("", "end", values=(row["id"], row["employee_name"], row["work_date"], row["check_in"], row["check_out"], f"{row['amount']:,.0f}", row["worked_hours"]))
        conn.close()

    # --- 3. KIRIM / CHIQIM TAB ---
    def init_transactions_tab(self):
        frame = self.tab_transactions
        
        top_frame = ttk.Frame(frame, padding=10)
        top_frame.pack(fill="x")
        
        btn_add_tx = ttk.Button(top_frame, text="➕ Kirim / Chiqim qo'shish", command=self.add_transaction_dialog)
        btn_add_tx.pack(side="left", padx=5)
        
        tree_frame = ttk.Frame(frame, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("ID", "Turi", "Modda / Kategoriya", "Kontragent", "Summa", "Sana")
        self.tx_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            self.tx_tree.heading(col, text=col)
            self.tx_tree.column(col, width=150, anchor="w")
        self.tx_tree.column("ID", width=50, anchor="center")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tx_tree.yview)
        self.tx_tree.configure(yscrollcommand=scrollbar.set)
        
        self.tx_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.load_transactions()

    def load_transactions(self):
        for row in self.tx_tree.get_children():
            self.tx_tree.delete(row)
            
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM transactions ORDER BY tx_date DESC")
        for row in cursor.fetchall():
            t_type = "Kirim 🟢" if row["type"] == "kirim" else "Chiqim 🔴"
            self.tx_tree.insert("", "end", values=(row["id"], t_type, row["category"], row["kontragent"], f"{row['amount']:,.0f}", row["tx_date"]))
        conn.close()

    def add_transaction_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Yangi Kirim / Chiqim qo'shish")
        dialog.geometry("450x380")
        dialog.grab_set()
        
        ttk.Label(dialog, text="Operatsiya turi:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        type_var = tk.StringVar(value="kirim")
        f_type = ttk.Frame(dialog)
        f_type.pack(anchor="w", padx=20)
        ttk.Radiobutton(f_type, text="Kirim", variable=type_var, value="kirim").pack(side="left", padx=5)
        ttk.Radiobutton(f_type, text="Chiqim", variable=type_var, value="chiqim").pack(side="left", padx=5)
        
        ttk.Label(dialog, text="Modda / Kategoriya (masalan: Savdo, Tovar, Arenda...):", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_cat = ttk.Entry(dialog, width=40)
        e_cat.pack(padx=20)
        
        ttk.Label(dialog, text="Kontragent:", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_kontr = ttk.Entry(dialog, width=40)
        e_kontr.pack(padx=20)
        
        ttk.Label(dialog, text="Summa (so'm):", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_amt = ttk.Entry(dialog, width=40)
        e_amt.pack(padx=20)
        
        ttk.Label(dialog, text="Sana (YYYY-MM-DD):", font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        e_date = ttk.Entry(dialog, width=40)
        e_date.insert(0, datetime.today().strftime('%Y-%m-%d'))
        e_date.pack(padx=20)
        
        def save():
            try:
                t_type = type_var.get()
                cat = e_cat.get().strip()
                kontr = e_kontr.get().strip()
                amt = float(e_amt.get().strip())
                dt = e_date.get().strip()
                
                if not cat or not dt:
                    messagebox.showerror("Xatolik", "Kategoriya va sana kiritilishi shart!", parent=dialog)
                    return
                    
                conn = self.get_db_connection()
                conn.execute("INSERT INTO transactions (type, category, kontragent, amount, tx_date) VALUES (?, ?, ?, ?, ?)",
                             (t_type, cat, kontr, amt, dt))
                conn.commit()
                conn.close()
                messagebox.showinfo("Muvaffaqiyat", "Operatsiya qo'shildi!", parent=dialog)
                dialog.destroy()
                self.load_transactions()
                self.load_balance()
            except ValueError:
                messagebox.showerror("Xatolik", "Summa faqat raqam bo'lishi kerak!", parent=dialog)
            except Exception as e:
                messagebox.showerror("Xatolik", f"Xatolik: {e}", parent=dialog)
                
        ttk.Button(dialog, text="Saqlash", command=save).pack(pady=20)

    # --- 4. BALANS VA DDS TAB ---
    def init_balance_tab(self):
        frame = self.tab_balance
        
        top_frame = ttk.Frame(frame, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="Restoran Balansi va Moliyaviy Hisobot", font=("Arial", 13, "bold")).pack(side="left", padx=5)
        
        btn_calc = ttk.Button(top_frame, text="🔄 Hisoblash", command=self.load_balance)
        btn_calc.pack(side="right", padx=5)
        
        # Balance details frame
        self.bal_frame = ttk.LabelFrame(frame, text="Oylik Balans Hisoboti", padding=15)
        self.bal_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        self.bal_labels = {}
        items = [
            ("Savdo", 0),
            ("Sotib olingan tovar", 0),
            ("Yalpi foyda (Savdo - Tovar)", 0),
            ("Ish haqi", 0),
            ("Kommunal", 0),
            ("Arenda", 0),
            ("Marketing", 0),
            ("Remont", 0),
            ("Nalog (Savdoning 2%)", 0),
            ("Jami xarajat", 0),
            ("Sof foyda", 0)
        ]
        
        for idx, (label, val) in enumerate(items):
            lbl_name = ttk.Label(self.bal_frame, text=label + ":", font=("Arial", 11, "bold" if "foyda" in label.lower() or "savdo" in label.lower() else "normal"))
            lbl_name.grid(row=idx, column=0, sticky="w", pady=4)
            
            lbl_val = ttk.Label(self.bal_frame, text="0 so'm", font=("Arial", 11, "bold" if "foyda" in label.lower() else "normal"))
            lbl_val.grid(row=idx, column=1, sticky="e", pady=4, padx=50)
            self.bal_labels[label] = lbl_val
            
        self.load_balance()

    def load_balance(self):
        conn = self.get_db_connection()
        cursor = conn.cursor()
        
        # Savdo
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='kirim' AND (category LIKE '%savdo%' OR category LIKE '%Савдо%')")
        savdo = cursor.fetchone()[0] or 0.0
        
        # Tovar
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='chiqim' AND (category LIKE '%tovar%' OR category LIKE '%товар%')")
        tovar = cursor.fetchone()[0] or 0.0
        
        # Ish haqi (total from work_records)
        cursor.execute("SELECT SUM(amount) FROM work_records")
        ish_haqi = cursor.fetchone()[0] or 0.0
        
        # Kommunal
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='chiqim' AND category LIKE '%kommunal%'")
        kommunal = cursor.fetchone()[0] or 0.0
        
        # Arenda
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='chiqim' AND category LIKE '%arenda%'")
        arenda = cursor.fetchone()[0] or 0.0
        
        # Marketing
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='chiqim' AND category LIKE '%marketing%'")
        marketing = cursor.fetchone()[0] or 0.0
        
        # Remont
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE type='chiqim' AND category LIKE '%remont%'")
        remont = cursor.fetchone()[0] or 0.0
        
        # Other chiqim categories if needed
        # Nalog = 2% of Savdo
        nalog = savdo * 0.02
        
        yalpi_foyda = savdo - tovar
        jami_xarajat = ish_haqi + kommunal + arenda + marketing + remont + nalog
        sof_foyda = yalpi_foyda - jami_xarajat
        
        conn.close()
        
        # Update UI labels
        self.bal_labels["Savdo"]["text"] = f"{savdo:,.0f} so'm"
        self.bal_labels["Sotib olingan tovar"]["text"] = f"{tovar:,.0f} so'm"
        self.bal_labels["Yalpi foyda (Savdo - Tovar)"]["text"] = f"{yalpi_foyda:,.0f} so'm"
        self.bal_labels["Ish haqi"]["text"] = f"{ish_haqi:,.0f} so'm"
        self.bal_labels["Kommunal"]["text"] = f"{kommunal:,.0f} so'm"
        self.bal_labels["Arenda"]["text"] = f"{arenda:,.0f} so'm"
        self.bal_labels["Marketing"]["text"] = f"{marketing:,.0f} so'm"
        self.bal_labels["Remont"]["text"] = f"{remont:,.0f} so'm"
        self.bal_labels["Nalog (Savdoning 2%)"]["text"] = f"{nalog:,.0f} so'm"
        self.bal_labels["Jami xarajat"]["text"] = f"{jami_xarajat:,.0f} so'm"
        self.bal_labels["Sof foyda"]["text"] = f"{sof_foyda:,.0f} so'm"

    def refresh_all(self):
        self.load_employees()
        self.load_salary()
        self.load_transactions()
        self.load_balance()
        messagebox.showinfo("Yangilandi", "Barcha ma'lumotlar yangilandi!")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            wb = openpyxl.Workbook()
            
            # Xodimlar
            ws_emp = wb.active
            ws_emp.title = "Xodimlar"
            ws_emp.append(["ID", "Ism familiya", "Lavozimi", "Telefon", "Tug'ilgan kun", "Status"])
            conn = self.get_db_connection()
            for row in conn.execute("SELECT * FROM employees"):
                ws_emp.append([row["id"], row["name"], row["position"], row["phone"], row["birth_date"], row["status"]])
                
            # Ish haqi
            ws_sal = wb.create_sheet(title="Ish haqi")
            ws_sal.append(["ID", "Xodim", "Sana", "Kirish", "Chiqish", "Kunbay summa", "Ishlagan vaqti"])
            for row in conn.execute("SELECT * FROM work_records"):
                ws_sal.append([row["id"], row["employee_name"], row["work_date"], row["check_in"], row["check_out"], row["amount"], row["worked_hours"]])
                
            # Tranzaksiyalar
            ws_tx = wb.create_sheet(title="Kirim-Chiqim")
            ws_tx.append(["ID", "Turi", "Modda", "Kontragent", "Summa", "Sana"])
            for row in conn.execute("SELECT * FROM transactions"):
                ws_tx.append([row["id"], row["type"], row["category"], row["kontragent"], row["amount"], row["tx_date"]])
                
            conn.close()
            wb.save(file_path)
            messagebox.SUCCESS = messagebox.showinfo("Muvaffaqiyat", f"Ma'lumotlar eksport qilindi:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Xatolik", f"Eksport qilishda xato: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RestaurantApp(root)
    root.mainloop()
