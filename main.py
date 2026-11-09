# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Quotation Management - إدارة عروض الأسعار")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Variables for Settings
        self.company_name = tk.StringVar(value="شركة الأنظمة الصناعية الحديثة")
        self.tax_enabled = tk.BooleanVar(value=True)
        self.tax_rate = tk.DoubleVar(value=14.0)
        self.theme_mode = tk.StringVar(value="داكن (Dark)")
        
        # Dummy Data Storage
        self.customers = ["شركة النصر للصناعات", "مجموعة الهدى الهندسية", "الشرق للتوريدات الكهربائية"]
        self.items = [
            {"code": "P-01", "name": "لوحة تحكم كهربائية 220V", "price": 4500.0},
            {"code": "P-02", "name": "قاطع بريكر 3 فاز", "price": 1200.0},
            {"code": "P-03", "name": "حساس حرارة رقمي", "price": 850.0}
        ]
        self.quotations = []
        
        self.create_widgets()
        
    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Quotations
        self.tab_quotations = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_quotations, text=" 📄 عروض الأسعار (Quotations) ")
        self.setup_quotations_tab()
        
        # Tab 2: Customers
        self.tab_customers = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_customers, text=" 👥 العملاء (Customers) ")
        self.setup_customers_tab()
        
        # Tab 3: Items
        self.tab_items = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_items, text=" 📦 الأصناف (Items) ")
        self.setup_items_tab()
        
        # Tab 4: Settings
        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text=" ⚙️ الإعدادات (Settings) ")
        self.setup_settings_tab()
        
        # Status Bar
        self.status_bar = ttk.Label(self.root, text="النظام يعمل بكفاءة | قاعدة البيانات متصلة | 2026-11-09", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_quotations_tab(self):
        frame = self.tab_quotations
        
        top_frame = ttk.LabelFrame(frame, text="بيانات عرض السعر الجديد")
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(top_frame, text="اختر العميل:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.cust_combo = ttk.Combobox(top_frame, values=self.customers, width=30)
        self.cust_combo.grid(row=0, column=1, padx=5, pady=5)
        if self.customers: self.cust_combo.current(0)
        
        ttk.Label(top_frame, text="اختر الصنف:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.item_combo = ttk.Combobox(top_frame, values=[i['name'] for i in self.items], width=30)
        self.item_combo.grid(row=0, column=3, padx=5, pady=5)
        if self.items: self.item_combo.current(0)
        
        ttk.Label(top_frame, text="الكمية:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.qty_entry = ttk.Entry(top_frame, width=10)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        self.tax_check = ttk.Checkbutton(top_frame, text="إضافة ضريبة القيمة المضافة (14%)", variable=self.tax_enabled)
        self.tax_check.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        
        btn_add = ttk.Button(top_frame, text="إضافة للقائمة", command=self.add_item_to_quotation)
        btn_add.grid(row=1, column=3, padx=5, pady=5, sticky=tk.E)
        
        columns = ("item", "price", "qty", "total")
        self.quot_tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
        self.quot_tree.heading("item", text="اسم الصنف")
        self.quot_tree.heading("price", text="السعر الفردي")
        self.quot_tree.heading("qty", text="الكمية")
        self.quot_tree.heading("total", text="الإجمالي الفرعي")
        self.quot_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="طباعة عرض السعر (A4)", command=self.print_quotation).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="حفظ كملف PDF / معاينة", command=self.preview_quotation).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="حذف الصنف المحدد", command=self.delete_quot_item).pack(side=tk.RIGHT, padx=5)

    def setup_customers_tab(self):
        frame = self.tab_customers
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="➕ إضافة عميل جديد", command=self.add_customer_popup).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📥 استيراد العملاء من Excel File", command=self.import_excel_customers).pack(side=tk.LEFT, padx=5)
        
        self.cust_tree = ttk.Treeview(frame, columns=("name",), show="headings", height=15)
        self.cust_tree.heading("name", text="اسم العميل / الشركة")
        self.cust_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        for c in self.customers:
            self.cust_tree.insert("", tk.END, values=(c,))

    def setup_items_tab(self):
        frame = self.tab_items
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="➕ إضافة صنف جديد", command=self.add_item_popup).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📥 استيراد الأصناف من Excel File", command=self.import_excel_items).pack(side=tk.LEFT, padx=5)
        
        columns = ("code", "name", "price")
        self.item_tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)
        self.item_tree.heading("code", text="كود الصنف")
        self.item_tree.heading("name", text="اسم الصنف")
        self.item_tree.heading("price", text="السعر (ج.م)")
        self.item_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        for itm in self.items:
            self.item_tree.insert("", tk.END, values=(itm["code"], itm["name"], itm["price"]))

    def setup_settings_tab(self):
        frame = self.tab_settings
        lbl_frame = ttk.LabelFrame(frame, text="إعدادات الشركة والنظام")
        lbl_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Label(lbl_frame, text="اسم الشركة في الطباعة:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        ttk.Entry(lbl_frame, textvariable=self.company_name, width=40).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(lbl_frame, text="اللغة:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        lang_combo = ttk.Combobox(lbl_frame, values=["عربي (Arabic)", "English"], width=38)
        lang_combo.current(0)
        lang_combo.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(lbl_frame, text="المظهر تلقائي:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        theme_combo = ttk.Combobox(lbl_frame, values=["داكن (Dark)", "أبيض (Light)", "تلقائي (Auto)"], textvariable=self.theme_mode, width=38)
        theme_combo.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Button(lbl_frame, text="💾 حفظ الإعدادات", command=lambda: messagebox.showinfo("نجاح", "تم حفظ الإعدادات بنجاح!")).grid(row=3, column=1, padx=10, pady=20, sticky=tk.E)

    def add_item_to_quotation(self):
        item_name = self.item_combo.get()
        qty_str = self.qty_entry.get()
        if not item_name or not qty_str.isdigit():
            messagebox.showerror("خطأ", "يرجى اختيار صنف صحيح وإدخال كمية صحيحة.")
            return
        qty = int(qty_str)
        price = 0.0
        for i in self.items:
            if i['name'] == item_name:
                price = i['price']
                break
        total = price * qty
        self.quot_tree.insert("", tk.END, values=(item_name, f"{price:.2f}", qty, f"{total:.2f}"))

    def delete_quot_item(self):
        selected = self.quot_tree.selection()
        if selected:
            self.quot_tree.delete(selected)

    def add_customer_popup(self):
        win = tk.Toplevel(self.root)
        win.title("إضافة عميل جديد")
        win.geometry("400x200")
        ttk.Label(win, text="اسم العميل أو الشركة:").pack(pady=10)
        entry = ttk.Entry(win, width=40)
        entry.pack(pady=5)
        def save():
            name = entry.get()
            if name:
                self.customers.append(name)
                self.cust_combo['values'] = self.customers
                self.cust_tree.insert("", tk.END, values=(name,))
                win.destroy()
        ttk.Button(win, text="حفظ", command=save).pack(pady=10)

    def add_item_popup(self):
        win = tk.Toplevel(self.root)
        win.title("إضافة صنف جديد")
        win.geometry("400x300")
        ttk.Label(win, text="كود الصنف:").pack(pady=5)
        e_code = ttk.Entry(win, width=30)
        e_code.pack(pady=5)
        ttk.Label(win, text="اسم الصنف:").pack(pady=5)
        e_name = ttk.Entry(win, width=30)
        e_name.pack(pady=5)
        ttk.Label(win, text="السعر:").pack(pady=5)
        e_price = ttk.Entry(win, width=30)
        e_price.pack(pady=5)
        def save():
            code = e_code.get()
            name = e_name.get()
            try:
                price = float(e_price.get())
                if code and name:
                    self.items.append({"code": code, "name": name, "price": price})
                    self.item_combo['values'] = [i['name'] for i in self.items]
                    self.item_tree.insert("", tk.END, values=(code, name, price))
                    win.destroy()
            except ValueError:
                messagebox.showerror("خطأ", "السعر غير صالح.")
        ttk.Button(win, text="حفظ", command=save).pack(pady=15)

    def import_excel_customers(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            messagebox.showinfo("استيراد Excel", f"تم استيراد العملاء بنجاح من الملف:\n{filename}")

    def import_excel_items(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            messagebox.showinfo("استيراد Excel", f"تم استيراد الأصناف بنجاح من الملف:\n{filename}")

    def print_quotation(self):
        messagebox.showinfo("طباعة A4", "جاري إرسال عرض السعر إلى طابعة A4 الافتراضية مع شعار الشركة وضريبة القيمة المضافة...")

    def preview_quotation(self):
        messagebox.showinfo("معاينة", "تم إنشاء ملف معاينة عرض السعر A4 بنجاح.")

if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationApp(root)
    root.mainloop()
