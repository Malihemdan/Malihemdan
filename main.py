import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from openpyxl import Workbook

# ================= قاعدة البيانات =================
conn = sqlite3.connect("purchases.db")
cursor = conn.cursor()

# حذف الجدول القديم وإنشاؤه من جديد
# إنشاء الجدول لو مش موجود بالفعل
cursor.execute("""
CREATE TABLE IF NOT EXISTS purchases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    purchase_order_no TEXT,
    purchase_order_date TEXT,
    purchase_order_amount REAL,
    company TEXT,
    supply_order_no TEXT,
    supply_order_amount REAL,
    supply_order_date TEXT,
    saving REAL,
    duration INTEGER
)
""")
conn.commit()


# ================= الدوال =================
def add_record():
    try:
        po_no = entry_po_no.get()
        po_date = entry_po_date.get()
        po_amount = float(entry_po_amount.get())
        company = entry_company.get()
        so_no = entry_so_no.get()
        so_amount = float(entry_so_amount.get())
        so_date = entry_so_date.get()

        # حساب التوفير
        saving = po_amount - so_amount

        # حساب المدة
        duration = (
            datetime.strptime(so_date, "%Y-%m-%d") - datetime.strptime(po_date, "%Y-%m-%d")
        ).days

        cursor.execute("""
        INSERT INTO purchases (purchase_order_no, purchase_order_date, purchase_order_amount,
        company, supply_order_no, supply_order_amount, supply_order_date, saving, duration)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (po_no, po_date, po_amount, company, so_no, so_amount, so_date, saving, duration))
        conn.commit()

        messagebox.showinfo("تم", "تم حفظ العملية بنجاح")
        show_records()

    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ: {e}")

def show_records():
    for row in tree.get_children():
        tree.delete(row)
    cursor.execute("SELECT * FROM purchases")
    for record in cursor.fetchall():
        tree.insert("", tk.END, values=record)

def export_excel():
    cursor.execute("SELECT * FROM purchases")
    records = cursor.fetchall()

    if not records:
        messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "التقرير"

    # عناوين الأعمدة
    headers = ["ID", "رقم طلب الشراء", "تاريخ طلب الشراء", "مبلغ طلب الشراء",
               "الشركة", "رقم امر التوريد", "مبلغ امر التوريد", "تاريخ امر التوريد",
               "مبلغ التوفير", "مدة العملية"]
    ws.append(headers)

    total_po = 0
    total_so = 0
    total_duration = 0

    for row in records:
        ws.append(row)
        total_po += row[3]
        total_so += row[6]
        total_duration += row[9]

    # تقرير إجمالي
    ws.append([])
    ws.append(["", "إجمالي طلبات الشراء", total_po])
    ws.append(["", "إجمالي أوامر التوريد", total_so])
    ws.append(["", "متوسط مدة العملية (يوم)", total_duration / len(records)])

    wb.save("purchase_report.xlsx")
    messagebox.showinfo("تم", "تم تصدير التقرير إلى Excel بنجاح")

# ================= واجهة المستخدم =================
root = tk.Tk()
root.title("✍برنامج إدارة المشتريات")
root.geometry("1000x600")
# Notebook (التبويبات)
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# ================= التبويب الأول (إدارة المشتريات) =================
frame_main = tk.Frame(notebook)
notebook.add(frame_main, text="إدارة المشتريات")


# الإدخالات
tk.Label(frame_main, text="رقم طلب الشراء").grid(row=0, column=0)
entry_po_no = tk.Entry(frame_main)
entry_po_no.grid(row=0, column=1)

tk.Label(frame_main, text="تاريخ طلب الشراء (YYYY-MM-DD)").grid(row=0, column=2)
entry_po_date = tk.Entry(frame_main)
entry_po_date.grid(row=0, column=3)

tk.Label(frame_main, text="مبلغ طلب الشراء").grid(row=1, column=0)
entry_po_amount = tk.Entry(frame_main)
entry_po_amount.grid(row=1, column=1)

tk.Label(frame_main, text="الشركة").grid(row=1, column=2)
entry_company = tk.Entry(frame_main)
entry_company.grid(row=1, column=3)

tk.Label(frame_main, text="رقم أمر التوريد").grid(row=2, column=0)
entry_so_no = tk.Entry(frame_main)
entry_so_no.grid(row=2, column=1)

tk.Label(frame_main, text="مبلغ أمر التوريد").grid(row=2, column=2)
entry_so_amount = tk.Entry(frame_main)
entry_so_amount.grid(row=2, column=3)

tk.Label(frame_main, text="تاريخ أمر التوريد (YYYY-MM-DD)").grid(row=3, column=0)
entry_so_date = tk.Entry(frame_main)
entry_so_date.grid(row=3, column=1)

# الأزرار
tk.Button(frame_main, text="إضافة عملية", command=add_record).grid(row=4, column=0, pady=10)
tk.Button(frame_main, text="عرض العمليات", command=show_records).grid(row=4, column=1, pady=10)
tk.Button(frame_main, text="تصدير Excel", command=export_excel).grid(row=4, column=2, pady=10)

# جدول عرض البيانات
columns = ("ID", "رقم طلب الشراء", "تاريخ طلب الشراء", "مبلغ طلب الشراء",
           "الشركة", "رقم أمر التوريد", "مبلغ أمر التوريد", "تاريخ أمر التوريد",
           "مبلغ التوفير", "مدة العملية")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)
tree.pack(fill=tk.BOTH, expand=True)

# ================= التبويب الثاني (حقوق الملكية) =================
frame_copy = tk.Frame(notebook)
notebook.add(frame_copy, text="© حقوق الملكية")

label_copy = tk.Label(
    frame_copy,
    text="© 2025 جميع الحقوق محفوظة\nتم تطوير هذا البرنامج بواسطة [محمود على \n mahmoud.ali20@hotmail.com\n +201069225114]",
    font=("Arial", 14),
    pady=50
)
label_copy.pack(expand=True)

# تشغيل البرنامج
root.mainloop()


