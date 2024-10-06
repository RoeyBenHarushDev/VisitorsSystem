import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import ttk as tk_ttk
import pandas as pd
from datetime import datetime
from ttkbootstrap.scrolled import ScrolledFrame
from PIL import Image, ImageTk
import os

class VisitorManagementSystem:
    def __init__(self, master):
        self.master = master
        self.master.title("מערכת ניהול מבקרים")
        self.master.geometry("2700x1700")
        self.style = ttk.Style("flatly")
        self.style.configure('.', font=('Segoe UI', 12))  # השתמש בגופן סגנון מינימליסטי
        self.visitors = []
        self.current_visitors = []

        # הגדרת צבעים בעיצוב Apple
        self.style.configure('TFrame', background='#F5F5F5')  # רקע לבן-אפור בהיר
        self.style.configure('TLabel', background='#F5F5F5', foreground='#333333', font=('Segoe UI', 14))
        self.style.configure('TButton', font=('Segoe UI', 12), padding=(10, 5), relief="flat", background='#007AFF')
        self.style.configure('Treeview', rowheight=40, font=('Segoe UI', 12), background='#ffffff')
        self.style.configure('Treeview.Heading', font=('Segoe UI', 12, 'bold'))

        self.create_widgets()

    def create_widgets(self):
        main_frame = ScrolledFrame(self.master)
        main_frame.pack(fill=BOTH, expand=YES)

        content_frame = ttk.Frame(main_frame, padding="30 20 30 20")
        content_frame.pack(fill=BOTH, expand=YES)

        # הוספת לוגו Apple-style
        logo_path = "audiocodes_log.png"
        if os.path.exists(logo_path):
            logo_image = Image.open(logo_path)
            logo_image = logo_image.resize((650, 150), Image.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(content_frame, image=logo_photo, background='#F5F5F5')
            logo_label.image = logo_photo
            logo_label.pack(pady=(0, 20))

        # כותרת עיצוב מינימליסטי
        ttk.Label(content_frame, text="מערכת ניהול מבקרים", font=("Segoe UI", 24, "bold"), foreground='#007AFF').pack(pady=20)

        # מסגרת לטופס
        form_frame = ttk.Frame(content_frame, style='TFrame')
        form_frame.pack(fill=X, pady=20)

        # שדות הקלט בעיצוב אפל
        fields = [("תאריך", "date"), ("שם מלא", "name"), ("מספר טלפון", "phone"), ("למי הגיע בחברה", "host")]
        self.entries = {}

        for i, (label, field) in enumerate(fields):
            frame = ttk.Frame(form_frame, style='TFrame')
            frame.pack(fill=X, pady=10)
            ttk.Label(frame, text=f"{label}:", width=15, anchor="e", style='TLabel').pack(side=RIGHT, padx=(0, 10))
            entry = ttk.Entry(frame, width=40, font=('Segoe UI', 12), bootstyle="default")
            entry.pack(side=RIGHT, expand=YES, fill=X)
            self.entries[field] = entry

        # הגדרת תאריך ברירת מחדל
        self.entries['date'].insert(0, datetime.now().strftime("%Y-%m-%d"))

        # כפתורים מעוצבים בסגנון אפל
        button_frame = ttk.Frame(content_frame, style='TFrame')
        button_frame.pack(anchor='center')

        ttk.Button(button_frame, text="הוסף מבקר", command=self.add_visitor, 
                   bootstyle="success-outline", width=15).pack(side=LEFT, padx=15, pady=10)

        ttk.Button(button_frame, text="Excel - ייצא ל", command=self.export_to_excel, 
                   bootstyle="info-outline", width=15).pack(side=LEFT, padx=15, pady=10)

        # מסגרת למבקרים נוכחיים
        ttk.Label(content_frame, text="מבקרים נוכחיים", font=("Segoe UI", 18, "bold"), foreground='#007AFF').pack(pady=(20, 10))
        self.current_visitors_frame = ScrolledFrame(content_frame, autohide=True, height=300)
        self.current_visitors_frame.pack(fill=X, expand=YES, pady=10)

        # טבלת כל המבקרים
        ttk.Label(content_frame, text="כל המבקרים", font=("Segoe UI", 18, "bold"), foreground='#007AFF').pack(pady=(20, 10))
        self.all_visitors_frame = ttk.Frame(content_frame, style='TFrame')
        self.all_visitors_frame.pack(fill=BOTH, expand=YES, pady=10)

        self.update_tables()

    def add_visitor(self):
        try:
            visitor = {
                "תאריך": self.entries['date'].get() or datetime.now().strftime("%Y-%m-%d"),
                "שם מלא": self.entries['name'].get(),
                "מספר טלפון": self.entries['phone'].get(),
                "למי הגיע": self.entries['host'].get(),
                "שעת כניסה": datetime.now().strftime("%H:%M:%S"),
                "שעת יציאה": ""
            }

            empty_fields = [field for field, value in visitor.items() if not value and field != "שעת יציאה"]
            
            if empty_fields:
                error_message = f"אנא מלא את השדות הבאים: {', '.join(empty_fields)}"
                Messagebox.show_error("שגיאה", error_message, parent=self.master)
                return

            self.visitors.append(visitor)
            self.current_visitors.append(visitor)
            self.clear_entries()
            self.update_tables()

        except Exception as e:
            error_message = f"אירעה שגיאה בעת הוספת המבקר: {str(e)}"
            Messagebox.show_error("שגיאה", error_message, parent=self.master)
            print(f"Error in add_visitor: {e}")

    def clear_entries(self):
        for entry in self.entries.values():
            entry.delete(0, ttk.END)
        self.entries['date'].insert(0, datetime.now().strftime("%Y-%m-%d"))

    def update_tables(self):
        self.update_current_visitors_table()
        self.update_all_visitors_table()

    def update_current_visitors_table(self):
        for widget in self.current_visitors_frame.winfo_children():
            widget.destroy()

        if self.current_visitors:
            grid_frame = ttk.Frame(self.current_visitors_frame, style='TFrame')
            grid_frame.pack(fill=BOTH, expand=YES)
            for i, visitor in enumerate(self.current_visitors):
                self.create_visitor_card(grid_frame, visitor, i)
        else:
            ttk.Label(self.current_visitors_frame, text="אין מבקרים נוכחיים", font=("Segoe UI", 14), style='TLabel').pack(pady=20)

    def create_visitor_card(self, parent, visitor, index):
        card = ttk.Frame(parent, style='TFrame', borderwidth=1, relief="solid", padding=10)
        card.grid(row=0, column=index, padx=10, pady=10, sticky="nsew")
        card.configure(width=200, height=220)

        # עיצוב כרטיסי מבקרים
        style = ttk.Style()
        style.configure(f"Card{index}.TFrame", background="#FFFFFF")
        card.configure(style=f"Card{index}.TFrame")

        ttk.Label(card, text=f"{visitor['שם מלא']}", font=('Segoe UI', 14, 'bold'), foreground='#007AFF', background='#ffffff').pack(pady=(10, 5))
        ttk.Label(card, text=f"למי הגיע: {visitor['למי הגיע']}", font=('Segoe UI', 12), background='#ffffff').pack(anchor="w", padx=10, pady=2)
        ttk.Label(card, text=f"שעת כניסה: {visitor['שעת כניסה']}", font=('Segoe UI', 12), background='#ffffff').pack(anchor="w", padx=10, pady=2)

        ttk.Button(card, text="יציאה", 
                   command=lambda v=visitor['שם מלא']: self.mark_visitor_exit(v),
                   bootstyle="danger-outline", width=10).pack(pady=(15, 10))

    def update_all_visitors_table(self):
        for widget in self.all_visitors_frame.winfo_children():
            widget.destroy()

        if self.visitors:
            columns = ("תאריך", "שם מלא", "מספר טלפון", "למי הגיע", "שעת כניסה", "שעת יציאה")
            tree = tk_ttk.Treeview(self.all_visitors_frame, columns=columns, show='headings')

            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor='center')

            for visitor in self.visitors:
                tree.insert('', 'end', values=(visitor["תאריך"], visitor["שם מלא"], visitor["מספר טלפון"],
                                               visitor["למי הגיע"], visitor["שעת כניסה"], visitor["שעת יציאה"]))

            tree.pack(fill=BOTH, expand=YES)
        else:
            ttk.Label(self.all_visitors_frame, text="אין מבקרים להצגה", font=("Segoe UI", 14), style='TLabel').pack(pady=20)

    def mark_visitor_exit(self, visitor_name):
        for visitor in self.current_visitors:
            if visitor["שם מלא"] == visitor_name:
                visitor["שעת יציאה"] = datetime.now().strftime("%H:%M:%S")
                self.current_visitors.remove(visitor)
                break
        for visitor in self.visitors:
            if visitor["שם מלא"] == visitor_name and visitor["שעת יציאה"] == "":
                visitor["שעת יציאה"] = datetime.now().strftime("%H:%M:%S")
                break
        self.update_tables()

    def export_to_excel(self):
        if self.visitors:
            df = pd.DataFrame(self.visitors)
            filename = f"visitors-{datetime.now().strftime('%Y%m%d')}.xlsx"
            df.to_excel(filename, index=False)
            Messagebox.show_info("ייצוא הצליח", f"הקובץ נשמר בשם: {filename}", parent=self.master)
        else:
            Messagebox.show_warning("אין נתונים", "אין מבקרים לייצא", parent=self.master)

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    app = VisitorManagementSystem(root)
    root.mainloop()
