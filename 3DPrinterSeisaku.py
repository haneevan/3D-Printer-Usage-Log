import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import os

# --- Visual Styling Constants ---
BG_COLOR = "#E6E8EA"        
HEADER_COLOR = "#7F8387"    
BORDER_COLOR = "#192E3B"    
TAB_ACTIVE_BG = "white"      

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3D プリンタ 製作記録")
        self.root.geometry("900x750") # Slightly wider for new columns
        self.root.configure(bg=BG_COLOR)
        
        self.editing_row_data = None
        self.dept_data = self.load_dept_database()

        style = ttk.Style()
        style.theme_use('clam') 
        style.configure('TNotebook', background=BG_COLOR, borderwidth=0)
        style.configure('TNotebook.Tab', padding=[5, 5], background="#CED4D9") 
        style.map('TNotebook.Tab', background=[('selected', TAB_ACTIVE_BG)])     
        
        style.configure("Treeview", rowheight=25, font=("Arial", 9))
        style.configure("Treeview.Heading", font=("MS Gothic", 10, "bold"))

        # --- Top Header Bar ---
        self.header_frame = tk.Frame(root, bg=HEADER_COLOR, height=60)
        self.header_frame.pack(fill=tk.X, side=tk.TOP, pady=0)
        self.header_frame.pack_propagate(False) 
        tk.Label(self.header_frame, text="3D プリンタ 製作記録", font=("MS Gothic", 18, "bold"), fg="white", bg=HEADER_COLOR).pack(expand=True)

        self.main_container = tk.Frame(root, bg="white", highlightbackground=BORDER_COLOR, highlightthickness=1)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.setup_input_tab()
        self.setup_history_tab()

        self.notebook.add(self.input_frame, text=" 入力 ")
        self.notebook.add(self.history_frame, text=" 履歴 ")

        self.load_history()

    def load_dept_database(self):
        """Loads department details from CSV into a dictionary."""
        db = {}
        if os.path.exists("departments.csv"):
            with open("departments.csv", mode="r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    db[row['Department']] = {
                        'dept_code': row['DeptCode'],
                        'room': row['RoomCode'],
                        'group': row['GroupCode']
                    }
        return db

    def setup_input_tab(self):
        self.input_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        tk.Frame(self.input_frame, height=10, bg=TAB_ACTIVE_BG).pack()
        form_frame = tk.Frame(self.input_frame, bg=TAB_ACTIVE_BG, padx=40)
        form_frame.pack(fill=tk.X)

        self.fields = {
            "product": ("品名:", ". . . . ."),
            "time": ("製作時間 (h):", "e.g., 2.5"),
            "filament": ("使用フィラメント:", "PLA, PETG..."),
            "weight": ("使用量 (g):", "e.g., 50"),
            "class": ("区分:", "改善, プレス, その他..."),
            "maker": ("製作者:", ". . . .")
        }

        self.entries = {}
        for key, (label_text, placeholder) in self.fields.items():
            self.entries[key] = self.create_entry_row(form_frame, label_text, placeholder)

        # Department Dropdown
        row_frame = tk.Frame(form_frame, bg=TAB_ACTIVE_BG, pady=8)
        row_frame.pack(fill=tk.X)
        tk.Label(row_frame, text="部署 (Department):", width=18, anchor='w', bg=TAB_ACTIVE_BG, font=("MS Gothic", 10)).pack(side=tk.LEFT)
        self.dept_var = tk.StringVar()
        self.dept_combo = ttk.Combobox(row_frame, textvariable=self.dept_var, width=28, state="readonly")
        self.dept_combo['values'] = list(self.dept_data.keys())
        self.dept_combo.pack(side=tk.LEFT, padx=10)

        tk.Frame(self.input_frame, height=20, bg=TAB_ACTIVE_BG).pack() 
        self.save_button = tk.Button(self.input_frame, text="データ保存", command=self.save_data, bg="#5CB85C", fg="white", width=25, font=("Arial", 12, "bold"))
        self.save_button.pack()

    def create_entry_row(self, parent_frame, label_text, placeholder):
        row_frame = tk.Frame(parent_frame, bg=TAB_ACTIVE_BG, pady=8)
        row_frame.pack(fill=tk.X)
        tk.Label(row_frame, text=label_text, width=18, anchor='w', bg=TAB_ACTIVE_BG, font=("MS Gothic", 10)).pack(side=tk.LEFT)
        entry = tk.Entry(row_frame, width=30, font=("Arial", 10))
        entry.pack(side=tk.LEFT, padx=10)
        entry.insert(0, placeholder); entry.config(fg='grey')
        entry.bind("<FocusIn>", lambda e: self.clear_placeholder(e, placeholder))
        entry.bind("<FocusOut>", lambda e: self.add_placeholder(e, placeholder))
        return entry

    def clear_placeholder(self, event, placeholder):
        if event.widget.get() == placeholder:
            event.widget.delete(0, tk.END); event.widget.config(fg='black')

    def add_placeholder(self, event, placeholder):
        if not event.widget.get():
            event.widget.insert(0, placeholder); event.widget.config(fg='grey')

    def setup_history_tab(self):
        self.history_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG, padx=10, pady=10)
        tools_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        tools_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(tools_frame, text="製作記録履歴 (最新順)", font=("MS Gothic", 12, "bold"), bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
        tk.Button(tools_frame, text="編集", command=self.prepare_edit, bg="#F0AD4E", width=15).pack(side=tk.RIGHT)

        # Columns include Dept info now
        columns = ("date", "product", "time", "filament", "weight", "class", "maker", "dept", "dcode", "rcode", "gcode")
        self.tree = ttk.Treeview(self.history_frame, columns=columns, show='headings')
        
        headers = ["製作日", "品名", "時間", "素材", "使用量", "区分", "製作者", "部署", "部ｺｰﾄﾞ", "室ｺｰﾄﾞ", "班ｺｰﾄﾞ"]
        for col, header in zip(columns, headers):
            self.tree.heading(col, text=header)
            self.tree.column(col, width=80, anchor="center")

        scrollbar = ttk.Scrollbar(self.history_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def save_data(self):
        data = {key: self.entries[key].get().strip() for key in self.entries}
        selected_dept = self.dept_var.get()
        date = datetime.now().strftime("%Y/%m/%d")

        if not selected_dept:
            messagebox.showwarning("入力エラー", "部署を選択してください。")
            return
        
        for key, val in data.items():
            if val == self.fields[key][1] or not val:
                messagebox.showwarning("入力エラー", f"「{self.fields[key][0]}」を入力してください。")
                return

        dept_info = self.dept_data[selected_dept]
        new_row = [date, data["product"], data["time"]+" h", data["filament"], data["weight"]+" g", 
                   data["class"], data["maker"], selected_dept, dept_info['dept_code'], dept_info['room'], dept_info['group']]

        file_path = "printer_history.csv"
        header = ["製作日", "品名", "製作時間", "使用フィラメント", "使用量", "区分", "製作者", "部署", "部コード", "室コード", "班コード"]

        try:
            if self.editing_row_data:
                rows = []
                with open(file_path, mode="r", encoding="utf-8-sig") as f:
                    content = list(csv.reader(f))
                    rows = content[1:]
                
                # Update logic
                for i, row in enumerate(rows):
                    if row == self.editing_row_data:
                        rows[i] = new_row
                        break
                
                with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header); writer.writerows(rows)
                self.editing_row_data = None
                self.save_button.config(text="データ保存", bg="#5CB85C")
            else:
                exists = os.path.isfile(file_path)
                with open(file_path, mode="a", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    if not exists: writer.writerow(header)
                    writer.writerow(new_row)

            messagebox.showinfo("成功", "保存されました。")
            self.load_history()
            # Clear fields
            for key, ent in self.entries.items():
                ent.delete(0, tk.END); ent.insert(0, self.fields[key][1]); ent.config(fg='grey')
            self.dept_var.set("")
            
        except PermissionError:
            messagebox.showerror("エラー", "Excelを閉じてから操作してください。")

    def load_history(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", mode="r", encoding="utf-8-sig") as f:
                reader = list(csv.reader(f))
                if len(reader) > 1:
                    for row in reversed(reader[1:]): self.tree.insert("", tk.END, values=row)

    def prepare_edit(self):
        sel = self.tree.selection()
        if not sel: return
        values = [str(v) for v in self.tree.item(sel[0])['values']]
        self.editing_row_data = values
        
        # Fill entries
        mapping = {"product": 1, "time": 2, "filament": 3, "weight": 4, "class": 5, "maker": 6}
        for key, idx in mapping.items():
            self.entries[key].delete(0, tk.END)
            clean_val = values[idx].replace(" h", "").replace(" g", "")
            self.entries[key].insert(0, clean_val); self.entries[key].config(fg='black')
        
        self.dept_var.set(values[7])
        self.notebook.select(0)
        self.save_button.config(text="更新を保存", bg="#0275D8")

if __name__ == "__main__":
    root = tk.Tk(); app = MyApp(root); root.mainloop()
