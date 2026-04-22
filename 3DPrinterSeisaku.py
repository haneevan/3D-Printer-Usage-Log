import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import os

# --- Visual Styling ---
BG_COLOR = "#E6E8EA"        
HEADER_COLOR = "#7F8387"    
TAB_ACTIVE_BG = "white"      

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3D プリンタ 製作記録")
        self.root.geometry("950x800")
        self.root.configure(bg=BG_COLOR)
        
        self.editing_row_data = None
        self.dept_db = self.load_dept_database()

        # UI Setup
        self.setup_styles()
        self.setup_main_ui()
        self.load_history()

    def load_dept_database(self):
        db = {}
        # Assuming departments.csv is in the same directory
        if os.path.exists("departments.csv"):
            with open("departments.csv", mode="r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    db[row['LookupCode']] = {
                        'dept': row['DeptName'],
                        'room': row['RoomName'],
                        'group': row['GroupName']
                    }
        return db

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=25, font=("Arial", 9))
        style.configure("Treeview.Heading", font=("MS Gothic", 10, "bold"))

    def setup_main_ui(self):
        header = tk.Frame(self.root, bg=HEADER_COLOR, height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="3D プリンタ 製作記録", font=("MS Gothic", 18, "bold"), fg="white", bg=HEADER_COLOR).pack(expand=True)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.input_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.history_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        
        self.notebook.add(self.input_frame, text=" 入力 ")
        self.notebook.add(self.history_frame, text=" 履歴 ")

        self.setup_input_tab()
        self.setup_history_tab()

    def setup_input_tab(self):
        # Allow the frame to center content
        self.input_frame.columnconfigure(0, weight=1)
        self.input_frame.rowconfigure(0, weight=1)

        # Central container for auto-scaling
        container = tk.Frame(self.input_frame, bg=TAB_ACTIVE_BG, padx=40, pady=20)
        container.grid(row=0, column=0, sticky="n") # Aligned to top-center
        
        container.columnconfigure(1, weight=1) # The entry column scales

        # 1. Date Input
        tk.Label(container, text="製作日:", anchor='w', bg=TAB_ACTIVE_BG, font=("MS Gothic", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.date_entry = tk.Entry(container, width=50, fg="grey")
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)")
        self.date_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=5)
        self.date_entry.bind("<FocusIn>", lambda e: self.on_date_focus_in())

        # 2. Main Input Fields (Standardized Column)
        self.fields = {
            "product": "品名:", "time": "製作時間 (h):", 
            "filament": "素材:", "weight": "使用量 (g):", 
            "class": "区分:", "maker": "製作者:"
        }
        self.entries = {}
        
        for i, (key, label) in enumerate(self.fields.items(), start=1):
            tk.Label(container, text=label, anchor='w', bg=TAB_ACTIVE_BG).grid(row=i, column=0, sticky="w", pady=5)
            ent = tk.Entry(container)
            ent.grid(row=i, column=1, sticky="ew", padx=10, pady=5)
            self.entries[key] = ent

        # 3. Department Code Input
        dept_row_idx = len(self.fields) + 1
        tk.Label(container, text="部署コード:", anchor='w', bg=TAB_ACTIVE_BG, font=("Arial", 10, "bold")).grid(row=dept_row_idx, column=0, sticky="w", pady=10)
        
        self.code_var = tk.StringVar()
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display) 
        self.code_entry = tk.Entry(container, textvariable=self.code_var, bg="#F0F8FF")
        self.code_entry.grid(row=dept_row_idx, column=1, sticky="ew", padx=10, pady=10)

        # 4. Result Display
        result_row_idx = dept_row_idx + 1
        self.result_display = tk.Frame(container, bg="#F9F9F9", pady=10, highlightbackground="#CCCCCC", highlightthickness=1)
        self.result_display.grid(row=result_row_idx, column=0, columnspan=2, sticky="ew", pady=15)
        
        self.result_label = tk.Label(self.result_display, text="結果 : --- ", font=("MS Gothic", 11), bg="#F9F9F9", fg="#555")
        self.result_label.pack(anchor='w', padx=20)

        # 5. Save Button
        save_row_idx = result_row_idx + 1
        self.save_button = tk.Button(container, text="データ保存", command=self.save_data, 
                                    bg="#5CB85C", fg="white", width=30, font=("Arial", 12, "bold"))
        self.save_button.grid(row=save_row_idx, column=0, columnspan=2, pady=20)

    def update_dept_display(self, *args):
        # Strip current hyphens to get clean digits
        raw_text = self.code_var.get().replace("-", "")[:12]
        
        formatted = ""
        if len(raw_text) <= 4:
            formatted = raw_text
        elif len(raw_text) <= 8:
            formatted = f"{raw_text[:4]}-{raw_text[4:]}"
        else:
            formatted = f"{raw_text[:4]}-{raw_text[4:8]}-{raw_text[8:]}"

        # Update text only if changed to avoid cursor issues
        self.code_var.trace_remove("write", self.trace_id)
        if self.code_var.get() != formatted:
            self.code_var.set(formatted)
            # Force cursor to the end so it doesn't jump left
            self.code_entry.after(1, lambda: self.code_entry.icursor(tk.END))
            
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display)

        # Results logic
        code = formatted.strip()
        if code in self.dept_db:
            info = self.dept_db[code]
            self.result_label.config(text=f"結果 : {info['dept']} || {info['room']} || {info['group']}", fg="blue")
        else:
            self.result_label.config(text="結果 : コードが見つかりません", fg="red")

    def on_date_focus_in(self):
        if "YYYY/MM/DD" in self.date_entry.get():
            self.date_entry.delete(0, tk.END)
            self.date_entry.config(fg="black")

    def setup_history_tab(self):
        cols = ("date", "prod", "time", "fil", "weight", "class", "maker", "dept", "room", "group")
        self.tree = ttk.Treeview(self.history_frame, columns=cols, show='headings')
        headers = ["製作日", "品名", "時間", "素材", "使用量", "区分", "製作者", "部署", "室", "グループ"]
        for col, h in zip(cols, headers):
            self.tree.heading(col, text=h)
            self.tree.column(col, width=90, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        btn_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(btn_frame, text="編集", command=self.prepare_edit, bg="#F0AD4E", width=15).pack(side=tk.RIGHT)

    def save_data(self):
        code = self.code_var.get().strip()
        if code not in self.dept_db:
            messagebox.showerror("エラー", "有効な部署コードを入力してください。")
            return

        input_date = self.date_entry.get().strip()
        date = input_date if input_date and "YYYY/MM/DD" not in input_date else datetime.now().strftime("%Y/%m/%d")

        dept_info = self.dept_db[code]
        row_to_save = [
            date, self.entries["product"].get(), self.entries["time"].get() + " h",
            self.entries["filament"].get(), self.entries["weight"].get() + " g",
            self.entries["class"].get(), self.entries["maker"].get(),
            dept_info['dept'], dept_info['room'], dept_info['group']
        ]

        file_path = "printer_history.csv"
        header = ["日付", "品名", "時間", "素材", "重量", "区分", "製作者", "部署", "室", "グループ"]

        try:
            if self.editing_row_data:
                all_rows = []
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    all_rows = list(csv.reader(f))
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    for r in all_rows[1:]:
                        if r == self.editing_row_data: writer.writerow(row_to_save)
                        else: writer.writerow(r)
                self.editing_row_data = None
                self.save_button.config(text="データ保存", bg="#5CB85C")
            else:
                exists = os.path.exists(file_path)
                with open(file_path, "a", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    if not exists: writer.writerow(header)
                    writer.writerow(row_to_save)

            messagebox.showinfo("成功", "データを保存しました。")
            self.load_history()
            self.clear_inputs()
        except PermissionError:
            messagebox.showerror("エラー", "CSVが開かれています。")

    def clear_inputs(self):
        for e in self.entries.values(): e.delete(0, tk.END)
        self.code_var.set("")
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)")
        self.date_entry.config(fg="grey")

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                data = list(csv.reader(f))[1:]
                for row in reversed(data): self.tree.insert("", tk.END, values=row)

    def prepare_edit(self):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])['values']
        self.editing_row_data = [str(v) for v in vals]
        
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, vals[0])
        self.date_entry.config(fg="black")

        self.entries["product"].insert(0, vals[1])
        self.entries["time"].insert(0, str(vals[2]).replace(" h", ""))
        self.entries["filament"].insert(0, vals[3])
        self.entries["weight"].insert(0, str(vals[4]).replace(" g", ""))
        self.entries["class"].insert(0, vals[5])
        self.entries["maker"].insert(0, vals[6])
        
        for code, info in self.dept_db.items():
            if info['dept'] == vals[7] and info['room'] == vals[8]:
                self.code_var.set(code)
                break
        
        self.notebook.select(0)
        self.save_button.config(text="更新を保存", bg="#0275D8")

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
