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
        self.root.geometry("800x750") 
        self.root.configure(bg=BG_COLOR)
        
        self.editing_index = None # Keeps track of which row we are editing

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

        self.title_label = tk.Label(self.header_frame, text="3D プリンタ 製作記録", 
                                   font=("MS Gothic", 18, "bold"), fg="white", bg=HEADER_COLOR)
        self.title_label.pack(expand=True)

        # --- Main Content Container ---
        self.main_container = tk.Frame(root, bg="white", highlightbackground=BORDER_COLOR, highlightthickness=1)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.setup_input_tab()
        self.setup_history_tab()

        self.notebook.add(self.input_frame, text=" 入力 ")
        self.notebook.add(self.history_frame, text=" 履歴 ")

        self.load_history()
    
    def setup_input_tab(self):
        self.input_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        tk.Frame(self.input_frame, height=20, bg=TAB_ACTIVE_BG).pack()

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

        tk.Frame(self.input_frame, height=30, bg=TAB_ACTIVE_BG).pack() 
        self.save_button = tk.Button(self.input_frame, text="データ保存", command=self.save_data, 
                                    bg="#5CB85C", fg="white", width=25, font=("Arial", 12, "bold"))
        self.save_button.pack()

    def create_entry_row(self, parent_frame, label_text, placeholder):
        row_frame = tk.Frame(parent_frame, bg=TAB_ACTIVE_BG, pady=8)
        row_frame.pack(fill=tk.X)
        tk.Label(row_frame, text=label_text, width=18, anchor='w', bg=TAB_ACTIVE_BG, font=("MS Gothic", 10)).pack(side=tk.LEFT)
        entry = tk.Entry(row_frame, width=30, font=("Arial", 10))
        entry.pack(side=tk.LEFT, padx=10)
        entry.insert(0, placeholder) 
        entry.config(fg='grey')
        entry.bind("<FocusIn>", lambda event: self.clear_placeholder(event, placeholder))
        entry.bind("<FocusOut>", lambda event: self.add_placeholder(event, placeholder))
        return entry

    def clear_placeholder(self, event, placeholder):
        if event.widget.get() == placeholder:
            event.widget.delete(0, tk.END)
            event.widget.config(fg='black')

    def add_placeholder(self, event, placeholder):
        if not event.widget.get():
            event.widget.insert(0, placeholder)
            event.widget.config(fg='grey')

    def setup_history_tab(self):
        self.history_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG, padx=10, pady=10)
        
        tools_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        tools_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(tools_frame, text="製作記録履歴", font=("MS Gothic", 12, "bold"), bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
        
        # New Edit Button
        tk.Button(tools_frame, text="編集", command=self.prepare_edit, bg="#F0AD4E", width=20).pack(side=tk.RIGHT)

        columns = ("date", "product", "time", "filament", "weight", "class", "maker")
        self.tree = ttk.Treeview(self.history_frame, columns=columns, show='headings')

        self.tree.heading("date", text="製作日")
        self.tree.heading("product", text="品名")
        self.tree.heading("time", text="製作時間")
        self.tree.heading("filament", text="使用フィラメント")
        self.tree.heading("weight", text="使用量")
        self.tree.heading("class", text="区分")
        self.tree.heading("maker", text="製作者")

        for col in columns:
            self.tree.column(col, width=100, anchor="center")

        scrollbar = ttk.Scrollbar(self.history_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def prepare_edit(self):
        """Takes selected item and puts it back into Input fields."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("編集エラー", "編集する行を選択してください。")
            return
        
        values = self.tree.item(selected[0])['values']
        
        # Fill Entry boxes with selected row data
        # Mapping tree values: 0:Date, 1:Product, 2:Time, 3:Filament, 4:Weight, 5:Class, 6:Maker
        self.entries["product"].delete(0, tk.END)
        self.entries["product"].insert(0, values[1])
        self.entries["product"].config(fg='black')

        self.entries["time"].delete(0, tk.END)
        self.entries["time"].insert(0, values[2])
        self.entries["time"].config(fg='black')

        self.entries["filament"].delete(0, tk.END)
        self.entries["filament"].insert(0, values[3])
        self.entries["filament"].config(fg='black')

        self.entries["weight"].delete(0, tk.END)
        self.entries["weight"].insert(0, str(values[4]).replace(" g", ""))
        self.entries["weight"].config(fg='black')

        self.entries["class"].delete(0, tk.END)
        self.entries["class"].insert(0, values[5])
        self.entries["class"].config(fg='black')

        self.entries["maker"].delete(0, tk.END)
        self.entries["maker"].insert(0, values[6])
        self.entries["maker"].config(fg='black')

        # Store the index of the row we are editing
        self.editing_index = self.tree.index(selected[0])
        
        # Switch to Input tab
        self.notebook.select(0)
        self.save_button.config(text="更新を保存 (Update Entry)", bg="#0275D8")

    def save_data(self):
        data = {key: self.entries[key].get().strip() for key in self.entries}
        date = datetime.now().strftime("%Y/%m/%d")

        # Validation
        for key, value in data.items():
            if value == self.fields[key][1] or not value:
                messagebox.showwarning("入力エラー", f"「{self.fields[key][0]}」を入力してください。")
                return

        file_path = "printer_history.csv"
        
        try:
            # If we are EDITING, we rewrite the whole file
            if self.editing_index is not None:
                rows = []
                with open(file_path, mode="r", encoding="utf-8-sig") as file:
                    reader = list(csv.reader(file))
                    header = reader[0]
                    rows = reader[1:]
                
                # Update the specific row (remember index is relative to data, not header)
                rows[self.editing_index] = [date, data["product"], data["time"], data["filament"], data["weight"]+" g", data["class"], data["maker"]]
                
                with open(file_path, mode="w", newline="", encoding="utf-8-sig") as file:
                    writer = csv.writer(file)
                    writer.writerow(header)
                    writer.writerows(rows)
                
                messagebox.showinfo("成功", "データが更新されました。")
                self.editing_index = None
                self.save_button.config(text="データ保存", bg="#5CB85C")

            else:
                # Regular new entry (APPEND mode)
                file_exists = os.path.isfile(file_path)
                with open(file_path, mode="a", newline="", encoding="utf-8-sig") as file:
                    writer = csv.writer(file)
                    if not file_exists:
                        writer.writerow(["製作日", "品名", "製作時間", "使用フィラメント", "使用量", "区分", "製作者"])
                    writer.writerow([date, data["product"], data["time"], data["filament"], data["weight"]+" g", data["class"], data["maker"]])
                messagebox.showinfo("成功", "データが保存されました。")

            # Reset fields
            for key, entry in self.entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, self.fields[key][1])
                entry.config(fg='grey')

            self.load_history()
            
        except PermissionError:
            messagebox.showerror("エラー", "Excelを閉じてから操作してください。")

    def load_history(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        file_path = "printer_history.csv"
        if os.path.exists(file_path):
            try:
                with open(file_path, mode="r", encoding="utf-8-sig") as file:
                    reader = csv.reader(file)
                    next(reader) 
                    for row in reader:
                        self.tree.insert("", tk.END, values=row)
            except Exception as e:
                print(f"Error loading CSV: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
