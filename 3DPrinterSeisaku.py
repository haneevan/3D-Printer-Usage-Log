import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
try:
    import requests
except ImportError:
    requests = None
import json
from tkcalendar import DateEntry

# --- Font Fallback Configurations ---
import matplotlib
matplotlib.rcParams['font.family'] = ['MS Gothic', 'IPAexGothic', 'Noto Sans CJK JP', 'sans-serif']

default_linux_font = "Noto Sans CJK JP"

TK_FONT_BOLD = ("MS Gothic" if os.name == "nt" else default_linux_font, 11, "bold")
TK_FONT_REGULAR = ("MS Gothic" if os.name == "nt" else default_linux_font, 11)
TK_FONT_HEADER = ("MS Gothic" if os.name == "nt" else default_linux_font, 18, "bold")
TK_FONT_TREEVIEW = ("MS Gothic" if os.name == "nt" else default_linux_font, 10, "bold")

BG_COLOR = "#E6E8EA"         
HEADER_COLOR = "#7F8387"    
TAB_ACTIVE_BG = "white"      

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3D プリンタ 製作記録")
        self.root.geometry("1150x850") 
        self.root.configure(bg=BG_COLOR)
        
        self.editing_row_data = None 
        self.dept_db = self.load_dept_database()

        # --- DYNAMICALLY LOAD PRICING ARRAYS FROM FILE ---
        pricing_data = self.load_pricing_config()
        
        # Pull out electricity cost separately
        self.elec_cost_per_hour = pricing_data.pop("electricity_per_hour", 15)
        # The remaining dictionary keys are your material per-kg prices
        self.material_prices = pricing_data

        self.setup_styles()
        self.setup_main_ui()
        self.load_history()
        self.setup_edit_tab()


    def load_pricing_config(self):
        """Loads material prices per kg and electricity cost from disk. Uses defaults if missing."""
        import os
        filename = "pricing_config.txt"
        defaults = {
            "PLA": "2800",
            "ABS": "3400",
            "PC": "7280",
            "PET-CF": "18360",
            "electricity_per_hour": "15"
        }
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8") as f:
                for k, v in defaults.items():
                    f.write(f"{k}={v}\n")
            # Convert values to integers for application memory
            return {k: int(v) for k, v in defaults.items()}
            
        pricing = {}
        with open(filename, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    try:
                        pricing[k] = int(v)
                    except ValueError:
                        pricing[k] = 0
        
        # Ensure fallback defaults exist if keys were missing or broken in file
        for k, v in defaults.items():
            if k not in pricing:
                pricing[k] = int(v)
        return pricing

    def save_pricing_config(self, pricing_dict):
        """Saves current memory pricing variables back into the text file."""
        with open("pricing_config.txt", "w", encoding="utf-8") as f:
            for k, v in pricing_dict.items():
                f.write(f"{k}={v}\n")

    def load_dept_database(self):
        db = {}
        if os.path.exists("departments.csv"):
            with open("departments.csv", mode="r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    db[row['LookupCode']] = {
                        'dept': row['DeptName'], 'room': row['RoomName'], 'group': row['GroupName']
                    }
        return db
    
    def load_classifications(self):
        """Loads classifications from the text file. Defaults if file missing."""
        import os
        filename = "classification_list.txt"
        default_list = ["内製化", "試作", "検討品", "品質", "改善"]
        
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8") as f:
                f.write("\n".join(default_list))
            return default_list
        
        with open(filename, "r", encoding="utf-8") as f:
            # Read lines and strip out invisible spaces/newlines
            choices = [line.strip() for line in f if line.strip()]
        return choices if choices else default_list

    def save_classifications(self, choices):
        """Saves the current classifications back to the text file."""
        with open("classification_list.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(choices))

    def load_materials(self):
        """Loads materials from the text file. Defaults if file missing."""
        import os
        filename = "material_list.txt"
        default_materials = ["PLA", "ABS", "PC", "PET-CF"]
        
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8") as f:
                f.write("\n".join(default_materials))
            return default_materials
        
        with open(filename, "r", encoding="utf-8") as f:
            choices = [line.strip() for line in f if line.strip()]
        return choices if choices else default_materials
    
    def save_materials(self, materials):
        """Saves the current materials list back to the text file."""
        with open("material_list.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(materials))

    def load_colors(self):
        """Loads colors from the text file. Defaults if file missing."""
        import os
        filename = "color_list.txt"
        default_colors = ["白", "黒", "赤", "青", "グレー"]
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8") as f:
                f.write("\n".join(default_colors))
            return default_colors
        with open(filename, "r", encoding="utf-8") as f:
            choices = [line.strip() for line in f if line.strip()]
        return choices if choices else default_colors

    def save_colors(self, colors):
        """Saves the current colors list back to the text file."""
        with open("color_list.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(colors))

    def load_producers(self):
        """Loads producers from the text file. Defaults if file missing."""
        import os
        filename = "producer_list.txt"
        default_producers = ["担当A", "担当B", "担当C"]
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8") as f:
                f.write("\n".join(default_producers))
            return default_producers
        with open(filename, "r", encoding="utf-8") as f:
            choices = [line.strip() for line in f if line.strip()]
        return choices if choices else default_producers

    def save_producers(self, producers):
        """Saves the current producers list back to the text file."""
        with open("producer_list.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(producers))


    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=35, font=TK_FONT_REGULAR)
        style.configure("Treeview.Heading", font=(TK_FONT_REGULAR[0], 10, "bold"))

    def setup_main_ui(self):
        header = tk.Frame(self.root, bg=HEADER_COLOR, height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="3D プリンタ 製作記録", font=TK_FONT_HEADER, fg="white", bg=HEADER_COLOR).pack(expand=True)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.input_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.history_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.stats_class_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.stats_dept_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        
        self.class_month_var = tk.StringVar(value="全期間")
        self.dept_month_var = tk.StringVar(value="全期間")
        
        self.notebook.add(self.input_frame, text=" 入力 ")
        self.notebook.add(self.history_frame, text=" 履歴 ")
        self.notebook.add(self.stats_class_frame, text=" 製作内容集計 ")
        self.notebook.add(self.stats_dept_frame, text=" 部署別利用率 ")

        self.setup_input_tab()
        self.setup_history_tab()
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def setup_input_tab(self):
        UI_CONFIG = {
            "entry_width": 45,
            "row_padding": 10,
            "label_font": TK_FONT_BOLD,
            "entry_font": TK_FONT_REGULAR
        }
        
        from tkcalendar import DateEntry
        from datetime import datetime

        # Main split container for input form (left) and keyboard (right)
        split_container = tk.Frame(self.input_frame, bg=TAB_ACTIVE_BG)
        split_container.pack(fill=tk.BOTH, expand=True)

        left_pane = tk.Frame(split_container, bg=TAB_ACTIVE_BG)
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_pane = tk.Frame(split_container, bg=TAB_ACTIVE_BG, padx=10, pady=10)
        right_pane.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        container = tk.Frame(left_pane, bg=TAB_ACTIVE_BG, padx=40, pady=20)
        container.pack(expand=True) 

        # Unified UI Mapping Dictionary
        self.fields = {
            "date": "製作日:", "product": "品名:", "filament": "素材:", 
            "color": "色:", "weight": "使用量 (g):", "class": "区分:", "maker": "製作者:"
        }
        self.entries = {}
        
        # --- FIXED ROW MAP: Unifies "date" loop creation at row index 0 ---
        row_map = ["date", "product", "time_placeholder", "filament", "color", "weight", "class", "maker"]
        
        for i, key in enumerate(row_map):
            if key == "time_placeholder":
                tk.Label(container, text="製作時間:", bg=TAB_ACTIVE_BG, font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
                time_frame = tk.Frame(container, bg=TAB_ACTIVE_BG)
                time_frame.grid(row=i, column=1, sticky="w", padx=10)
                
                self.hour_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.hour_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" h   ", bg=TAB_ACTIVE_BG, font=UI_CONFIG["entry_font"]).pack(side=tk.LEFT)
                
                self.min_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.min_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" m", bg=TAB_ACTIVE_BG, font=UI_CONFIG["entry_font"]).pack(side=tk.LEFT)
                continue

            # Standard Label Generation
            tk.Label(container, text=self.fields[key], bg=TAB_ACTIVE_BG, font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
            
            # --- FIXED SELECTION STRUCT (if / elif / else chain) ---
            if key == "date":
                ent = DateEntry(
                    container, 
                    width=UI_CONFIG["entry_width"] - 3, 
                    font=UI_CONFIG["entry_font"],
                    background='darkblue', 
                    foreground='white', 
                    borderwidth=2,
                    date_pattern='yyyy/m/d'
                )
                ent.set_date(datetime.now())
                self.date_entry = ent  # Maintained assignment shortcut safely
                
            elif key == "class":
                ent = ttk.Combobox(container, width=UI_CONFIG["entry_width"] - 3, font=UI_CONFIG["entry_font"], state="readonly")
                current_options = self.load_classifications()
                ent['values'] = tuple(current_options)
                if current_options: ent.current(0)
                
            elif key == "filament":
                ent = ttk.Combobox(container, width=UI_CONFIG["entry_width"] - 3, font=UI_CONFIG["entry_font"], state="readonly")
                current_materials = self.load_materials()
                ent['values'] = tuple(current_materials)
                if current_materials: ent.current(0)
                
            elif key == "color":
                ent = ttk.Combobox(container, width=UI_CONFIG["entry_width"] - 3, font=UI_CONFIG["entry_font"], state="readonly")
                current_colors = self.load_colors()
                ent['values'] = tuple(current_colors)
                if current_colors: ent.current(0)
                
            elif key == "maker":
                ent = ttk.Combobox(container, width=UI_CONFIG["entry_width"] - 3, font=UI_CONFIG["entry_font"], state="readonly")
                current_producers = self.load_producers()
                ent['values'] = tuple(current_producers)
                if current_producers: ent.current(0)
                
            else:
                ent = tk.Entry(container, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"])
                
            ent.grid(row=i, column=1, sticky="w", padx=10)
            self.entries[key] = ent

        # Department Code Frame Elements
        tk.Label(container, text="部署コード:", font=UI_CONFIG["label_font"], bg=TAB_ACTIVE_BG).grid(row=8, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
        self.code_var = tk.StringVar()
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display)
        self.code_entry = tk.Entry(container, textvariable=self.code_var, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"], bg="#F0F8FF")
        self.code_entry.grid(row=8, column=1, sticky="w", padx=10)

        self.result_label = tk.Label(container, text="結果 : --- ", font=TK_FONT_REGULAR, bg="#F9F9F9", width=UI_CONFIG["entry_width"] + 15)
        self.result_label.grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)

        self.save_button = tk.Button(container, text="データ保存", command=self.save_data, bg="#5CB85C", fg="white", width=30, font=TK_FONT_BOLD)
        self.save_button.grid(row=10, column=0, columnspan=2, pady=20)
        
        # --- CLEANED KEYPAD BINDINGS ---
        # Bound ONLY to writable text fields. Dropped dropdown menus and the calendar component.
        self.entries["product"].bind("<Button-1>", lambda e: self.show_keyboard(self.entries["product"]))
        self.hour_entry.bind("<Button-1>", lambda e: self.show_keyboard(self.hour_entry))
        self.min_entry.bind("<Button-1>", lambda e: self.show_keyboard(self.min_entry))
        self.entries["weight"].bind("<Button-1>", lambda e: self.show_keyboard(self.entries["weight"]))
        self.code_entry.bind("<Button-1>", lambda e: self.show_keyboard(self.code_entry))

        # Initialize and embed the keyboard UI
        self.keyboard_ui = VirtualKeyboard(right_pane, self.entries["product"])
        self.keyboard_ui.pack(fill=tk.BOTH, expand=True)

    def show_keyboard(self, entry):
        if entry == self.date_entry:
            self.on_date_focus_in()
        self.keyboard_ui.update_target(entry)

    def update_dept_display(self, *args):
        raw_text = self.code_var.get().replace("-", "")[:12]
        formatted = ""
        if len(raw_text) <= 4: formatted = raw_text
        elif len(raw_text) <= 8: formatted = f"{raw_text[:4]}-{raw_text[4:]}"
        else: formatted = f"{raw_text[:4]}-{raw_text[4:8]}-{raw_text[8:]}"

        self.code_var.trace_remove("write", self.trace_id)
        if self.code_var.get() != formatted:
            self.code_var.set(formatted)
            self.code_entry.after(1, lambda: self.code_entry.icursor(tk.END))
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display)

        code = formatted.strip()
        if code in self.dept_db:
            info = self.dept_db[code]
            self.result_label.config(text=f"結果 : {info['dept']} || {info['room']} || {info['group']}", fg="blue")
        else:
            self.result_label.config(text="結果 : コードが見つかりません", fg="red")

    def on_date_focus_in(self):
        if "YYYY" in self.date_entry.get():
            self.date_entry.delete(0, tk.END)
            self.date_entry.config(fg="black")

    def calculate_cost(self):
        try:
            weight_str = self.entries["weight"].get().replace("g", "").strip()
            weight_g = float(weight_str or 0)
            material = self.entries["filament"].get().upper().strip()
            
            unit_price = self.material_prices.get(material, 3000) 
            mat_cost = (weight_g / 1000) * unit_price
            
            hours = float(self.hour_entry.get().strip() or 0)
            minutes = float(self.min_entry.get().strip() or 0)
            total_hours = hours + (minutes / 60)
            elec_cost = total_hours * self.elec_cost_per_hour
            
            total = mat_cost + elec_cost
            return f"¥{total:.2f}"
        except ValueError:
            return "¥0.00"

    def save_data(self):
        code = self.code_var.get().strip()
        if code not in self.dept_db:
            messagebox.showerror("エラー", "有効な部署コードを入力してください。")
            return

        input_date = self.date_entry.get().strip()
        date = input_date if input_date and "YYYY" not in input_date else datetime.now().strftime("%Y/%m/%d")
        h = self.hour_entry.get().strip() or "0"
        m = self.min_entry.get().strip() or "0"
        formatted_time = f"{h} h {m} m"

        dept_info = self.dept_db[code]
        row_to_save = [
            date, self.entries["product"].get(), formatted_time,
            self.entries["filament"].get(), self.entries["color"].get(), 
            self.entries["weight"].get().replace(" g","").strip() + " g",
            self.entries["class"].get(), self.entries["maker"].get(),
            dept_info['dept'], dept_info['room'], dept_info['group'],
            self.calculate_cost()
        ]

        file_path = "printer_history.csv"
        header = ["日付", "品名", "時間", "素材", "色", "重量", "区分", "製作者", "部署", "室", "グループ", "コスト"]

        try:
            if self.editing_row_data:
                all_rows = []
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    all_rows = list(csv.reader(f))
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    for r in all_rows[1:]:
                        if [str(x) for x in r] == self.editing_row_data:
                            writer.writerow(row_to_save)
                        else:
                            writer.writerow(r)
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
            messagebox.showerror("エラー", "CSVファイルが開かれています。閉じてからやり直してください。")

    def clear_inputs(self):
        for key, e in self.entries.items():
            if key == "class":
                e.set("") # Safely clears a Combobox, or you can use e.current(0) to reset to default
            else:
                e.delete(0, tk.END)
        self.hour_entry.delete(0, tk.END)
        self.min_entry.delete(0, tk.END)
        self.code_var.set("")
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)")
        self.date_entry.config(fg="grey")
        self.save_button.config(text="データ保存", bg="#5CB85C")
        self.editing_row_data = None

    def setup_history_tab(self):
        cols = ("date", "prod", "time", "fil", "color", "weight", "class", "maker", "dept", "room", "group", "cost")
        self.tree = ttk.Treeview(self.history_frame, columns=cols, show='headings')
        headers = ["製作日", "品名", "時間", "素材", "色", "使用量", "区分", "製作者", "部署", "室", "グループ", "コスト"]
        for col, h in zip(cols, headers):
            self.tree.heading(col, text=h)
            if col == "date": self.tree.column(col, width=100, anchor="center")
            elif col == "prod": self.tree.column(col, width=180, anchor="w")
            else: self.tree.column(col, width=80, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        btn_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        btn_frame.pack(fill=tk.X, padx=10, side=tk.BOTTOM, pady=10)

        tk.Button(btn_frame, text="選択行を編集", command=self.prepare_edit, 
                  bg="#F0AD4E", fg="white", width=15, font=TK_FONT_TREEVIEW).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(btn_frame, text="選択行を削除", command=self.delete_entry, 
                  bg="#D9534F", fg="white", width=15, font=TK_FONT_TREEVIEW).pack(side=tk.RIGHT, padx=5)

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                data = list(csv.reader(f))[1:]
                for row in reversed(data): self.tree.insert("", tk.END, values=row)

    def prepare_edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("警告", "編集する行を選択してください。")
            return
        vals = self.tree.item(sel[0])['values']
        self.clear_inputs()
        self.editing_row_data = [str(v) for v in vals]
        
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, vals[0])
        self.date_entry.config(fg="black")
        self.entries["product"].insert(0, vals[1])

        time_str = str(vals[2])
        parts = time_str.split()
        self.hour_entry.insert(0, parts[0] if len(parts) > 0 else "0")
        self.min_entry.insert(0, parts[2] if len(parts) > 2 else "0")
        
        self.entries["filament"].insert(0, vals[3])
        self.entries["color"].insert(0, vals[4])
        self.entries["weight"].insert(0, str(vals[5]).replace(" g", ""))
        self.entries["class"].insert(0, vals[6])
        self.entries["maker"].insert(0, vals[7])
        
        for code, info in self.dept_db.items():
            if info['dept'] == vals[8] and info['room'] == vals[9]:
                self.code_var.set(code)
                break
        
        self.notebook.select(0)
        self.save_button.config(text="更新を保存", bg="#0275D8")

    def delete_entry(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("警告", "削除する行を選択してください。")
            return
        
        vals = [str(v) for v in self.tree.item(sel[0])['values']]
        if not messagebox.askyesno("確認", "選択した行を削除してもよろしいですか？"):
            return

        file_path = "printer_history.csv"
        try:
            all_rows = []
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    header = next(reader)
                    for r in reader:
                        if [str(x) for x in r] != vals:
                            all_rows.append(r)

                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    writer.writerows(all_rows)

                messagebox.showinfo("成功", "データを削除しました。")
                self.load_history()
        except PermissionError:
            messagebox.showerror("エラー", "CSVファイルが開かれています。閉じてからやり直してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"削除中にエラーが発生しました: {e}")

    def get_available_months(self):
        months = ["全期間"]
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader)
                for row in reader:
                    if row:
                        month = row[0][:7] 
                        if month not in months: months.append(month)
        return sorted(months, reverse=True)

    def on_tab_change(self, event):
        idx = self.notebook.index(self.notebook.select())
        if idx == 2: 
            self.class_pie_tab(self.stats_class_frame, 6, "製作内容集計", self.class_month_var)
        elif idx == 3: 
            self.dept_pie_tab(self.stats_dept_frame)

    def class_pie_tab(self, frame, column_index, title, month_var):
        for widget in frame.winfo_children(): widget.destroy()
        
        ctrl_frame = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(ctrl_frame, text="表示月:", bg=TAB_ACTIVE_BG, font=TK_FONT_REGULAR).pack(side=tk.LEFT)
        month_filter = ttk.Combobox(ctrl_frame, textvariable=month_var, values=self.get_available_months(), state="readonly", font=TK_FONT_REGULAR)
        month_filter.pack(side=tk.LEFT, padx=5)
        month_filter.bind("<<ComboboxSelected>>", lambda e: self.class_pie_tab(frame, column_index, title, month_var))

        container = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        container.pack(fill=tk.BOTH, expand=True)

        left_side = tk.Frame(container, bg=TAB_ACTIVE_BG)
        left_side.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_side = tk.Frame(container, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
        right_side.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=20)

        data_counts, total_count, selected_month = {}, 0, month_var.get()
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader)
                for row in reader:
                    if selected_month != "全期間" and not row[0].startswith(selected_month): continue
                    if len(row) > column_index:
                        val = row[column_index]
                        if val and val != "N/A":
                            data_counts[val] = data_counts.get(val, 0) + 1
                            total_count += 1

        if not data_counts:
            tk.Label(left_side, text="指定期間のデータがありません", bg=TAB_ACTIVE_BG, font=TK_FONT_REGULAR).pack(pady=50)
            return

        fig, ax = plt.subplots(figsize=(6, 6), dpi=100)
        fig.patch.set_facecolor(TAB_ACTIVE_BG)
        colors = plt.cm.Set2(np.linspace(0, 1, len(data_counts)))
        
        font_name_target = "MS Gothic" if os.name == "nt" else default_linux_font
        wedges, texts, autotexts = ax.pie(
            data_counts.values(), 
            labels=data_counts.keys(), 
            autopct='%1.1f%%', 
            startangle=90, 
            colors=colors,
            pctdistance=0.75, 
            wedgeprops={'width': 0.45, 'edgecolor': 'w', 'linewidth': 2}, 
            textprops={'fontsize': 11, 'fontname': font_name_target} 
        )
        
        for autotext in autotexts:
            autotext.set_color('black') 
            autotext.set_fontsize(12)      
            autotext.set_fontweight('bold') 
        
        center_text = "合計製作数" if selected_month == "全期間" else f"{selected_month}\n製作数"
        ax.text(0, 0, f"{center_text}\n{total_count}個", ha='center', va='center', 
                fontweight='bold', fontsize=13, fontname=font_name_target)
        
        ax.set_title(title, pad=20, fontname=font_name_target, fontsize=15, fontweight='bold')
        
        canvas = FigureCanvasTkAgg(fig, master=left_side)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

        style = ttk.Style()
        style.configure("Class.Treeview", font=TK_FONT_REGULAR, rowheight=30)
        tree = ttk.Treeview(right_side, columns=("label", "count"), show="headings", style="Class.Treeview")
        tree.heading("label", text="区分/項目")
        tree.heading("count", text="製作数 (個)")
        tree.column("label", width=150, anchor=tk.W)
        tree.column("count", width=100, anchor=tk.E)
        tree.pack(fill=tk.BOTH, expand=True)

        for label, count in sorted(data_counts.items(), key=lambda item: item[1], reverse=True):
            tree.insert("", tk.END, values=(label, f"{count} 個"))

    def dept_pie_tab(self, frame):
        for w in frame.winfo_children(): w.destroy()

        ctrl_frame = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(ctrl_frame, text="表示月:", bg=TAB_ACTIVE_BG, font=TK_FONT_REGULAR).pack(side=tk.LEFT)
        
        month_filter = ttk.Combobox(ctrl_frame, textvariable=self.dept_month_var, 
                                    values=self.get_available_months(), state="readonly", font=TK_FONT_REGULAR)
        month_filter.pack(side=tk.LEFT, padx=5)
        month_filter.bind("<<ComboboxSelected>>", lambda e: self.dept_pie_tab(frame))

        content_container = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        content_container.pack(fill=tk.BOTH, expand=True)

        left_side = tk.Frame(content_container, bg=TAB_ACTIVE_BG)
        left_side.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_side = tk.Frame(content_container, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
        right_side.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 10), pady=10)

        dept_data, total_count, selected_month = {}, 0, self.dept_month_var.get()
        log_report = {} 

        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if selected_month != "全期間" and not row["日付"].startswith(selected_month): 
                        continue
                    
                    d, r = row.get("部署"), row.get("室")
                    if d:
                        if d not in dept_data: dept_data[d] = {}
                        dept_data[d][r] = dept_data[d].get(r, 0) + 1
                        total_count += 1
                        
                        group_key = f"{d}({r})" if r else d
                        mat = row.get("素材", "不明").upper()
                        
                        try:
                            g_str = row.get("重量", "0").replace("g", "").replace(" ","").strip()
                            c_str = row.get("コスト", "0").replace("¥", "").replace(",","").strip()
                            grams = float(g_str) if g_str else 0
                            cost_val = float(c_str) if c_str else 0
                        except: 
                            grams, cost_val = 0, 0

                        if group_key not in log_report: 
                            log_report[group_key] = {}
                        if mat not in log_report[group_key]: 
                            log_report[group_key][mat] = {'g': 0, 'c': 0}
                        
                        log_report[group_key][mat]['g'] += grams
                        log_report[group_key][mat]['c'] += cost_val

        if not dept_data:
            tk.Label(left_side, text="指定期間のデータがありません", bg=TAB_ACTIVE_BG, font=TK_FONT_REGULAR).pack(pady=50)
            return

        fig, ax = plt.subplots(figsize=(5.5, 5.5), dpi=100)
        fig.patch.set_facecolor(TAB_ACTIVE_BG)
        base_colors = {"技術部": plt.cm.Blues, "製造部": plt.cm.Oranges, "その他": plt.cm.Greens}
        labels, sizes, colors = [], [], []

        for dept, rooms in dept_data.items():
            cmap = base_colors.get(dept, plt.cm.Greys)
            room_items = list(rooms.items())
            color_shades = cmap(np.linspace(0.4, 0.7, len(room_items)))
            for i, (room_name, count) in enumerate(room_items):
                labels.append(f"{dept}\n({room_name})" if room_name else dept)
                sizes.append(count)
                colors.append(color_shades[i])

        font_name_target = "MS Gothic" if os.name == "nt" else default_linux_font
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors, 
            pctdistance=0.75,
            wedgeprops={'width': 0.4, 'edgecolor': 'w'}, 
            textprops={'fontsize': 11, 'fontname': font_name_target} 
        )

        for autotext in autotexts:
            autotext.set_color('black')
            autotext.set_fontsize(14)
            autotext.set_fontweight('bold')
        
        center_text = "合計製作数" if selected_month == "全期間" else f"{selected_month}\n製作数"
        ax.text(0, 0, f"{center_text}\n{total_count}個", ha='center', va='center', 
                fontweight='bold', fontname=font_name_target, fontsize=12)
        
        ax.set_title("部署・室別利用率", pad=15, fontname=font_name_target, fontsize=14, fontweight='bold')
        
        canvas = FigureCanvasTkAgg(fig, master=left_side)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

        style = ttk.Style()
        style.configure("Treeview", font=TK_FONT_REGULAR, rowheight=28)
        style.configure("Treeview.Heading", font=(TK_FONT_REGULAR[0], 10, "bold"))

        columns = ("material", "weight", "cost")
        tree = ttk.Treeview(right_side, columns=columns, show="headings")
        
        tree.heading("material", text="項目/素材")
        tree.heading("weight", text="使用量 (g)")
        tree.heading("cost", text="コスト")

        tree.column("material", width=150, anchor=tk.W)
        tree.column("weight", width=80, anchor=tk.E)
        tree.column("cost", width=100, anchor=tk.E)

        scrollbar = ttk.Scrollbar(right_side, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        grand_cost, grand_weight = 0, 0

        for group, mats in sorted(log_report.items()):
            group_total_cost = sum(v['c'] for v in mats.values())
            tree.insert("", tk.END, values=(f"【{group}】", "", f"¥{group_total_cost:,.0f}"), tags=('header',))
            
            for mat, vals in mats.items():
                tree.insert("", tk.END, values=(f"  {mat}", f"{vals['g']:.1f}g", f"¥{vals['c']:,.0f}"))
                grand_cost += vals['c']
                grand_weight += vals['g']

        tree.tag_configure('header', background='#F0F0F0') 
        tree.insert("", tk.END, values=("", "", ""), tags=('sep',))
        tree.insert("", tk.END, values=("総合計", f"{grand_weight:.1f}g", f"¥{grand_cost:,.0f}"), tags=('total',))
        tree.tag_configure('total', background='#D1FFD1', font=(font_name_target, 10, "bold"))

    def setup_edit_tab(self):
        # Create a new frame container for the tab
        self.edit_frame = tk.Frame(self.notebook, bg="#F5F5F5")
        self.notebook.add(self.edit_frame, text=" リスト設定 ")

        split_container = tk.Frame(self.edit_frame, bg="#F5F5F5")
        split_container.pack(fill=tk.BOTH, expand=True)

        left_pane = tk.Frame(split_container, bg="#F5F5F5", padx=20, pady=20)
        left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_pane = tk.Frame(split_container, bg="#F5F5F5", padx=10, pady=10)
        right_pane.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- LEFT PANE SUB-TAB NOTEBOOK LAYOUT ---
        self.sub_notebook = ttk.Notebook(left_pane)
        self.sub_notebook.pack(fill=tk.BOTH, expand=True)

        # Setup Five Sub-Frames (Added Cost Panel)
        self.sub_class_frame = tk.Frame(self.sub_notebook, bg="#F5F5F5", pady=10)
        self.sub_mat_frame = tk.Frame(self.sub_notebook, bg="#F5F5F5", pady=10)
        self.sub_color_frame = tk.Frame(self.sub_notebook, bg="#F5F5F5", pady=10)
        self.sub_maker_frame = tk.Frame(self.sub_notebook, bg="#F5F5F5", pady=10)
        self.sub_cost_frame = tk.Frame(self.sub_notebook, bg="#F5F5F5", pady=10) # <-- New!

        self.sub_notebook.add(self.sub_class_frame, text=" 区分設定 ")
        self.sub_notebook.add(self.sub_mat_frame, text=" 素材設定 ")
        self.sub_notebook.add(self.sub_color_frame, text=" 色設定 ")
        self.sub_notebook.add(self.sub_maker_frame, text=" 製作者設定 ")
        self.sub_notebook.add(self.sub_cost_frame, text=" コスト設定 ") # <-- New!

        # Build UI onto frames using type tracking flags
        self.build_list_manager_ui(self.sub_class_frame, list_type="class")
        self.build_list_manager_ui(self.sub_mat_frame, list_type="material")
        self.build_list_manager_ui(self.sub_color_frame, list_type="color")
        self.build_list_manager_ui(self.sub_maker_frame, list_type="maker")
        self.build_list_manager_ui(self.sub_cost_frame, list_type="cost") # <-- New!

        # Set default focus field for keyboard to class input
        self.edit_keyboard_ui = VirtualKeyboard(right_pane, self.class_entry)
        self.edit_keyboard_ui.pack(fill=tk.BOTH, expand=True)

    def build_list_manager_ui(self, parent_frame, list_type="class"):
        """Factory method structural pattern modified to handle core text logs and numeric cost management."""
        
        if list_type == "cost":
            tk.Label(parent_frame, text="素材価格と電気料金の管理", font=TK_FONT_BOLD, bg="#F5F5F5").pack(anchor="w", pady=(0, 15))
            form_frame = tk.Frame(parent_frame, bg="#F5F5F5")
            form_frame.pack(fill=tk.BOTH, expand=True)

            # --- 1. Electricity Entry Box ---
            tk.Label(form_frame, text="電気料金単価 (円/時間):", font=TK_FONT_REGULAR, bg="#F5F5F5").grid(row=0, column=0, sticky="w", pady=5)
            self.elec_entry = tk.Entry(form_frame, font=TK_FONT_REGULAR, width=15)
            self.elec_entry.insert(0, str(self.elec_cost_per_hour))
            self.elec_entry.grid(row=0, column=1, sticky="w", padx=10, pady=5)
            
            # Switch keyboard and select text on click
            self.elec_entry.bind("<Button-1>", lambda e: [
                self.show_edit_keyboard(self.elec_entry),
                self.elec_entry.selection_range(0, tk.END),
                self.elec_entry.icursor(tk.END)
            ])
            # FIXED: If left empty on focus out, restore original value
            self.elec_entry.bind("<FocusOut>", lambda e: self.elec_entry.insert(0, str(self.elec_cost_per_hour)) if not self.elec_entry.get().strip() else None)

            # --- Material Selector Dropdown ---
            tk.Label(form_frame, text="対象の素材:", font=TK_FONT_REGULAR, bg="#F5F5F5").grid(row=1, column=0, sticky="w", pady=(20, 5))
            self.cost_mat_combo = ttk.Combobox(form_frame, font=TK_FONT_REGULAR, width=13, state="readonly")
            self.cost_mat_combo['values'] = tuple(self.load_materials())
            if self.load_materials(): self.cost_mat_combo.current(0)
            self.cost_mat_combo.grid(row=1, column=1, sticky="w", padx=10, pady=(20, 5))
            self.cost_mat_combo.bind("<<ComboboxSelected>>", lambda e: self.on_cost_material_selected())

            # --- 2. Material Price Entry Box ---
            tk.Label(form_frame, text="素材単価 (円/kg):", font=TK_FONT_REGULAR, bg="#F5F5F5").grid(row=2, column=0, sticky="w", pady=5)
            self.mat_price_entry = tk.Entry(form_frame, font=TK_FONT_REGULAR, width=15)
            self.mat_price_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
            
            # Switch keyboard and select text on click
            self.mat_price_entry.bind("<Button-1>", lambda e: [
                self.show_edit_keyboard(self.mat_price_entry),
                self.mat_price_entry.selection_range(0, tk.END),
                self.mat_price_entry.icursor(tk.END)
            ])
            # FIXED: If left empty on focus out, restore original value for selected material
            self.mat_price_entry.bind("<FocusOut>", lambda e: self.restore_empty_material_price())
            
            self.on_cost_material_selected()

            tk.Button(form_frame, text="料金設定を保存", font=TK_FONT_BOLD, bg="#5CB85C", fg="white", width=20, 
                      command=self.save_pricing_action).grid(row=3, column=0, columnspan=2, pady=30, sticky="w")
            return
            
        title_map = {
            "class": "区分リストの編集", "material": "素材リストの編集",
            "color": "色リストの編集", "maker": "製作者リストの編集"
        }
        tk.Label(parent_frame, text=title_map.get(list_type, "リスト編集"), font=TK_FONT_BOLD, bg="#F5F5F5").pack(anchor="w", pady=(0, 10))

        layout_container = tk.Frame(parent_frame, bg="#F5F5F5")
        layout_container.pack(fill=tk.BOTH, expand=True)

        # FIXED FRAME PACKING: Control frame packed first to claim its right-side real estate
        control_frame = tk.Frame(layout_container, bg="#F5F5F5")
        control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))

        list_frame = tk.Frame(layout_container, bg="#F5F5F5")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        lbox = tk.Listbox(list_frame, font=TK_FONT_REGULAR, width=22, height=10)
        lbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=lbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        lbox.config(yscrollcommand=scrollbar.set)

        tk.Label(control_frame, text="新しく追加する項目:", font=TK_FONT_REGULAR, bg="#F5F5F5").pack(anchor="w")
        ent = tk.Entry(control_frame, font=TK_FONT_REGULAR, width=18)
        ent.pack(anchor="w", pady=(5, 15))
        ent.bind("<Button-1>", lambda e: self.show_edit_keyboard(e.widget))

        # Dynamically save references to class properties using unified naming arrays
        if not hasattr(self, 'list_entries'): self.list_entries = {}
        if not hasattr(self, 'list_boxes'): self.list_boxes = {}
        
        self.list_entries[list_type] = ent
        self.list_boxes[list_type] = lbox

        # Backward compatibility tracking hooks
        if list_type == "class": self.class_listbox, self.class_entry = lbox, ent
        elif list_type == "material": self.mat_listbox, self.mat_entry = lbox, ent
        elif list_type == "color": self.color_listbox, self.color_entry = lbox, ent
        elif list_type == "maker": self.maker_listbox, self.maker_entry = lbox, ent

        tk.Button(control_frame, text="項目を追加", font=TK_FONT_BOLD, bg="#5CB85C", fg="white", width=14, command=lambda t=list_type: self.add_item_action(t)).pack(fill=tk.X, pady=4)
        tk.Button(control_frame, text="選択項目を削除", font=TK_FONT_BOLD, bg="#D9534F", fg="white", width=14, command=lambda t=list_type: self.remove_item_action(t)).pack(fill=tk.X, pady=4)

        self.refresh_list_views()

    def restore_empty_material_price(self):
        """Restores the saved configuration price if the field is left completely empty."""
        if not self.mat_price_entry.get().strip():
            selected_mat = self.cost_mat_combo.get()
            if selected_mat and selected_mat in self.material_prices:
                saved_price = self.material_prices[selected_mat]
                self.mat_price_entry.insert(0, str(saved_price))

    def on_cost_material_selected(self):
        """Triggers whenever user changes material option dropdown inside cost management sub-tab."""
        selected_mat = self.cost_mat_combo.get()
        if not selected_mat: return
        
        # Look up price in app memory object; fall back to 0 if material is completely new
        current_price = self.material_prices.get(selected_mat, 0)
        self.mat_price_entry.delete(0, tk.END)
        self.mat_price_entry.insert(0, str(current_price))

    def save_pricing_action(self):
        """Gathers text field values inside cost frame and flushes updates directly to config disk."""
        try:
            new_elec = int(self.elec_entry.get().strip())
            selected_mat = self.cost_mat_combo.get()
            new_price = int(self.mat_price_entry.get().strip())
        except ValueError:
            # Native Error Pop-up
            messagebox.showerror("入力エラー", "エラー:\n金額には半角数字のみを入力してください。")
            return

        # Commit directly back to app memory structures
        self.elec_cost_per_hour = new_elec
        if selected_mat:
            self.material_prices[selected_mat] = new_price

        # Compile data array and push straight to file
        export_payload = {"electricity_per_hour": self.elec_cost_per_hour}
        export_payload.update(self.material_prices)
        self.save_pricing_config(export_payload)
        
        # --- NATIVE SUCCESS POP-UP ALERT ---
        messagebox.showinfo("成功", "データ保存しました。")

    def show_edit_keyboard(self, target_entry):
        """Dynamically shifts the virtual keyboard's focus to the active text field."""
        # Update the keyboard's internal target if the method exists
        if hasattr(self, 'edit_keyboard_ui'):
            # If your VirtualKeyboard class has a target update method or property:
            self.edit_keyboard_ui.target_entry = target_entry

    def refresh_list_views(self):
        """Syncs all four data pools dynamically onto app interfaces."""
        # 1. Classifications
        if hasattr(self, 'class_listbox'):
            self.class_listbox.delete(0, tk.END)
            try:
                classes = self.load_classifications()
                for item in classes: self.class_listbox.insert(tk.END, item)
                if hasattr(self, 'entries') and "class" in self.entries: self.entries["class"]['values'] = tuple(classes)
            except AttributeError: pass

        # 2. Materials
        if hasattr(self, 'mat_listbox'):
            self.mat_listbox.delete(0, tk.END)
            try:
                materials = self.load_materials()
                for item in materials: self.mat_listbox.insert(tk.END, item)
                if hasattr(self, 'entries') and "filament" in self.entries: self.entries["filament"]['values'] = tuple(materials)
            except AttributeError: pass

        # 3. Colors
        if hasattr(self, 'color_listbox'):
            self.color_listbox.delete(0, tk.END)
            try:
                colors = self.load_colors()
                for item in colors: self.color_listbox.insert(tk.END, item)
                if hasattr(self, 'entries') and "color" in self.entries: self.entries["color"]['values'] = tuple(colors)
            except AttributeError: pass

        # 4. Makers/Producers
        if hasattr(self, 'maker_listbox'):
            self.maker_listbox.delete(0, tk.END)
            try:
                producers = self.load_producers()
                for item in producers: self.maker_listbox.insert(tk.END, item)
                if hasattr(self, 'entries') and "maker" in self.entries: self.entries["maker"]['values'] = tuple(producers)
            except AttributeError: pass

    def add_item_action(self, list_type):
        """Fetches value from targeted sub-tab entry box and commits it to disk storage."""
        # Dynamic lookup from active structural components
        entry_widget = self.list_entries.get(list_type)
        if not entry_widget: return
        
        new_text = entry_widget.get().strip()
        if not new_text: return
        
        # Pull existing entries depending on mode
        if list_type == "class":
            items = self.load_classifications()
            if new_text not in items:
                items.append(new_text)
                self.save_classifications(items)
        elif list_type == "material":
            items = self.load_materials()
            if new_text not in items:
                items.append(new_text)
                self.save_materials(items)
                if hasattr(self, 'cost_mat_combo'):
                    self.cost_mat_combo['values'] = tuple(items)
        elif list_type == "color":
            items = self.load_colors()
            if new_text not in items:
                items.append(new_text)
                self.save_colors(items)
        elif list_type == "maker":
            items = self.load_producers()
            if new_text not in items:
                items.append(new_text)
                self.save_producers(items)
                
        entry_widget.delete(0, tk.END)
        self.refresh_list_views()
        messagebox.showinfo("成功", "項目を追加しました。")

    def remove_item_action(self, list_type):
        """Deletes selected element inside current active tab listbox container."""
        box_widget = self.list_boxes.get(list_type)
        if not box_widget or not box_widget.curselection(): return
        
        selected_index = box_widget.curselection()[0]
        selected_text = box_widget.get(selected_index)
        
        if list_type == "class":
            items = self.load_classifications()
            if selected_text in items: items.remove(selected_text)
            self.save_classifications(items)
        elif list_type == "material":
            items = self.load_materials()
            if selected_text in items: items.remove(selected_text)
            self.save_materials(items)
        elif list_type == "color":
            items = self.load_colors()
            if selected_text in items: items.remove(selected_text)
            self.save_colors(items)
        elif list_type == "maker":
            items = self.load_producers()
            if selected_text in items: items.remove(selected_text)
            self.save_producers(items)
            
        self.refresh_list_views()
        messagebox.showinfo("成功", "選択した項目を削除しました。")

class VirtualKeyboard(tk.Frame):
    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
        self.configure(bg=BG_COLOR)

        # Handwriting data
        self.strokes = []
        self.current_stroke = []
        self.last_x, self.last_y = None, None
        self.recognition_timer = None

        self.shift_active = False
        self.qwerty_buttons = [] # Store references to letter buttons to update case

        self.setup_ui()

    def update_target(self, target_entry):
        self.target_entry = target_entry
        self.target_entry.focus_set()

    def setup_ui(self):
        # The tabs for switching between layouts (1 and 2 in your image)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_jp = tk.Frame(self.notebook, bg="#F5F5F5")
        self.tab_en = tk.Frame(self.notebook, bg="#F5F5F5")

        self.notebook.add(self.tab_jp, text="   1   ")
        self.notebook.add(self.tab_en, text="   2   ")

        self.build_tab_1_custom()
        self.build_tab_2_qwerty()

    def build_tab_1_custom(self):
        # --- Top Blank Canvas Area ---
        canvas_frame = tk.Frame(self.tab_jp, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))
        """
        clear_text_btn = tk.Button(canvas_frame, text="✖", relief="flat", bg="white", font=("Arial", 12, "bold"), command=self.clear_field)
        clear_text_btn.pack(side="top", anchor="ne", padx=5, pady=5)
        """
        # Allow basic pen drawing in the blank space
        self.drawing_canvas = tk.Canvas(canvas_frame, bg="white", highlightthickness=0)
        self.drawing_canvas.pack(fill=tk.BOTH, expand=True)
        self.drawing_canvas.bind("<Button-1>", self.start_stroke)
        self.drawing_canvas.bind("<B1-Motion>", self.paint)
        self.drawing_canvas.bind("<ButtonRelease-1>", self.end_stroke)
        
        clear_canvas_btn = tk.Button(canvas_frame, text="削除", relief="flat", bg="white", command=self.clear_handwriting)
        clear_canvas_btn.pack(side="top", anchor="sw", padx=5, pady=5)

        # --- Suggestion Row (Replacing Punctuation) ---
        self.suggestion_frame = tk.Frame(self.tab_jp, bg="#F5F5F5")
        self.suggestion_frame.pack(fill=tk.X, padx=10, pady=5)
        self.suggestion_buttons = []
        
        # Create 8 placeholder buttons for suggestions
        for i in range(8):
            btn = tk.Button(self.suggestion_frame, text="", font=("MS Gothic", 14, "bold"), 
                            height=2, command=lambda v="": self.select_suggestion(v))
            btn.pack(side="left", expand=True, fill="both", padx=2)
            self.suggestion_buttons.append(btn)
        
        self.update_suggestion_ui(["、", "。", "？", "！", "：", "「", "」", "；"]) # Default view

        # --- Action Row ---
        action_frame = tk.Frame(self.tab_jp, bg="#F5F5F5")
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        tk.Button(action_frame, text="⌫", font=("Arial", 12, "bold"), height=2, command=self.backspace).pack(side="left", expand=True, fill="both", padx=2)
        tk.Button(action_frame, text="クリア (Clear)", font=("Arial", 11, "bold"), height=2, bg="#FFCDD2", command=self.clear_field).pack(side="left", expand=True, fill="both", padx=2)
        tk.Button(action_frame, text="日本語 (Space)", font=("Arial", 12), height=2, command=lambda: self.insert_text(" ")).pack(side="left", expand=True, fill="both", padx=10)
        tk.Button(action_frame, text="↵", font=("Arial", 12, "bold"), height=2, command=self.clear_handwriting).pack(side="left", expand=True, fill="both", padx=2)

    def build_tab_2_qwerty(self):
        keys = [
            ['`', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '-', '=', '⌫'],
            ['q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', '[', ']', '\\'],
            ['a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', ';', "'"],
            ['⇧', 'z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.', '/', '⇧'],
            ['Ctrl + Alt', 'Space', 'Clear', 'Ctrl + Alt']
        ]

        for row in keys:
            frame = tk.Frame(self.tab_en, bg="#F5F5F5")
            frame.pack(fill=tk.X, padx=10, pady=2)
            
            for key in row:
                if key == '⌫':
                    btn = tk.Button(frame, text=key, width=6, height=2, font=("Arial", 11), command=self.backspace)
                elif key == 'Space':
                    btn = tk.Button(frame, text="", width=30, height=2, font=("Arial", 11), command=lambda: self.insert_text(" "))
                elif key == 'Clear':
                    btn = tk.Button(frame, text="Clear", width=10, height=2, font=("Arial", 11), bg="#FFCDD2", command=self.clear_field)
                elif key == '⇧':
                    btn = tk.Button(frame, text=key, width=6, height=2, font=("Arial", 11), command=self.toggle_shift)
                elif key == 'Ctrl + Alt':
                    btn = tk.Button(frame, text=key, width=8, height=2, font=("Arial", 10), state=tk.DISABLED)
                else:
                    btn = tk.Button(frame, text=key, width=4, height=2, font=("Arial", 12), command=lambda v=key: self.insert_text(v))
                    if key.isalpha() and len(key) == 1:
                        self.qwerty_buttons.append({'button': btn, 'char': key})
                        
                btn.pack(side="left", expand=True, fill="both", padx=2)

    def start_stroke(self, event):
        self.last_x, self.last_y = event.x, event.y
        self.current_stroke = [[event.x], [event.y], [0]] # x, y, time (dummy)
        if self.recognition_timer:
            self.after_cancel(self.recognition_timer)

    def paint(self, event):
        if self.last_x and self.last_y:
            self.drawing_canvas.create_line(self.last_x, self.last_y, event.x, event.y, 
                                            width=3, capstyle=tk.ROUND, smooth=True, fill="black")
        self.last_x, self.last_y = event.x, event.y
        self.current_stroke[0].append(event.x)
        self.current_stroke[1].append(event.y)
        self.current_stroke[2].append(0)

    def end_stroke(self, event):
        self.strokes.append(self.current_stroke)
        self.last_x, self.last_y = None, None
        # Wait 500ms after drawing stops to recognize
        self.recognition_timer = self.after(500, self.fetch_recognition)

    def fetch_recognition(self):
        if not self.strokes: return

        if requests is None:
            self.update_suggestion_ui(["Install", "requests", "via pip", "to use", "handwriting", "", "", ""])
            return
        
        url = "https://www.google.com.tw/inputtools/request?ime=handwriting&app=mobilesearch&cs=1&oe=UTF-8"
        payload = {
            "options": "enable_pre_space",
            "requests": [{
                "writing_guide": {"writing_area_width": self.drawing_canvas.winfo_width(), "writing_area_height": self.drawing_canvas.winfo_height()},
                "ink": self.strokes,
                "language": "ja"
            }]
        }
        
        try:
            response = requests.post(url, json=payload, timeout=2)
            if response.status_code == 200:
                results = response.json()
                candidates = results[1][0][1]
                self.update_suggestion_ui(candidates[:8])
        except Exception as e:
            print(f"Recognition error: {e}")

    def update_suggestion_ui(self, candidates):
        # Fill buttons with candidates, hide if no candidate
        for i, btn in enumerate(self.suggestion_buttons):
            if i < len(candidates):
                char = candidates[i]
                btn.config(text=char, state=tk.NORMAL, command=lambda v=char: self.select_suggestion(v))
            else:
                btn.config(text="", state=tk.DISABLED)

    def select_suggestion(self, char):
        if char:
            self.insert_text(char)
            self.clear_handwriting()

    def clear_handwriting(self):
        self.drawing_canvas.delete("all")
        self.strokes = []
        if self.recognition_timer:
            self.after_cancel(self.recognition_timer)
        # Reset to default punctuations after clearing
        self.update_suggestion_ui(['、', '。', '？', '！', '：', '「', '」', '；'])

    def toggle_shift(self):
        self.shift_active = not self.shift_active
        for item in self.qwerty_buttons:
            new_char = item['char'].upper() if self.shift_active else item['char'].lower()
            item['button'].config(text=new_char)

    def clear_field(self):
        self.target_entry.delete(0, tk.END)
        self.target_entry.focus_set()

    def insert_text(self, char):
        if self.shift_active and char.isalpha():
            char = char.upper()
        self.target_entry.insert(tk.INSERT, char)
        self.target_entry.focus_set()
    
    def backspace(self):
        """Removes a single character behind the current cursor inside the active entry."""
        if hasattr(self, 'target_entry') and self.target_entry:
            current_text = self.target_entry.get()
            if current_text:
                # Find out where the typing cursor currently is
                cursor_pos = self.target_entry.index(tk.INSERT)
                
                if cursor_pos > 0:
                    # Delete the character right before the cursor
                    self.target_entry.delete(cursor_pos - 1, cursor_pos)
                else:
                    # Fallback: if cursor tracking is hidden, shave off the last element
                    self.target_entry.delete(len(current_text) - 1, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
