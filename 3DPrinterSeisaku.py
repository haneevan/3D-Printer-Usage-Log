import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

# Fix Japanese font display for Matplotlib charts
import matplotlib
matplotlib.rc('font', family='MS Gothic')

# --- Visual Styling ---
BG_COLOR = "#E6E8EA"         
HEADER_COLOR = "#7F8387"    
TAB_ACTIVE_BG = "white"      

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3D プリンタ 製作記録")
        self.root.geometry("1100x850") # Wide enough for dashboard
        self.root.configure(bg=BG_COLOR)
        
        self.editing_row_data = None # Stores data of the row being edited
        self.dept_db = self.load_dept_database()

        # Material Prices per kg (from your Excel formulas)
        self.material_prices = {
            "PLA": 2800, "ABS": 3400, "PC": 7280, "PET-CF": 18360
        }
        self.elec_cost_per_hour = 15 

        # UI Setup
        self.setup_styles()
        self.setup_main_ui()
        self.load_history()

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

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=35, font=("Arial", 9))
        style.configure("Treeview.Heading", font=("MS Gothic", 10, "bold"))

    def setup_main_ui(self):
        # 1. Header (Common for all tabs)
        header = tk.Frame(self.root, bg=HEADER_COLOR, height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="3D プリンタ 製作記録", font=("MS Gothic", 18, "bold"), fg="white", bg=HEADER_COLOR).pack(expand=True)

        # 2. Create the Notebook (must be done before adding frames)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 3. Create frames for each tab
        self.input_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.history_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.stats_class_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        self.stats_dept_frame = tk.Frame(self.notebook, bg=TAB_ACTIVE_BG)
        
        # 4. Add frames to notebook with localized tab text
        self.notebook.add(self.input_frame, text=" 入力 ")
        self.notebook.add(self.history_frame, text=" 履歴 ")
        self.notebook.add(self.stats_class_frame, text=" 製作内容集計 ")
        self.notebook.add(self.stats_dept_frame, text=" 部署別利用率 ")

        # 5. Initialize the specific contents of each tab
        self.setup_input_tab()
        self.setup_history_tab()
        
        # 6. Bind notebook selection to automatically refresh charts
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    # ==================== INPUT TAB LOGIC (Restored Style) ====================
    def setup_input_tab(self):
        # --- UI CONFIGURATION ---
        UI_CONFIG = {
            "entry_width": 55,       # Width of input boxes
            "row_padding": 12,       # Spacing between rows
            "label_font": ("MS Gothic", 11, "bold"),
            "entry_font": ("Arial", 11)
        }

        # Frame centering logic
        container = tk.Frame(self.input_frame, bg=TAB_ACTIVE_BG, padx=40, pady=20)
        container.pack(expand=True) 

        # 1. Date Field
        tk.Label(container, text="製作日:", bg=TAB_ACTIVE_BG, 
                 font=UI_CONFIG["label_font"]).grid(row=0, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
        self.date_entry = tk.Entry(container, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"], fg="grey")
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)")
        self.date_entry.grid(row=0, column=1, sticky="w", padx=10)
        self.date_entry.bind("<FocusIn>", lambda e: self.on_date_focus_in())

        # 2. Main Data Fields (The Loop)
        self.fields = {
            "product": "品名:", "filament": "素材:", "color": "色:", 
            "weight": "使用量 (g):", "class": "区分:", "maker": "製作者:"
        }
        self.entries = {}
        
        row_map = ["product", "time_placeholder", "filament", "color", "weight", "class", "maker"]
        for i, key in enumerate(row_map, start=1):
            if key == "time_placeholder":
                tk.Label(container, text="製作時間:", bg=TAB_ACTIVE_BG, 
                         font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
                time_frame = tk.Frame(container, bg=TAB_ACTIVE_BG)
                time_frame.grid(row=i, column=1, sticky="w", padx=10)
                
                self.hour_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.hour_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" h   ", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
                
                self.min_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.min_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" m", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
                continue

            tk.Label(container, text=self.fields[key], bg=TAB_ACTIVE_BG, 
                     font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
            
            ent = tk.Entry(container, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"])
            ent.grid(row=i, column=1, sticky="w", padx=10)
            self.entries[key] = ent

        # 3. Department Code (FIXED: Added Trace back in)
        tk.Label(container, text="部署コード:", font=UI_CONFIG["label_font"], bg=TAB_ACTIVE_BG).grid(row=8, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
        
        self.code_var = tk.StringVar()
        # --- THIS LINE RESTORES THE HYPHEN LOGIC ---
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display)
        
        self.code_entry = tk.Entry(container, textvariable=self.code_var, width=UI_CONFIG["entry_width"], 
                                   font=UI_CONFIG["entry_font"], bg="#F0F8FF")
        self.code_entry.grid(row=8, column=1, sticky="w", padx=10)

        # 4. Result Label
        self.result_label = tk.Label(container, text="結果 : --- ", font=("MS Gothic", 11), bg="#F9F9F9", width=UI_CONFIG["entry_width"] + 15)
        self.result_label.grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)

        # 5. Save Button
        self.save_button = tk.Button(container, text="データ保存", command=self.save_data, 
                                     bg="#5CB85C", fg="white", width=30, font=("Arial", 11, "bold"))
        self.save_button.grid(row=10, column=0, columnspan=2, pady=20)

    # ==================== GENERAL LOGIC (Consolidated) ====================
    def update_dept_display(self, *args):
        # Dynamic hyphen formatting (Good as is)
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

        # Results logic
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

    def save_data(self):
        code = self.code_var.get().strip()
        if code not in self.dept_db:
            messagebox.showerror("エラー", "有効な部署コードを入力してください。")
            return

        input_date = self.date_entry.get().strip()
        date = input_date if input_date and "YYYY" not in input_date else datetime.now().strftime("%Y/%m/%d")

        # Format the time string: e.g., "1 h 17 m"
        h = self.hour_entry.get().strip() or "0"
        m = self.min_entry.get().strip() or "0"
        formatted_time = f"{h} h {m} m"

        dept_info = self.dept_db[code]
        row_to_save = [
            date, 
            self.entries["product"].get(), 
            formatted_time,
            self.entries["filament"].get(), 
            self.entries["color"].get(), 
            self.entries["weight"].get() + " g",
            self.entries["class"].get(), 
            self.entries["maker"].get(),
            dept_info['dept'], 
            dept_info['room'], 
            dept_info['group'],
            self.calculate_cost() # Dynamic calculation and lock
        ]

        file_path = "printer_history.csv"
        # Column names (added 'color')
        header = ["日付", "品名", "時間", "素材", "色", "重量", "区分", "製作者", "部署", "室", "グループ", "コスト"]

        try:
            if self.editing_row_data:
                # Update logic
                all_rows = []
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    all_rows = list(csv.reader(f))
                
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    for r in all_rows[1:]:
                        if r == self.editing_row_data:
                            writer.writerow(row_to_save)
                        else:
                            writer.writerow(r)
                self.editing_row_data = None
                self.save_button.config(text="データ保存", bg="#5CB85C")
            else:
                # Simple append logic
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

    def calculate_cost(self):
        try:
            # 1. Get Weight (clean "g" if needed)
            weight_str = self.entries["weight"].get().replace(" g", "")
            weight_g = float(weight_str or 0)
            material = self.entries["filament"].get().upper()
            
            # 2. Get Material Cost
            unit_price = self.material_prices.get(material, 3000) 
            mat_cost = (weight_g / 1000) * unit_price
            
            # 3. Get Time and Electricity Cost
            hours = float(self.hour_entry.get() or 0)
            minutes = float(self.min_entry.get() or 0)
            total_hours = hours + (minutes / 60)
            elec_cost = total_hours * self.elec_cost_per_hour
            
            # 4. Total and lock
            total = mat_cost + elec_cost
            return f"¥{total:.2f}"
        except ValueError:
            return "¥0.00"

    def clear_inputs(self):
        # Clear main fields
        for e in self.entries.values(): 
            e.delete(0, tk.END)
        self.hour_entry.delete(0, tk.END); self.min_entry.delete(0, tk.END)
        self.code_var.set("")
        self.result_label.config(text="結果 : --- ", fg="#555") 
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)"); self.date_entry.config(fg="grey")
        self.save_button.config(text="データ保存", bg="#5CB85C")
        self.editing_row_data = None

    # ==================== HISTORY TAB LOGIC ====================
    def setup_history_tab(self):
        # Added color and cost columns
        cols = ("date", "prod", "time", "fil", "color", "weight", "class", "maker", "dept", "room", "group", "cost")
        self.tree = ttk.Treeview(self.history_frame, columns=cols, show='headings')
        
        headers = ["製作日", "品名", "時間", "素材", "色", "使用量", "区分", "製作者", "部署", "室", "グループ", "コスト"]
        for col, h in zip(cols, headers):
            self.tree.heading(col, text=h)
            if col == "date": self.tree.column(col, width=100, anchor="center")
            elif col == "prod": self.tree.column(col, width=200, anchor="w")
            else: self.tree.column(col, width=85, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        btn_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        btn_frame.pack(fill=tk.X, padx=10, side=tk.BOTTOM, pady=10)
        tk.Button(btn_frame, text="選択行を編集", command=self.prepare_edit, bg="#F0AD4E", width=15).pack(side=tk.RIGHT)

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                # list() avoids errors withDictReader during file ops
                data = list(csv.reader(f))[1:]
                for row in reversed(data): self.tree.insert("", tk.END, values=row)

    def prepare_edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("警告", "編集する行を選択してください。")
            return
            
        vals = self.tree.item(sel[0])['values']
        self.clear_inputs()
        
        # 1. Lock this row data for update comparison
        self.editing_row_data = [str(v) for v in vals]
        
        # 2. Date
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, vals[0])
        self.date_entry.config(fg="black")

        # 3. Product Name
        self.entries["product"].insert(0, vals[1])

        # 4. Time (Safely parse "X h Y m")
        time_str = str(vals[2])
        parts = time_str.split()
        h_val = parts[0] if len(parts) > 0 else "0"
        m_val = parts[2] if len(parts) > 2 else "0"
        self.hour_entry.insert(0, h_val)
        self.min_entry.insert(0, m_val)
        
        # 5. Filament, Color, Weight, Class, Maker
        self.entries["filament"].insert(0, vals[3])
        self.entries["color"].insert(0, vals[4]) # New Color Column
        self.entries["weight"].insert(0, str(vals[5]).replace(" g", ""))
        self.entries["class"].insert(0, vals[6])
        self.entries["maker"].insert(0, vals[7])
        
        # 6. Dept Code Lookup (Matching Dept AND Room for accuracy)
        found_code = False
        for code, info in self.dept_db.items():
            if info['dept'] == vals[8] and info['room'] == vals[9]:
                self.code_var.set(code)
                found_code = True
                break
        
        if not found_code: 
            self.code_var.set("") # Clear if not found instead of "N/A"

        # UI state change
        self.notebook.select(0)
        self.save_button.config(text="更新を保存", bg="#0275D8")

    # ==================== DASHBOARD LOGIC ====================
    def on_tab_change(self, event):
        selected_tab = self.notebook.index(self.notebook.select())
        # Production Content Tab (Index 6: Class)
        if selected_tab == 2: 
            self.draw_pie_chart(self.stats_class_frame, column_index=6, title="3Dプリント 製作内容集計")
        # Dept Usage Tab (Index 8: Dept)
        elif selected_tab == 3: 
            self.draw_pie_chart(self.stats_dept_frame, column_index=8, title="3Dプリント 部署別利用率")

    def draw_pie_chart(self, frame, column_index, title):
        # Clear frame to prevent layering
        for widget in frame.winfo_children():
            widget.destroy()

        data_counts = {}
        total_count = 0
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader) # Skip header
                for row in reader:
                    if len(row) > column_index:
                        val = row[column_index]
                        if val and val != "N/A": # Basic cleaning
                            data_counts[val] = data_counts.get(val, 0) + 1
                            total_count += 1

        if not data_counts:
            tk.Label(frame, text="データがありません", bg=TAB_ACTIVE_BG).pack(pady=50)
            return

        fig, ax = plt.subplots(figsize=(6, 5), dpi=100)
        fig.patch.set_facecolor(TAB_ACTIVE_BG)
        
        # Donut style
        ax.pie(data_counts.values(), labels=data_counts.keys(), autopct='%1.1f%%', 
               startangle=90, wedgeprops={'width':0.4}, pctdistance=0.8)
        
        # Central total count text
        ax.text(0, 0, f"累計製作数\n{total_count}個", ha='center', va='center', fontweight='bold', fontsize=12)
        ax.set_title(title, pad=20, fontname="MS Gothic", fontsize=14)
        
        # Embed
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig) # Free memory

    def draw_dept_pie_chart(self, frame):
        """Specially handled hierarchical pie chart for Departments and Rooms."""
        for w in frame.winfo_children(): w.destroy()
        
        dept_data = {} # Format: {Dept: {Room: Count}}
        total_count = 0
        
        if os.path.exists("printer_history.csv"):
            with open("printer_history.csv", "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader) 
                for row in reader:
                    if len(row) > 9:
                        d, r = row[8], row[9]
                        if d:
                            if d not in dept_data: dept_data[d] = {}
                            dept_data[d][r] = dept_data[d].get(r, 0) + 1
                            total_count += 1

        if not dept_data:
            tk.Label(frame, text="データがありません", bg=TAB_ACTIVE_BG).pack(pady=50)
            return

        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(TAB_ACTIVE_BG)

        # Define Base Colors
        base_colors = {
            "技術部": plt.cm.Blues,
            "製造部": plt.cm.Oranges,
            "その他": plt.cm.Greens
        }

        labels = []
        sizes = []
        colors = []

        for dept, rooms in dept_data.items():
            cmap = base_colors.get(dept, plt.cm.Greys)
            room_items = list(rooms.items())
            # Generate different depths of the same color family
            color_shades = cmap(np.linspace(0.4, 0.8, len(room_items)))
            
            for i, (room_name, count) in enumerate(room_items):
                # If room name is empty, label it with the Dept name
                display_label = f"{dept}\n({room_name})" if room_name else dept
                labels.append(display_label)
                sizes.append(count)
                colors.append(color_shades[i])

        # Draw the Pie Chart
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct='%1.1f%%', 
            startangle=90, colors=colors,
            wedgeprops={'width': 0.5, 'edgecolor': 'w'},
            pctdistance=0.75, textprops={'fontsize': 8, 'fontname': "MS Gothic"}
        )

        # Center text
        ax.text(0, 0, f"累計製作数\n{total_count}個", ha='center', va='center', 
                fontweight='bold', fontname="MS Gothic", fontsize=12)
        
        ax.set_title("3Dプリント 部署・室別利用率", pad=20, fontname="MS Gothic", fontsize=14)
        
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

    def on_tab_change(self, event):
        idx = self.notebook.index(self.notebook.select())
        if idx == 2: 
            self.draw_pie_chart(self.stats_class_frame, 6, "製作内容集計")
        elif idx == 3: 
            # Use the new specialized hierarchical chart
            self.draw_dept_pie_chart(self.stats_dept_frame)

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
