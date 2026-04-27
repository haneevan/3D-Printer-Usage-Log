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
        self.root.geometry("1150x850") 
        self.root.configure(bg=BG_COLOR)
        
        self.editing_row_data = None 
        self.dept_db = self.load_dept_database()

        # Material Prices per kg
        self.material_prices = {
            "PLA": 2800, "ABS": 3400, "PC": 7280, "PET-CF": 18360
        }
        self.elec_cost_per_hour = 15 

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
        header = tk.Frame(self.root, bg=HEADER_COLOR, height=60)
        header.pack(fill=tk.X)
        tk.Label(header, text="3D プリンタ 製作記録", font=("MS Gothic", 18, "bold"), fg="white", bg=HEADER_COLOR).pack(expand=True)

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
            "entry_width": 55,
            "row_padding": 12,
            "label_font": ("MS Gothic", 11, "bold"),
            "entry_font": ("Arial", 11)
        }

        container = tk.Frame(self.input_frame, bg=TAB_ACTIVE_BG, padx=40, pady=20)
        container.pack(expand=True) 

        tk.Label(container, text="製作日:", bg=TAB_ACTIVE_BG, font=UI_CONFIG["label_font"]).grid(row=0, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
        self.date_entry = tk.Entry(container, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"], fg="grey")
        self.date_entry.insert(0, "YYYY/MM/DD (空欄で今日)")
        self.date_entry.grid(row=0, column=1, sticky="w", padx=10)
        self.date_entry.bind("<FocusIn>", lambda e: self.on_date_focus_in())

        self.fields = {
            "product": "品名:", "filament": "素材:", "color": "色:", 
            "weight": "使用量 (g):", "class": "区分:", "maker": "製作者:"
        }
        self.entries = {}
        
        row_map = ["product", "time_placeholder", "filament", "color", "weight", "class", "maker"]
        for i, key in enumerate(row_map, start=1):
            if key == "time_placeholder":
                tk.Label(container, text="製作時間:", bg=TAB_ACTIVE_BG, font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
                time_frame = tk.Frame(container, bg=TAB_ACTIVE_BG)
                time_frame.grid(row=i, column=1, sticky="w", padx=10)
                self.hour_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.hour_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" h   ", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
                self.min_entry = tk.Entry(time_frame, width=10, font=UI_CONFIG["entry_font"])
                self.min_entry.pack(side=tk.LEFT)
                tk.Label(time_frame, text=" m", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
                continue

            tk.Label(container, text=self.fields[key], bg=TAB_ACTIVE_BG, font=UI_CONFIG["label_font"]).grid(row=i, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
            ent = tk.Entry(container, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"])
            ent.grid(row=i, column=1, sticky="w", padx=10)
            self.entries[key] = ent

        tk.Label(container, text="部署コード:", font=UI_CONFIG["label_font"], bg=TAB_ACTIVE_BG).grid(row=8, column=0, sticky="w", pady=UI_CONFIG["row_padding"])
        self.code_var = tk.StringVar()
        self.trace_id = self.code_var.trace_add("write", self.update_dept_display)
        self.code_entry = tk.Entry(container, textvariable=self.code_var, width=UI_CONFIG["entry_width"], font=UI_CONFIG["entry_font"], bg="#F0F8FF")
        self.code_entry.grid(row=8, column=1, sticky="w", padx=10)

        self.result_label = tk.Label(container, text="結果 : --- ", font=("MS Gothic", 11), bg="#F9F9F9", width=UI_CONFIG["entry_width"] + 15)
        self.result_label.grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)

        self.save_button = tk.Button(container, text="データ保存", command=self.save_data, bg="#5CB85C", fg="white", width=30, font=("Arial", 11, "bold"))
        self.save_button.grid(row=10, column=0, columnspan=2, pady=20)

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
        for e in self.entries.values(): e.delete(0, tk.END)
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
        
        # --- Button Frame ---
        btn_frame = tk.Frame(self.history_frame, bg=TAB_ACTIVE_BG)
        btn_frame.pack(fill=tk.X, padx=10, side=tk.BOTTOM, pady=10)

        # Edit Button
        tk.Button(btn_frame, text="選択行を編集", command=self.prepare_edit, 
                  bg="#F0AD4E", fg="white", width=15, font=("MS Gothic", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        
        # Delete Button (New)
        tk.Button(btn_frame, text="選択行を削除", command=self.delete_entry, 
                  bg="#D9534F", fg="white", width=15, font=("MS Gothic", 9, "bold")).pack(side=tk.RIGHT, padx=5)

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
        
        # Get the data of the selected row
        vals = [str(v) for v in self.tree.item(sel[0])['values']]
        
        # Confirm deletion
        if not messagebox.askyesno("確認", "選択した行を削除してもよろしいですか？"):
            return

        file_path = "printer_history.csv"
        try:
            all_rows = []
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    header = next(reader)
                    # Filter out the row that matches our selection
                    for r in reader:
                        if [str(x) for x in r] != vals:
                            all_rows.append(r)

                # Rewrite the file without the deleted row
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    writer.writerows(all_rows)

                messagebox.showinfo("成功", "データを削除しました。")
                self.load_history() # Refresh the list
            
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
        
        # --- Control Header ---
        ctrl_frame = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(ctrl_frame, text="表示月:", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
        month_filter = ttk.Combobox(ctrl_frame, textvariable=month_var, values=self.get_available_months(), state="readonly")
        month_filter.pack(side=tk.LEFT, padx=5)
        month_filter.bind("<<ComboboxSelected>>", lambda e: self.class_pie_tab(frame, column_index, title, month_var))

        # --- Main Container ---
        container = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        container.pack(fill=tk.BOTH, expand=True)

        left_side = tk.Frame(container, bg=TAB_ACTIVE_BG)
        left_side.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_side = tk.Frame(container, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
        right_side.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20, pady=20)

        # --- Data Processing ---
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
            tk.Label(left_side, text="指定期間のデータがありません", bg=TAB_ACTIVE_BG).pack(pady=50)
            return

        # --- Chart (Left) ---
        fig, ax = plt.subplots(figsize=(6, 6), dpi=100) # Increased figure size
        fig.patch.set_facecolor(TAB_ACTIVE_BG)
        
        # CHANGED: Using 'Set2' for much more vibrant, saturated colors
        # Alternatives: plt.cm.Accent, plt.cm.Dark2, or plt.cm.tab10
        colors = plt.cm.Set2(np.linspace(0, 1, len(data_counts)))
        
        wedges, texts, autotexts = ax.pie(
            data_counts.values(), 
            labels=data_counts.keys(), 
            autopct='%1.1f%%', 
            startangle=90, 
            colors=colors,
            pctdistance=0.75, # Adjusted for fatter donut visibility
            wedgeprops={'width': 0.45, 'edgecolor': 'w', 'linewidth': 2}, # Slightly fatter, cleaner edges
            textprops={'fontsize': 11, 'fontname': "MS Gothic"} 
        )
        
        # Percentage Visibility Tuning
        for autotext in autotexts:
            autotext.set_color('black') 
            autotext.set_fontsize(12)      # Bigger font
            autotext.set_fontweight('bold') 
        
        # Center Text Tuning
        center_text = "合計製作数" if selected_month == "全期間" else f"{selected_month}\n製作数"
        ax.text(0, 0, f"{center_text}\n{total_count}個", ha='center', va='center', 
                fontweight='bold', fontsize=13, fontname="MS Gothic")
        
        ax.set_title(title, pad=20, fontname="MS Gothic", fontsize=15, fontweight='bold')
        
        canvas = FigureCanvasTkAgg(fig, master=left_side)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

        # --- Data Table (Right) ---
        # (Table logic remains same, but increased height/font for visibility)
        style = ttk.Style()
        style.configure("Class.Treeview", font=("MS Gothic", 10), rowheight=30)
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

        # --- Control Header ---
        ctrl_frame = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(ctrl_frame, text="表示月:", bg=TAB_ACTIVE_BG).pack(side=tk.LEFT)
        
        month_filter = ttk.Combobox(ctrl_frame, textvariable=self.dept_month_var, 
                                    values=self.get_available_months(), state="readonly")
        month_filter.pack(side=tk.LEFT, padx=5)
        month_filter.bind("<<ComboboxSelected>>", lambda e: self.dept_pie_tab(frame))

        # --- Main Content Container ---
        content_container = tk.Frame(frame, bg=TAB_ACTIVE_BG)
        content_container.pack(fill=tk.BOTH, expand=True)

        left_side = tk.Frame(content_container, bg=TAB_ACTIVE_BG)
        left_side.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_side = tk.Frame(content_container, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
        right_side.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 10), pady=10)

        # --- Data Processing ---
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
            tk.Label(left_side, text="指定期間のデータがありません", bg=TAB_ACTIVE_BG).pack(pady=50)
            return

        # --- Draw Pie Chart (Left) ---
        fig, ax = plt.subplots(figsize=(5.5, 5.5), dpi=100) # Increased figure size
        fig.patch.set_facecolor(TAB_ACTIVE_BG)
        base_colors = {"技術部": plt.cm.Blues, "製造部": plt.cm.Oranges, "その他": plt.cm.Greens}
        labels, sizes, colors = [], [], []

        for dept, rooms in dept_data.items():
            cmap = base_colors.get(dept, plt.cm.Greys)
            room_items = list(rooms.items())
            color_shades = cmap(np.linspace(0.4, 0.7, len(room_items))) # Vibrant range
            for i, (room_name, count) in enumerate(room_items):
                labels.append(f"{dept}\n({room_name})" if room_name else dept)
                sizes.append(count)
                colors.append(color_shades[i])

        # Enhanced Pie Parameters
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors, 
            pctdistance=0.75,
            wedgeprops={'width': 0.4, 'edgecolor': 'w'}, # Fattened donut
            textprops={'fontsize': 11, 'fontname': "MS Gothic"} # Boosted labels
        )

        # Make percentages more visible
        for autotext in autotexts:
            autotext.set_color('black')
            autotext.set_fontsize(14)
            autotext.set_fontweight('bold')
        
        center_text = "合計製作数" if selected_month == "全期間" else f"{selected_month}\n製作数"
        ax.text(0, 0, f"{center_text}\n{total_count}個", ha='center', va='center', 
                fontweight='bold', fontname="MS Gothic", fontsize=12) # Boosted center
        
        ax.set_title("部署・室別利用率", pad=15, fontname="MS Gothic", fontsize=14, fontweight='bold')
        
        canvas = FigureCanvasTkAgg(fig, master=left_side)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

        # --- Draw Detailed Log Treeview (Right) ---
        style = ttk.Style()
        style.configure("Treeview", font=("MS Gothic", 10), rowheight=28) # Cleaned up table fonts
        style.configure("Treeview.Heading", font=("MS Gothic", 10, "bold"))

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

        tree.tag_configure('header', background='#F0F0F0') # Slightly lighter header
        tree.insert("", tk.END, values=("", "", ""), tags=('sep',))
        tree.insert("", tk.END, values=("総合計", f"{grand_weight:.1f}g", f"¥{grand_cost:,.0f}"), tags=('total',))
        tree.tag_configure('total', background='#D1FFD1', font=("MS Gothic", 10, "bold"))

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()