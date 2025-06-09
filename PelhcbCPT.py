import tkinter as tk
from tkinter import messagebox, ttk
import math
import pandas as pd
from tkinter import filedialog, font

# Color constants
PRIMARY_COLOR = "#165DFF"  # Primary color - Blue
SECONDARY_COLOR = "#4080FF"  # Secondary blue
ACCENT_COLOR = "#00B42A"  # No liquefaction - Green
WARNING_COLOR = "#FFC107"  # Slight liquefaction - Yellow
MEDIUM_COLOR = "#FF7D00"  # Moderate liquefaction - Orange
DANGER_COLOR = "#F53F3F"  # Severe liquefaction - Red
NEUTRAL_COLOR = "#F2F3F5"  # Neutral color
DARK_TEXT = "#1D2129"  # Dark text
LIGHT_TEXT = "#86909C"  # Light text
DELETE_BUTTON_COLOR = "#F53F3F"  # Delete button color - Red
BUTTON_BORDER_WIDTH = 2  # Button border width
BUTTON_BG_COLOR = "#FFFFFF"  # Button background color


class SoilLiquefactionCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Program for Earthquake-Induced Sand Liquefaction Hazard Classification Based on Cone Penetration Testing V1")
        self.root.geometry("1300x750")
        self.root.configure(bg=NEUTRAL_COLOR)

        # Set font
        self.default_font = font.nametofont("TkDefaultFont")
        self.default_font.configure(family="Segoe UI", size=11)
        self.root.option_add("*Font", self.default_font)

        # Configure ttk styles
        self.style = ttk.Style()

        # Main operation button style - Blue border, blue text, white background
        self.style.configure("MainButton.TButton", padding=8, relief="raised",
                             background=BUTTON_BG_COLOR, foreground=PRIMARY_COLOR,
                             font=("Segoe UI", 11, "bold"), borderwidth=BUTTON_BORDER_WIDTH,
                             bordercolor=PRIMARY_COLOR)
        self.style.map("MainButton.TButton",
                       background=[("active", "#E6F7FF"), ("pressed", "#CCE4FF")],
                       foreground=[("pressed", "#0E42D2"), ("!pressed", PRIMARY_COLOR)],
                       relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Delete point button style - Red border, red text, white background
        self.style.configure("DeleteButton.TButton", padding=10, relief="raised",
                             background=BUTTON_BG_COLOR, foreground=DELETE_BUTTON_COLOR,
                             font=("Segoe UI", 12, "bold"), borderwidth=BUTTON_BORDER_WIDTH,
                             bordercolor=DELETE_BUTTON_COLOR)
        self.style.map("DeleteButton.TButton",
                       background=[("active", "#FFF2F2"), ("pressed", "#FEE2E2")],
                       foreground=[("pressed", "#DC2626"), ("!pressed", DELETE_BUTTON_COLOR)],
                       relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Add soil layer button style - Blue border, blue text, white background
        self.style.configure("AddLayerButton.TButton", padding=10, relief="raised",
                             background=BUTTON_BG_COLOR, foreground=PRIMARY_COLOR,
                             font=("Segoe UI", 12, "bold"), borderwidth=BUTTON_BORDER_WIDTH,
                             bordercolor=PRIMARY_COLOR)
        self.style.map("AddLayerButton.TButton",
                       background=[("active", "#E6F7FF"), ("pressed", "#CCE4FF")],
                       foreground=[("pressed", "#0E42D2"), ("!pressed", PRIMARY_COLOR)],
                       relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Centered add button style
        self.style.configure("CenterAddButton.TButton", padding=10, relief="raised",
                             background=BUTTON_BG_COLOR, foreground=PRIMARY_COLOR,
                             font=("Segoe UI", 12, "bold"), borderwidth=BUTTON_BORDER_WIDTH,
                             bordercolor=PRIMARY_COLOR, anchor="center")
        self.style.map("CenterAddButton.TButton",
                       background=[("active", "#E6F7FF"), ("pressed", "#CCE4FF")],
                       foreground=[("pressed", "#0E42D2"), ("!pressed", PRIMARY_COLOR)],
                       relief=[("pressed", "sunken"), ("!pressed", "raised")])

        self.style.configure("TFrame", background=NEUTRAL_COLOR)
        self.style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"),
                             background=NEUTRAL_COLOR, foreground=DARK_TEXT)
        self.style.configure("Subheader.TLabel", font=("Segoe UI", 13, "bold"),
                             background=NEUTRAL_COLOR, foreground=DARK_TEXT)
        self.style.configure("Result.TLabel", font=("Segoe UI", 12),
                             background="white", foreground=DARK_TEXT)

        # Create UI
        self.create_ui()

    def create_ui(self):
        # Top title
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=20, pady=15)

        title_label = ttk.Label(header_frame, text="Earthquake Sandy Soil Liquefaction Risk Index Rating Program 1.0 Based on Static Cone Penetration Test", style="Header.TLabel")
        title_label.pack(side=tk.LEFT)

        # Main content area
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Left input area
        input_frame = ttk.LabelFrame(main_frame, text="Input Parameters", padding=15)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Create a container frame to hold the Canvas and scrollbars
        canvas_container = ttk.Frame(input_frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)

        # Create Canvas and scrollbars
        self.canvas = tk.Canvas(canvas_container, bg="white", highlightthickness=0)
        h_scrollbar = ttk.Scrollbar(canvas_container, orient="horizontal", command=self.canvas.xview)
        v_scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)

        # Grid layout configuration
        canvas_container.grid_rowconfigure(0, weight=1)
        canvas_container.grid_columnconfigure(0, weight=1)

        # Place Canvas and scrollbars
        self.canvas.grid(row=0, column=0, sticky="nsew")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")

        # Create frame to hold points
        self.points_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.points_frame, anchor="nw")

        # Bind event handlers
        self.points_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        # Add mouse wheel event handlers
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self.on_shift_mousewheel)

        # Independent button frame
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(fill=tk.X, pady=15)

        self.add_point_btn = ttk.Button(button_frame, text="Add New Point",
                                        command=self.add_new_point,
                                        style="MainButton.TButton")
        self.add_point_btn.pack(side=tk.LEFT, padx=5)

        self.calculate_btn = ttk.Button(button_frame, text="Calculate Liquefaction Level",
                                        command=self.calculate_all_points,
                                        style="MainButton.TButton")
        self.calculate_btn.pack(side=tk.LEFT, padx=15)

        self.import_btn = ttk.Button(button_frame, text="Import Excel File",
                                     command=self.import_excel,
                                     style="MainButton.TButton")
        self.import_btn.pack(side=tk.LEFT, padx=15)

        # New help button
        help_btn = ttk.Button(button_frame, text="Help",
                              command=self.show_help,
                              style="MainButton.TButton")
        help_btn.pack(side=tk.LEFT, padx=15)

        # Right result area
        result_frame = ttk.LabelFrame(main_frame, text="Calculation Results", padding=15)
        result_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))

        columns = ("Point", "PointName", "Layer", "Contribution", "Total_I", "Level")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        self.tree.heading("Point", text="Point Number")
        self.tree.heading("PointName", text="Point Name")
        self.tree.heading("Layer", text="Soil Layer")
        self.tree.heading("Contribution", text="Contribution Value")
        self.tree.heading("Total_I", text="Point Liquefaction Index")
        self.tree.heading("Level", text="Liquefaction Level")

        self.tree.column("Point", width=60, anchor=tk.CENTER)
        self.tree.column("PointName", width=100, anchor=tk.CENTER)
        self.tree.column("Layer", width=120, anchor=tk.CENTER)
        self.tree.column("Contribution", width=100, anchor=tk.CENTER)
        self.tree.column("Total_I", width=120, anchor=tk.CENTER)
        self.tree.column("Level", width=120, anchor=tk.CENTER)

        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scrollbar.set)

        # New clear button
        clear_btn = ttk.Button(result_frame, text="Clear Calculation Results", command=self.clear_results,
                               style="MainButton.TButton")
        clear_btn.pack(side=tk.TOP, pady=5, anchor="ne")

        # Create left raw data table
        self.raw_data_tree = ttk.Treeview(input_frame, columns=("Point", "q_c", "dw", "a_max", "ms", "R_f", "di", "zi"), show="headings")
        self.raw_data_tree.heading("Point", text="Point Number")
        self.raw_data_tree.heading("q_c", text="Measured Cone Tip Resistance (MPa)")
        self.raw_data_tree.heading("dw", text="Groundwater Level Depth (m)")
        self.raw_data_tree.heading("a_max", text="Peak Ground Acceleration (g)")
        self.raw_data_tree.heading("ms", text="Surface Wave Magnitude Ms")
        self.raw_data_tree.heading("R_f", text="Friction Ratio")
        self.raw_data_tree.heading("di", text="Liquefiable Layer Thickness (m)")
        self.raw_data_tree.heading("zi", text="Middle Depth of Liquefiable Layer (m)")
        self.raw_data_tree.column("Point", width=60, anchor=tk.CENTER)
        self.raw_data_tree.column("q_c", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("dw", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("a_max", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("ms", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("R_f", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("di", width=120, anchor=tk.CENTER)
        self.raw_data_tree.column("zi", width=120, anchor=tk.CENTER)
        self.raw_data_tree.pack_forget()  # Hide raw data table initially

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        if self.points_frame.winfo_reqwidth() > self.canvas.winfo_width():
            self.canvas.itemconfig(self.canvas_window, width=self.points_frame.winfo_reqwidth())
        else:
            self.canvas.itemconfig(self.canvas_window, width=self.canvas.winfo_width())

    def on_canvas_configure(self, event):
        if self.points_frame.winfo_reqwidth() > event.width:
            self.canvas.itemconfig(self.canvas_window, width=self.points_frame.winfo_reqwidth())
        else:
            self.canvas.itemconfig(self.canvas_window, width=event.width)

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def calculate_q_ccr(self, beta, a_max, dw, zi, R_f):
        term1 = beta * (35 * a_max) / (a_max + 0.17)
        term2 = (1 - 0.05 * dw)
        term3 = 0.1 + (0.9 * zi) / (zi + 6)
        term4 = math.sqrt(4 / (7.4 * R_f + 1.04))
        return term1 * term2 * term3 * term4

    def calculate_liquefaction_index(self, entry_group):
        try:
            q_c = float(entry_group[0].get())
            dw = float(entry_group[1].get())
            a_max = float(entry_group[2].get())
            ms = float(entry_group[3].get())
            R_f = float(entry_group[4].get())
            if R_f < 0.4:
                R_f = 0.4
            di = float(entry_group[5].get())
            zi = float(entry_group[6].get())

            beta = 0.2 * ms - 0.5
            q_ccr = self.calculate_q_ccr(beta, a_max, dw, zi, R_f)
            q_c_effective = min(q_c, q_ccr)
            Wi = 10 - 0.5 * zi
            return (1 - q_c_effective / q_ccr) * di * Wi

        except (ValueError, ZeroDivisionError):
            messagebox.showerror("Input Error", "Please ensure all inputs are valid numbers and the divisor is not zero!")
            return None

    def determine_liquefaction_level(self, I):
        if 0 < I <= 5:
            return "Slight Liquefaction"
        elif 5 < I <= 15:
            return "Moderate Liquefaction"
        elif I > 15:
            return "Severe Liquefaction"
        else:
            return "No Liquefaction"

    def get_level_color(self, level):
        if level == "No Liquefaction":
            return ACCENT_COLOR  # Green
        elif level == "Slight Liquefaction":
            return WARNING_COLOR  # Yellow
        elif level == "Moderate Liquefaction":
            return MEDIUM_COLOR  # Orange
        elif level == "Severe Liquefaction":
            return DANGER_COLOR  # Red
        return DARK_TEXT

    def calculate_all_points(self):
        self.tree.delete(*self.tree.get_children())

        while len(self.point_names) < len(self.all_points_entry_groups):
            self.point_names.append(tk.StringVar())

        for point_num, point_data in enumerate(self.all_points_entry_groups, 1):
            point_entry_groups = [entry_group for entry_group, _ in point_data]
            point_name_var = self.point_names[point_num - 1]
            point_name = point_name_var.get() or f"Point {point_num}"

            # Create point header row
            point_item = self.tree.insert("", "end", values=(
                f"{point_num}", point_name, "", "", "", ""
            ))
            self.tree.item(point_item, tags=("point_header",))

            # Calculate total liquefaction index for the point
            total_I = 0.0
            for entry_group in point_entry_groups:
                contribution = self.calculate_liquefaction_index(entry_group)
                if contribution is None:
                    return
                total_I += contribution

            # Determine liquefaction level for the point
            level = self.determine_liquefaction_level(total_I)

            # Insert all soil layer rows, only the last row shows the point liquefaction index and level
            for idx, entry_group in enumerate(point_entry_groups, 1):
                contribution = self.calculate_liquefaction_index(entry_group)
                if contribution is None:
                    return

                total_I_str = f"{total_I:.2f}" if idx == len(point_entry_groups) else ""
                level_str = level if idx == len(point_entry_groups) else ""

                item = self.tree.insert("", "end", values=(
                    "", "", f"Layer {idx}", f"{contribution:.2f}", total_I_str, level_str
                ))
                self.tree.item(item, tags=(level,))  # Add level tag to all rows

            # Add separator row
            self.tree.insert("", "end", values=("", "", "", "", "", ""))

        # Configure tag colors
        self.tree.tag_configure("point_header", background="#E6F7FF", font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("No Liquefaction", foreground=ACCENT_COLOR, font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("Slight Liquefaction", foreground=WARNING_COLOR, font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("Moderate Liquefaction", foreground=MEDIUM_COLOR, font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("Severe Liquefaction", foreground=DANGER_COLOR, font=("Segoe UI", 11, "bold"))

        # Show raw data table (if there is imported data)
        if self.raw_data_tree.get_children():
            self.raw_data_tree.pack(fill=tk.BOTH, expand=True, pady=5)

    def clear_results(self):
        """Clear the results table and raw data table, and hide the raw data table"""
        self.tree.delete(*self.tree.get_children())
        self.raw_data_tree.delete(*self.raw_data_tree.get_children())
        self.raw_data_tree.pack_forget()  # Hide raw data table

    def add_input_group(self, point_frame):
        """
        Add soil layer input group, including layer number, parameters, and delete button.
        :param point_frame: Point frame
        """
        point_idx = self.get_point_index(point_frame)
        layer_num = len(self.all_points_entry_groups[point_idx]) + 1

        layer_frame = ttk.Frame(point_frame)
        layer_frame.pack(fill=tk.X, pady=5)

        # Layer header frame (includes number)
        layer_header_frame = ttk.Frame(layer_frame, name="layer_header")
        layer_header_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(layer_header_frame, text=f"Layer {layer_num}:",
                  font=("Segoe UI", 12, "bold"), foreground=DARK_TEXT).pack(side=tk.LEFT, anchor="w")

        # Parameter input frame
        params_frame = ttk.Frame(layer_frame)
        params_frame.pack(fill=tk.X, padx=10, pady=5, side=tk.LEFT)

        param_labels = [
            "Measured Cone Tip Resistance (MPa):",
            "Groundwater Level Depth (m):",
            "Peak Ground Acceleration (g):",
            "Surface Wave Magnitude Ms:",
            "Friction Ratio:",
            "Liquefaction Layer Thickness (m):",
            "Middle Depth of Liquefaction Layer (m):"
        ]

        entry_group = []
        for text in param_labels:
            param_frame = ttk.Frame(params_frame)
            param_frame.pack(side=tk.LEFT, padx=5)
            ttk.Label(param_frame, text=text, anchor="w").pack(side=tk.LEFT, padx=2)
            entry = ttk.Entry(param_frame, width=10)
            entry.pack(side=tk.LEFT, padx=2)
            entry_group.append(entry)

        # Add delete layer button
        delete_btn_frame = ttk.Frame(layer_frame)
        delete_btn_frame.pack(side=tk.RIGHT, padx=5)
        delete_btn = ttk.Button(delete_btn_frame, text="Delete Layer",
                                command=lambda pf=point_frame, lidx=layer_num - 1:
                                self.delete_layer(pf, lidx),  # Current layer index is layer_num-1
                                style="DeleteButton.TButton")
        delete_btn.pack()

        # Save entry_group and layer_frame to data structure
        self.all_points_entry_groups[point_idx].append((entry_group, layer_frame))

        # Update scroll region
        self.points_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def add_new_point(self):
        """
        Add a new point and show the button to add a new soil layer.
        """
        point_num = len(self.all_points_entry_groups) + 1
        point_frame = ttk.LabelFrame(self.points_frame, text=f"Point {point_num}", padding=10)
        point_frame.pack(fill=tk.X, pady=10, ipady=5)

        # Point name and delete button
        name_and_btn_frame = ttk.Frame(point_frame)
        name_and_btn_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(name_and_btn_frame, text="Point Name:").pack(side=tk.LEFT, padx=5)
        name_var = tk.StringVar()
        ttk.Entry(name_and_btn_frame, textvariable=name_var, width=20).pack(side=tk.LEFT, padx=5)
        self.point_names.append(name_var)

        delete_btn = ttk.Button(name_and_btn_frame, text="Delete Point",
                                command=lambda pf=point_frame: self.delete_point(pf),
                                style="DeleteButton.TButton")
        delete_btn.pack(side=tk.RIGHT, padx=5)

        self.all_points_entry_groups.append([])

        # Add first soil layer
        self.add_input_group(point_frame)

        # Add new soil layer button
        add_layer_btn = ttk.Button(point_frame, text="Add New Layer",
                                   command=lambda pf=point_frame: self.add_input_group(pf),
                                   style="CenterAddButton.TButton")
        add_layer_btn.pack(side=tk.BOTTOM, pady=5, fill=tk.X, expand=True)

        self.points_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def delete_point(self, point_frame):
        """
        Delete an entire point.
        :param point_frame: Point frame to delete
        """
        point_idx = self.get_point_index(point_frame)
        if point_idx != -1:
            del self.all_points_entry_groups[point_idx]
            del self.point_names[point_idx]
            point_frame.destroy()
            self.update_point_numbers()
            self.points_frame.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def delete_layer(self, point_frame, layer_idx):
        """
        Delete a specified soil layer from a specified point.
        :param point_frame: Point frame
        :param layer_idx: Soil layer index
        """
        point_idx = self.get_point_index(point_frame)
        if point_idx == -1 or layer_idx < 0 or layer_idx >= len(self.all_points_entry_groups[point_idx]):
            return

        # Delete corresponding layer frame and data
        _, layer_frame = self.all_points_entry_groups[point_idx].pop(layer_idx)
        layer_frame.destroy()
        self.update_layer_numbers(point_frame)  # Update remaining layer numbers
        self.points_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_point_numbers(self):
        """Update the numbers of all points"""
        points = [child for child in self.points_frame.winfo_children()
                  if isinstance(child, ttk.LabelFrame) and "Point" in child['text']]
        for idx, point in enumerate(points, 1):
            point['text'] = f"Point {idx}"

    def update_layer_numbers(self, point_frame):
        """
        Update the numbers of all soil layers within a specified point.
        :param point_frame: Point frame
        """
        point_idx = self.get_point_index(point_frame)
        point_data = self.all_points_entry_groups[point_idx] if point_idx != -1 else []
        for new_idx, (_, layer_frame) in enumerate(point_data, 1):
            header_frames = [child for child in layer_frame.winfo_children()
                             if isinstance(child, ttk.Frame) and child.winfo_name() == "layer_header"]
            if header_frames:
                labels = [child for child in header_frames[0].winfo_children()
                          if isinstance(child, ttk.Label)]
                if labels:
                    labels[0].config(text=f"Layer {new_idx}:")

    def get_point_index(self, point_frame):
        """
        Get the index of the point frame in the UI.
        :param point_frame: Point frame
        :return: Index value, returns -1 if not found
        """
        points = [child for child in self.points_frame.winfo_children()
                  if isinstance(child, ttk.LabelFrame) and "Point" in child['text']]
        return points.index(point_frame) if point_frame in points else -1

    def import_excel(self):
        """Import Excel file"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.tree.delete(*self.tree.get_children())
                self.raw_data_tree.delete(*self.raw_data_tree.get_children())  # Clear raw data table

                if 'Point' not in df.columns:
                    messagebox.showerror("Format Error", "The Excel file must contain a 'Point' column!")
                    return

                # Check if necessary columns exist
                required_columns = ['q_c', 'dw', 'a_max', 'ms', 'R_f', 'di', 'zi']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    messagebox.showerror("Format Error", f"The Excel file is missing the following necessary columns: {', '.join(missing_columns)}")
                    return

                # Display raw data in the left table
                for _, row in df.iterrows():
                    point = row.get('Point', '')
                    q_c = row.get('q_c', '')
                    dw = row.get('dw', '')
                    a_max = row.get('a_max', '')
                    ms = row.get('ms', '')
                    R_f = row.get('R_f', '')
                    di = row.get('di', '')
                    zi = row.get('zi', '')

                    self.raw_data_tree.insert("", "end", values=(
                        point, q_c, dw, a_max, ms, R_f, di, zi
                    ))

                # Calculate and display results
                for point_idx, (point_name, group_data) in enumerate(df.groupby('Point'), 1):
                    total_I = 0.0
                    for _, row in group_data.iterrows():
                        q_c = row.get('q_c', 0.0)
                        dw = row.get('dw', 0.0)
                        a_max = row.get('a_max', 0.0)
                        ms = row.get('ms', 0.0)
                        R_f = row.get('R_f', 0.4)
                        if R_f < 0.4:
                            R_f = 0.4
                        di = row.get('di', 0.0)
                        zi = row.get('zi', 0.0)

                        beta = 0.2 * ms - 0.5
                        q_ccr = self.calculate_q_ccr(beta, a_max, dw, zi, R_f)
                        q_c_effective = min(q_c, q_ccr)
                        Wi = 10 - 0.5 * zi
                        contribution = (1 - q_c_effective / q_ccr) * di * Wi
                        total_I += contribution

                    level = self.determine_liquefaction_level(total_I)

                    # Create point header row
                    point_item = self.tree.insert("", "end", values=(
                        f"{point_idx}", point_name, "", "", "", ""
                    ))
                    self.tree.item(point_item, tags=("point_header",))

                    for layer_idx, (_, row) in enumerate(group_data.iterrows(), 1):
                        q_c = row.get('q_c', 0.0)
                        dw = row.get('dw', 0.0)
                        a_max = row.get('a_max', 0.0)
                        ms = row.get('ms', 0.0)
                        R_f = row.get('R_f', 0.4)
                        if R_f < 0.4:
                            R_f = 0.4
                        di = row.get('di', 0.0)
                        zi = row.get('zi', 0.0)

                        beta = 0.2 * ms - 0.5
                        q_ccr = self.calculate_q_ccr(beta, a_max, dw, zi, R_f)
                        q_c_effective = min(q_c, q_ccr)
                        Wi = 10 - 0.5 * zi
                        contribution = (1 - q_c_effective / q_ccr) * di * Wi

                        total_I_str = f"{total_I:.2f}" if layer_idx == len(group_data) else ""
                        level_str = level if layer_idx == len(group_data) else ""

                        item = self.tree.insert("", "end", values=(
                            "", "", f"Layer {layer_idx}", f"{contribution:.2f}", total_I_str, level_str
                        ))
                        self.tree.item(item, tags=(level,))

                    self.tree.insert("", "end", values=("", "", "", "", "", ""))

                # Configure tag colors
                self.tree.tag_configure("point_header", background="#E6F7FF", font=("Segoe UI", 11, "bold"))
                self.tree.tag_configure("No Liquefaction", foreground=ACCENT_COLOR, font=("Segoe UI", 11, "bold"))
                self.tree.tag_configure("Slight Liquefaction", foreground=WARNING_COLOR, font=("Segoe UI", 11, "bold"))
                self.tree.tag_configure("Moderate Liquefaction", foreground=MEDIUM_COLOR, font=("Segoe UI", 11, "bold"))
                self.tree.tag_configure("Severe Liquefaction", foreground=DANGER_COLOR, font=("Segoe UI", 11, "bold"))

                # Show raw data table
                self.raw_data_tree.pack(fill=tk.BOTH, expand=True, pady=5)

            except Exception as e:
                messagebox.showerror("Import Error", f"An error occurred while importing the file: {str(e)}")

    def show_help(self):
        """Display help dialog"""
        help_text = """
        Instructions for the Earthquake Sandy Soil Liquefaction Risk Index Rating Program Based on Static Cone Penetration Test

        I. Program Overview
        This program is used to calculate and evaluate the possibility of sandy soil liquefaction under earthquake action and its rating, based on the data of static cone penetration test (CPT).

        II. Description of Input Parameters
        1. Measured Cone Tip Resistance (MPa): The cone tip resistance value obtained from the static cone penetration test.
        2. Groundwater Level Depth (m): The depth of the groundwater level at the test point.
        3. Peak Ground Acceleration (g): The peak ground acceleration during the earthquake.
        4. Surface Wave Magnitude Ms: The surface wave magnitude of the earthquake.
        5. Friction Ratio: The ratio of the cone tip resistance to the sidewall frictional resistance.
        6. Liquefaction Layer Thickness (m): The thickness of the soil layer that may liquefy.
        7. Middle Depth of Liquefaction Layer (m): The middle depth of the liquefaction layer.

        III. Operation Steps
        1. Add Test Point: Click the "Add New Point" button.
        2. Add Soil Layer: Click the "Add New Layer" button under each test point.
        3. Input Parameters: Enter the parameters of each soil layer in the corresponding input boxes.
        4. Calculate Results: Click the "Calculate Liquefaction Level" button to obtain the calculation results.

        IV. Result Explanation
        1. Contribution Value: The contribution of each soil layer to the overall liquefaction possibility.
        2. Point Liquefaction Index: The comprehensive liquefaction possibility index of each soil layer.
        3. Liquefaction Level: The level of liquefaction (No Liquefaction, Slight Liquefaction, Moderate Liquefaction, Severe Liquefaction) divided according to the liquefaction index.

        V. Import and Export
        1. Import: Click the "Import Excel File" button to import the pre-prepared test data.
        2. Export: Click the "Export Results" button to export the calculation results as an Excel file.

        VI. Notes
        1. All input parameters should be valid numbers.
        2. The minimum value of the friction ratio is 0.4.
        3. Ensure that all necessary parameters are entered before calculation.
        """

        # Create help dialog
        help_window = tk.Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("600x500")
        help_window.resizable(True, True)
        # 创建文本框显示帮助内容
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)  # 设置为只读

        # 添加滚动条
        scrollbar = ttk.Scrollbar(text_widget, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)

    def export_results(self):
        """Export calculation results to Excel file"""
        if not self.tree.get_children():
            messagebox.showinfo("Export", "There are no calculation results to export!")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if file_path:
                # Create DataFrame
                data = []
                current_point = None
                current_point_name = None
                for item_id in self.tree.get_children():
                    values = self.tree.item(item_id, "values")
                    tags = self.tree.item(item_id, "tags")

                    if not values or all(v == "" for v in values):
                        continue

                    if tags and tags[0] == "point_header":
                        current_point = values[0]
                        current_point_name = values[1]
                        continue

                    if current_point and current_point_name:
                        layer = values[2]
                        contribution = values[3]
                        total_i = values[4]
                        level = values[5]

                        data.append({
                            "Point Number": current_point,
                            "Point Name": current_point_name,
                            "Soil Layer": layer,
                            "Contribution Value": contribution,
                            "Point Liquefaction Index": total_i,
                            "Liquefaction Level": level
                        })

                if data:
                    df = pd.DataFrame(data)
                    df.to_excel(file_path, index=False)
                    messagebox.showinfo("Export", f"Results exported successfully to {file_path}")
                else:
                    messagebox.showinfo("Export", "No valid data to export!")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting the results: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SoilLiquefactionCalculator(root)

    # Initialize variables
    app.all_points_entry_groups = []  # Stores all point soil layer entry groups
    app.point_names = []  # Stores all point name variables

    # Add the first point
    app.add_new_point()

    root.mainloop()