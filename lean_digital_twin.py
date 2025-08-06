import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re

# --- PLATFORM-SPECIFIC CHECK ---
if os.name != 'nt':
    # This check is important because window embedding is a Windows-only feature.
    messagebox.showerror("Unsupported OS",
                         "This version of the application uses Windows-specific features (pywin32) to embed Excel and can only run on Windows.")
    exit()

# --- STABLE DEPENDENCIES ---
import xlwings as xw
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from SPARQLWrapper import SPARQLWrapper, JSON
import win32gui
import win32con


class LeanDigitalTwin(tk.Tk):
    """
    An application for interacting with a GraphDB repository, visualizing data models,
    and linking data to an embedded Excel datasheet.
    [VERSION: XLWINGS - Embedded & Refined]
    """

    def __init__(self):
        super().__init__()
        self.initialization_ok = True  # Flag to track if setup is successful

        # First, check for Excel installation before proceeding
        try:
            # Use a more reliable Excel check
            app_check = xw.App(visible=False, add_book=False)
            app_check.quit()
        except Exception as e:
            self.withdraw()  # Hide the root window before showing the error
            messagebox.showerror("Excel Not Found",
                                 "Could not connect to Microsoft Excel. Please ensure it is installed and accessible.\n\n"
                                 f"Error: {e}")
            self.initialization_ok = False
            return  # Stop the initialization process

        self.title("LEAN Digital Twin (Embedded Excel Edition)")
        self.geometry("1400x850")

        self.properties = []
        self.current_excel_path = None
        self.tag_associations = {}
        self.graph = nx.DiGraph()
        self.xl_app = None
        self.current_xl_db = None
        self.active_node = None

        self.storage_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_files")
        os.makedirs(self.storage_dir, exist_ok=True)

        self._configure_styles()
        self._create_main_ui()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.after(100, self._update_file_lists)
        self.after(100, self._update_tag_lists)

    def on_closing(self):
        """Custom closing method to ensure Excel quits."""
        self._cleanup_xlwings()
        self.destroy()

    def _cleanup_xlwings(self):
        """Closes the Excel instance started by this application."""
        if self.xl_app:
            try:
                # Un-parent the window before quitting to avoid graphical glitches
                if hasattr(self.xl_app, 'hwnd') and self.xl_app.hwnd:
                    try:
                        win32gui.SetParent(self.xl_app.hwnd, 0)
                    except:
                        pass
                self.xl_app.quit()
                self.xl_app = None
                self._update_status("Excel instance closed.", 2000)
            except Exception as e:
                print(f"Could not quit Excel gracefully: {e}")

    def _configure_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('TLabel', font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        self.style.configure('TEntry', font=('Segoe UI', 10))
        self.style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=[10, 5])
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Info.TLabel', font=('Segoe UI', 9, 'italic'))
        self.style.configure('ActiveNode.TLabel', font=('Segoe UI', 11, 'bold'), foreground='blue')

    def _create_main_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both")

        self.main_notebook = ttk.Notebook(main_frame)
        self.main_notebook.pack(expand=True, fill="both")

        self.tab_graphical = ttk.Frame(self.main_notebook, padding=10)
        self.tab_graphdb = ttk.Frame(self.main_notebook, padding=10)
        self.tab_excel = ttk.Frame(self.main_notebook, padding=10)
        self.tab_functionalities = ttk.Frame(self.main_notebook, padding=10)

        self.main_notebook.add(self.tab_graphical, text="Graphical Model")
        self.main_notebook.add(self.tab_graphdb, text="GraphDB → Excel")
        self.main_notebook.add(self.tab_excel, text="Datasheet Editor")
        self.main_notebook.add(self.tab_functionalities, text="Functionalities")

        self._build_graphical_model_tab(self.tab_graphical)
        self._build_graphdb_tab(self.tab_graphdb)
        self._build_datasheet_editor_tab(self.tab_excel)
        self._build_functionalities_tab(self.tab_functionalities)

        self.status_bar = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor='w', padding=5)
        self.status_bar.pack(side="bottom", fill="x")

    def _update_status(self, message, clear_after_ms=None):
        self.status_bar.config(text=message)
        if clear_after_ms:
            self.after(clear_after_ms, lambda: self.status_bar.config(text="Ready"))

    def _build_graphical_model_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)

        controls_frame = ttk.Frame(parent)
        controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(controls_frame, text="Select Node(s) to Visualize:", style='Header.TLabel').pack(anchor="w")

        self.node_listbox_display = tk.Listbox(controls_frame, height=6, selectmode="extended", exportselection=False)
        self.node_listbox_display.pack(fill="x", expand=True, pady=5)
        self.node_listbox_display.bind("<<ListboxSelect>>", lambda event: self._update_data_model(event))

        button_bar = ttk.Frame(controls_frame)
        button_bar.pack(fill="x", pady=5)

        ttk.Button(button_bar, text="Refresh Model", command=self._update_data_model).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_bar, text="Clear Graph", command=self._clear_graph).pack(side=tk.LEFT)

        self.graph_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=1)
        self.graph_frame.grid(row=1, column=0, sticky="nsew")

    def _build_graphdb_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        conn_frame = ttk.LabelFrame(parent, text="1. Connection", padding=10)
        conn_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        conn_frame.columnconfigure(1, weight=1)

        ttk.Label(conn_frame, text="Repo URL:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_repo = ttk.Entry(conn_frame)
        self.entry_repo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(conn_frame, text="SPARQL Prefix:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_prefix = ttk.Entry(conn_frame)
        self.entry_prefix.insert(0, "PREFIX ex: <http://example.org/pumps#>")
        self.entry_prefix.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        node_select_frame = ttk.LabelFrame(parent,
                                           text="2. Master Node List (Double-click to set Active Node manually)",
                                           padding=10)
        node_select_frame.grid(row=1, column=0, sticky="nsew", pady=10)
        node_select_frame.rowconfigure(1, weight=1)
        node_select_frame.columnconfigure(0, weight=1)

        ttk.Button(node_select_frame, text="Fetch All Nodes from Repo", command=self._fetch_nodes).grid(row=0, column=0,
                                                                                                        sticky="w",
                                                                                                        padx=5, pady=5)
        self.node_listbox = tk.Listbox(node_select_frame, height=6, exportselection=False)
        self.node_listbox.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.node_listbox.bind("<Double-1>", lambda event: self._on_node_manual_select(event))

    def _build_datasheet_editor_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(1, weight=1)

        left_pane = ttk.Frame(parent, padding=5)
        left_pane.grid(row=0, column=0, rowspan=2, sticky="ns", pady=5)
        left_pane.rowconfigure(3, weight=1)

        ttk.Label(left_pane, text="Select a Tag to View:", style='Header.TLabel').grid(row=0, column=0, sticky="w",
                                                                                       pady=(0, 5))

        self.tag_selector_combobox = ttk.Combobox(left_pane, state="readonly")
        self.tag_selector_combobox.grid(row=1, column=0, sticky="ew")
        self.tag_selector_combobox.bind("<<ComboboxSelected>>",
                                        lambda event: self._on_tag_selected_in_editor_tab(event))

        ttk.Label(left_pane, text="Datasheets for Selected Tag:", style='Header.TLabel').grid(row=2, column=0,
                                                                                              sticky="w", pady=(10, 5))

        self.datasheet_listbox_for_tag = tk.Listbox(left_pane, height=15, exportselection=False)
        self.datasheet_listbox_for_tag.grid(row=3, column=0, sticky="nsew")
        self.datasheet_listbox_for_tag.bind("<Double-1>", lambda event: self._load_file_from_list(event))

        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.grid(row=1, column=1, sticky="nsew", padx=10, pady=5)

        table_container = ttk.Frame(main_pane)
        main_pane.add(table_container, weight=3)
        table_container.rowconfigure(1, weight=1)
        table_container.columnconfigure(0, weight=1)

        sheet_controls_frame = ttk.Frame(table_container)
        sheet_controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        ttk.Label(sheet_controls_frame, text="Activate Sheet:").pack(side=tk.LEFT, padx=(0, 5))

        self.sheet_selector_combobox = ttk.Combobox(sheet_controls_frame, state="readonly")
        self.sheet_selector_combobox.pack(side=tk.LEFT, fill="x", expand=True)
        self.sheet_selector_combobox.bind("<<ComboboxSelected>>", lambda event: self._on_sheet_selected(event))

        self.excel_frame = ttk.Frame(table_container, relief=tk.SUNKEN, borderwidth=1)
        self.excel_frame.grid(row=1, column=0, sticky="nsew")
        self.excel_frame.bind("<Configure>", self._resize_excel_window)

        info_frame = ttk.Frame(main_pane, padding=10)
        info_frame.columnconfigure(0, weight=1)
        main_pane.add(info_frame, weight=1)

        ttk.Label(info_frame, text="Tag Information", style='Header.TLabel').grid(row=0, column=0, sticky="w",
                                                                                  pady=(0, 5))
        self.tag_text_embed = tk.Text(info_frame, height=4, width=30, font=('Segoe UI', 10), relief=tk.SOLID,
                                      borderwidth=1, state='disabled')
        self.tag_text_embed.grid(row=1, column=0, sticky="ew")

        active_node_frame = ttk.LabelFrame(info_frame, text="Active Node for Mapping", padding=10)
        active_node_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.active_node_display_label = ttk.Label(active_node_frame, text="None Selected", style='ActiveNode.TLabel')
        self.active_node_display_label.pack(pady=2)

        ttk.Label(info_frame, text="Associated Node Properties (Double-click to preview):", style='Header.TLabel').grid(
            row=3, column=0, sticky="w", pady=(10, 5))
        self.properties_text_embed = tk.Text(info_frame, height=8, width=30, font=('Segoe UI', 10), relief=tk.SOLID,
                                             borderwidth=1, state='disabled')
        self.properties_text_embed.grid(row=4, column=0, sticky="nsew")
        self.properties_text_embed.bind("<Double-1>", self._on_property_double_click)

        paste_frame = ttk.LabelFrame(info_frame, text="Copy/Paste Live Data", padding=10)
        paste_frame.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        paste_frame.columnconfigure(1, weight=1)

        ttk.Label(paste_frame, text="1. Double-click property to copy.", style='Info.TLabel').grid(row=0, column=0,
                                                                                                   columnspan=2,
                                                                                                   sticky="w")
        ttk.Label(paste_frame, text="2. Type target cell (e.g., B5).", style='Info.TLabel').grid(row=1, column=0,
                                                                                                 columnspan=2,
                                                                                                 sticky="w")
        ttk.Label(paste_frame, text="3. Click Paste.", style='Info.TLabel').grid(row=2, column=0, columnspan=2,
                                                                                 sticky="w")

        self.live_value_button = ttk.Button(paste_frame, text="Copy Value: (none)")
        self.live_value_button.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 5))
        self.live_unit_button = ttk.Button(paste_frame, text="Copy Unit: (none)")
        self.live_unit_button.grid(row=4, column=0, columnspan=2, sticky="ew", pady=5)

        ttk.Label(paste_frame, text="Cell:").grid(row=5, column=0, sticky="w", padx=(0, 5), pady=5)
        self.paste_cell_entry = ttk.Entry(paste_frame, width=10)
        self.paste_cell_entry.grid(row=5, column=1, sticky="ew", pady=5)

        self.paste_button = ttk.Button(paste_frame, text="Paste to Active Excel Sheet",
                                       command=self._paste_from_clipboard)
        self.paste_button.grid(row=6, column=0, columnspan=2, sticky="ew", pady=5)

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=1, sticky="ew", pady=(10, 0), padx=10)

        ttk.Button(button_frame, text="Import Datasheet to Library...", command=self._import_datasheet_to_library).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Save As (Make a Copy)...", command=self._save_excel_as).pack(side=tk.LEFT,
                                                                                                    padx=5)

    def _build_functionalities_tab(self, parent):
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        sub_notebook = ttk.Notebook(parent)
        sub_notebook.grid(row=0, column=0, sticky="nsew")

        frame_tags = ttk.Frame(sub_notebook, padding=10)
        frame_excel_files = ttk.Frame(sub_notebook, padding=10)

        sub_notebook.add(frame_tags, text="Tag Management")
        sub_notebook.add(frame_excel_files, text="File Management")

        self._build_tags_subtab(frame_tags)
        self._build_excel_files_subtab(frame_excel_files)

    def _build_excel_files_subtab(self, parent):
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)

        ttk.Label(parent, text="Datasheet Library", style='Header.TLabel').grid(row=0, column=0, sticky="w",
                                                                                pady=(0, 5))
        ttk.Label(parent, text="This list contains all datasheets imported into the application's library.",
                  style='Info.TLabel').grid(row=1, column=0, sticky="w", pady=(0, 10))

        self.file_listbox_manage = tk.Listbox(parent, height=10)
        self.file_listbox_manage.grid(row=2, column=0, sticky="nsew", pady=5)

        ttk.Button(parent, text="Remove Selected File from Library", command=self._remove_file).grid(row=3, column=0,
                                                                                                     sticky="w",
                                                                                                     pady=10)

    def _build_tags_subtab(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)

        create_frame = ttk.LabelFrame(parent, text="Create or Update Tag", padding=10)
        create_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        create_frame.columnconfigure(1, weight=1)

        ttk.Label(create_frame, text="Tag Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_tag = ttk.Entry(create_frame)
        self.entry_tag.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(create_frame, text="Associate Node(s):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.node_combobox = ttk.Combobox(create_frame, state="readonly")
        self.node_combobox.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(create_frame, text="Associate Datasheet(s):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.datasheet_combobox = ttk.Combobox(create_frame, state="readonly")
        self.datasheet_combobox.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        ttk.Button(create_frame, text="Create/Update Tag", command=self._add_tag).grid(row=3, column=1, sticky="e",
                                                                                       padx=5, pady=10)

        view_frame = ttk.LabelFrame(parent, text="View Tag Associations", padding=10)
        view_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        view_frame.columnconfigure(1, weight=1)
        view_frame.rowconfigure(1, weight=1)

        ttk.Label(view_frame, text="Existing Tags (Double-click to view)").grid(row=0, column=0, columnspan=2,
                                                                                sticky="w", padx=5)
        self.tag_listbox = tk.Listbox(view_frame, height=5, exportselection=False)
        self.tag_listbox.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tag_listbox.bind("<Double-1>", lambda event: self._show_tag_connections(event))

        ttk.Label(view_frame, text="Associated Nodes").grid(row=2, column=0, sticky="w", padx=5, pady=(10, 0))
        self.nodes_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.nodes_display.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        view_frame.rowconfigure(3, weight=1)

        ttk.Label(view_frame, text="Associated Datasheets").grid(row=2, column=1, sticky="w", padx=5, pady=(10, 0))
        self.datasheets_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.datasheets_display.grid(row=3, column=1, sticky="nsew", padx=(5, 5), pady=5)
        self.datasheets_display.bind("<Double-1>", lambda event: self._load_datasheet_from_functionalities_tab(event))

    def _resize_excel_window(self, event=None):
        if self.xl_app and hasattr(self.xl_app, 'hwnd') and self.xl_app.hwnd:
            try:
                win32gui.MoveWindow(self.xl_app.hwnd, 0, 0, self.excel_frame.winfo_width(),
                                    self.excel_frame.winfo_height(), True)
            except (win32gui.error, AttributeError):
                # If we can't resize, the Excel window might be gone
                pass

    def _load_excel_file(self, file_path):
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                messagebox.showerror("File Not Found", f"The file could not be found at:\n{abs_path}")
                return

            # Clean up any existing Excel instance first
            if self.xl_app is not None:
                self._cleanup_xlwings()

            # Create Excel app with visible=True first
            self.xl_app = xw.App(visible=True, add_book=False)
            self.xl_app.display_alerts = False
            
            # Small delay to ensure Excel is fully loaded
            self.after(200, lambda: self._continue_excel_loading(abs_path))

        except Exception as e:
            messagebox.showerror("xlwings Load Error",
                                 f"Failed to initialize Excel.\n\nError: {e}")
            self._cleanup_xlwings()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

    def _continue_excel_loading(self, abs_path):
        try:
            # Open the workbook
            self.current_xl_db = self.xl_app.books.open(abs_path)
            
            # Get Excel window handle
            excel_hwnd = None
            max_attempts = 10
            attempt = 0
            
            # Try to get the Excel window handle with retries
            while excel_hwnd is None and attempt < max_attempts:
                try:
                    excel_hwnd = self.xl_app.hwnd
                    if excel_hwnd:
                        break
                except:
                    pass
                attempt += 1
                self.after(100)  # Wait 100ms before retry
            
            if not excel_hwnd:
                raise Exception("Could not get Excel window handle after multiple attempts")
            
            # Get the frame to embed into
            self.excel_frame.update_idletasks()  # Ensure frame is rendered
            frame_hwnd = self.excel_frame.winfo_id()
            
            # Embed Excel window into our frame
            win32gui.SetParent(excel_hwnd, frame_hwnd)

            # Remove window decorations (title bar, borders)
            style = win32gui.GetWindowLong(excel_hwnd, win32con.GWL_STYLE)
            style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME | win32con.WS_SYSMENU)
            win32gui.SetWindowLong(excel_hwnd, win32con.GWL_STYLE, style)

            # Activate the workbook and resize the window
            self.current_xl_db.activate()
            self.after(100, self._resize_excel_window)  # Delay resize to ensure embedding is complete

            # Update sheet selector
            sheet_names = [sheet.name for sheet in self.current_xl_db.sheets]
            self.sheet_selector_combobox['values'] = sheet_names
            if sheet_names:
                self.sheet_selector_combobox.set(sheet_names[0])
                self._display_sheet(sheet_names[0])

            self.current_excel_path = abs_path
            self._update_status(f"Embedded '{os.path.basename(abs_path)}'.", 4000)

        except Exception as e:
            messagebox.showerror("xlwings Embed Error",
                                 f"Failed to open or embed the Excel file.\n\nError: {e}")
            self._cleanup_xlwings()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

    def _on_sheet_selected(self, event=None):
        selected_sheet = self.sheet_selector_combobox.get()
        if selected_sheet:
            self._display_sheet(selected_sheet)

    def _display_sheet(self, sheet_name):
        if not self.current_xl_db:
            return
        try:
            self.current_xl_db.sheets[sheet_name].activate()
            self._update_status(f"Activated sheet '{sheet_name}'.", 3000)
        except Exception as e:
            messagebox.showerror("Sheet Activation Error", f"Could not activate sheet '{sheet_name}'.\n\nError: {e}")

    def _paste_from_clipboard(self):
        if not self.current_xl_db:
            messagebox.showwarning("No Datasheet", "Please open and embed a datasheet first.")
            return

        cell_address = self.paste_cell_entry.get().strip().upper()
        if not re.match(r"^[A-Z]+[1-9][0-9]*$", cell_address):
            messagebox.showwarning("Invalid Cell", "Please enter a valid cell address (e.g., A1, B5, C10).")
            return

        try:
            clipboard_content = self.clipboard_get()
            sheet = self.current_xl_db.sheets.active

            target_cell = sheet.range(cell_address)

            # Handle merged cells properly
            if hasattr(target_cell, 'merge_area') and target_cell.merge_area:
                write_cell = target_cell.merge_area.cells[0]
            else:
                write_cell = target_cell

            write_cell.value = clipboard_content
            self.current_xl_db.save()

            self._update_status(f"Pasted to {write_cell.address.replace('$', '')} in '{sheet.name}' and saved.", 4000)
            self.paste_cell_entry.delete(0, tk.END)

        except tk.TclError:
            messagebox.showwarning("Empty Clipboard", "The clipboard is empty or contains no text.")
        except Exception as e:
            messagebox.showerror("Paste Error", f"An error occurred while pasting to Excel: {e}")

    def _update_data_model(self, event=None):
        selected_indices = self.node_listbox_display.curselection()
        if not selected_indices: return

        nodes = [self.node_listbox_display.get(i) for i in selected_indices]
        repo_url, prefix = self.entry_repo.get().strip(), self.entry_prefix.get().strip()

        if not all([repo_url, prefix, nodes]):
            messagebox.showwarning("Missing Data", "Repository URL, Prefix, and a selected node are required.")
            return

        self._clear_graph()
        self._update_status("Fetching data model...")

        def get_local_name(uri):
            if not isinstance(uri, str): return uri
            return uri.split('#')[-1].split('/')[-1]

        def task():
            node_conditions = " || ".join(
                [f"sameTerm(?subject, ex:{node}) || sameTerm(?object, ex:{node})" for node in nodes])
            query = f"{prefix}\nSELECT ?subject ?predicate ?object WHERE {{ ?subject ?predicate ?object . FILTER({node_conditions}) }}"
            results = self._run_sparql_query(query)
            if results is None: return

            self.graph.clear()
            for res in results:
                subject = get_local_name(res.get("subject", {}).get("value", ""))
                predicate = get_local_name(res.get("predicate", {}).get("value", ""))
                obj = get_local_name(res.get("object", {}).get("value", ""))
                if all([subject, predicate, obj]):
                    self.graph.add_edge(subject, obj, label=predicate)

            self.after(0, self._draw_graph)

        threading.Thread(target=task, daemon=True).start()

    def _draw_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        if not self.graph.nodes():
            self._update_status("No data found for selected nodes.", 4000)
            return

        try:
            fig, ax = plt.subplots(figsize=(10, 8))
            pos = nx.spring_layout(self.graph, k=0.7, iterations=50)
            nx.draw(self.graph, pos, ax=ax, with_labels=True, node_color='#a0cbe2', node_size=2500, font_size=10,
                    font_weight='bold', width=1.5, edge_color='gray', arrows=True)
            edge_labels = nx.get_edge_attributes(self.graph, 'label')
            nx.draw_networkx_edge_labels(self.graph, pos, edge_labels=edge_labels, font_color='firebrick', font_size=9)
            fig.tight_layout()

            self.canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(expand=True, fill="both")
            self._update_status("Data model loaded.", 4000)
            plt.close(fig)
        except Exception as e:
            messagebox.showerror("Graphing Error", f"An error occurred while drawing the graph: {e}")

    def _clear_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        self.graph.clear()
        if hasattr(self, 'canvas'):
            del self.canvas
        self._update_status("Graph cleared.")

    def _run_sparql_query(self, query):
        repo_url = self.entry_repo.get().strip()
        if not repo_url:
            self.after(0, lambda: messagebox.showerror("Error", "Repository URL is not set."))
            return None
        try:
            sparql = SPARQLWrapper(repo_url)
            sparql.setQuery(query)
            sparql.setReturnFormat(JSON)
            return sparql.query().convert()["results"]["bindings"]
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("SPARQL Error", f"Failed to execute query:\n{e}"))
            self.after(0, lambda: self._update_status(f"SPARQL Error: {e}", 5000))
            return None

    def _fetch_nodes(self):
        prefix_str = self.entry_prefix.get().strip()
        if not prefix_str or '<' not in prefix_str or '>' not in prefix_str:
            messagebox.showerror("Invalid Prefix",
                                 "Please provide a valid SPARQL Prefix (e.g., PREFIX ex: <http://example.org#>)")
            return
        self._update_status("Fetching all nodes from repository...")

        def task():
            try:
                uri_base = prefix_str.split('<')[1].split('>')[0]
                query = f'SELECT DISTINCT ?resource WHERE {{ {{ ?resource ?p ?o . }} UNION {{ ?s ?p ?resource . }} FILTER(ISIRI(?resource) && STRSTARTS(STR(?resource), "{uri_base}")) }} ORDER BY ?resource'
                results = self._run_sparql_query(query)
                if results is None:
                    self.after(0, lambda: self._update_status("Failed to fetch nodes.", 4000))
                    return
                nodes = sorted(list(set(res["resource"]["value"].split('#')[-1].split('/')[-1] for res in results)))
                self.after(0, lambda: self._update_node_lists(nodes))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to parse prefix or fetch nodes: {e}"))
                self.after(0, lambda: self._update_status("Error fetching nodes.", 4000))

        threading.Thread(target=task, daemon=True).start()

    def _update_node_lists(self, nodes):
        self.node_listbox.delete(0, tk.END)
        self.node_listbox_display.delete(0, tk.END)
        for node in nodes:
            self.node_listbox.insert(tk.END, node)
            self.node_listbox_display.insert(tk.END, node)
        self.node_combobox['values'] = nodes
        self._update_status(f"{len(nodes)} nodes fetched.", 4000)

    def _update_file_lists(self):
        try:
            files = sorted([f for f in os.listdir(self.storage_dir) if f.endswith(('.xlsx', '.xls', '.xlsm'))])
            self.file_listbox_manage.delete(0, tk.END)
            for f in files:
                self.file_listbox_manage.insert(tk.END, f)
            self.datasheet_combobox['values'] = files
        except Exception as e:
            print(f"Error updating file list: {e}")

    def _update_tag_lists(self):
        try:
            tags = sorted(self.tag_associations.keys())
            self.tag_listbox.delete(0, tk.END)
            for tag in tags:
                self.tag_listbox.insert(tk.END, tag)
            self.tag_selector_combobox['values'] = tags
        except Exception as e:
            print(f"Error updating tag lists: {e}")

    def _import_datasheet_to_library(self):
        paths = filedialog.askopenfilenames(title="Select Datasheet(s) to Import",
                                            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if not paths: return
        imported_count = 0
        for path in paths:
            try:
                filename, internal_path = os.path.basename(path), os.path.join(self.storage_dir, os.path.basename(path))
                if not os.path.exists(internal_path):
                    shutil.copy(path, internal_path)
                    imported_count += 1
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import {filename}:\n{e}")
        if imported_count > 0:
            self._update_file_lists()
            messagebox.showinfo("Import Successful", f"{imported_count} new datasheet(s) added to the library.")
            self._update_status("Datasheet library updated.", 4000)

    def _save_excel_as(self):
        if not self.current_excel_path or not self.current_xl_db:
            messagebox.showwarning("No File", "A datasheet must be open in Excel to use 'Save As'.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx *.xlsm")],
                                                 title="Save Datasheet As")
        if not save_path: return
        try:
            abs_save_path = os.path.abspath(save_path)
            self.current_xl_db.save(abs_save_path)

            filename = os.path.basename(abs_save_path)
            internal_path = os.path.join(self.storage_dir, filename)

            if os.path.normpath(abs_save_path).lower() != os.path.normpath(internal_path).lower():
                shutil.copyfile(abs_save_path, internal_path)

            self._update_file_lists()
            messagebox.showinfo("Success",
                                f"File exported as '{filename}' and a copy was added/updated in the library.")
        except Exception as e:
            messagebox.showerror("Save As Error", f"Failed to save the file: {e}")

    def _load_file_from_list(self, event=None):
        if not event.widget.curselection(): return
        file_name = event.widget.get(event.widget.curselection())

        if self.current_xl_db and self.current_xl_db.name == file_name:
            try:
                self.current_xl_db.activate()
            except:
                # If activation fails, reload the file
                self._load_excel_file(os.path.join(self.storage_dir, file_name))
        else:
            self._load_excel_file(os.path.join(self.storage_dir, file_name))

    def _on_tag_selected_in_editor_tab(self, event=None):
        tag_name = self.tag_selector_combobox.get()
        if not tag_name: return

        nodes = self.tag_associations.get(tag_name, {}).get('nodes', [])
        if len(nodes) == 1:
            self._set_active_node(nodes[0])
        else:
            self._set_active_node(None)

        self.datasheet_listbox_for_tag.delete(0, tk.END)
        if tag_name in self.tag_associations:
            for datasheet in self.tag_associations[tag_name].get('datasheets', []):
                self.datasheet_listbox_for_tag.insert(tk.END, datasheet)
        self._display_tag_info_in_editor_view(tag_name)

    def _on_property_double_click(self, event=None):
        try:
            index = self.properties_text_embed.index(f"@{event.x},{event.y} linestart")
            line_end = self.properties_text_embed.index(f"{index} lineend")
            line_text = self.properties_text_embed.get(index, line_end)

            match = re.search(r"^\s*[-*]?\s*(\w+)", line_text)
            if not match: return
            prop = match.group(1)

            if not prop: return

            if not self.active_node:
                self.live_value_button.config(text="Copy Value: (No Active Node)", command=lambda: None)
                self.live_unit_button.config(text="Copy Unit: (No Active Node)", command=lambda: None)
                messagebox.showinfo("Info",
                                    "An active node must be set to see a live value preview.")
                return

            node = self.active_node
            self._update_status(f"Fetching preview for {prop}...")

            def task():
                prefix = self.entry_prefix.get().strip()
                query_bnode = f"{prefix} SELECT ?value ?unit WHERE {{ ex:{node} ex:{prop} ?bnode . ?bnode ex:hasValue ?value . OPTIONAL {{ ?bnode ex:hasUnit ?unit . }} }}"
                res_bnode = self._run_sparql_query(query_bnode)
                val, uni = (
                    res_bnode[0]["value"]["value"], res_bnode[0].get("unit", {}).get("value")) if res_bnode else (
                    None, None)

                if val is None:
                    query_direct = f"{prefix} SELECT ?value WHERE {{ ex:{node} ex:{prop} ?value . FILTER(isLiteral(?value)) }}"
                    res_direct = self._run_sparql_query(query_direct)
                    if res_direct:
                        val = res_direct[0]["value"]["value"]

                def update_ui():
                    self.live_value_button.config(text=f"Copy Value: {val or '(none)'}",
                                                  command=lambda v=val: self._copy_to_clipboard(v))
                    self.live_unit_button.config(text=f"Copy Unit: {uni or '(none)'}",
                                                 command=lambda u=uni: self._copy_to_clipboard(u))
                    self._update_status("Preview loaded.", 4000)

                self.after(0, update_ui)

            threading.Thread(target=task, daemon=True).start()

        except (tk.TclError, IndexError):
            pass

    def _display_tag_info_in_editor_view(self, tag_name):
        self.tag_text_embed.config(state='normal')
        self.properties_text_embed.config(state='normal')
        self.tag_text_embed.delete(1.0, tk.END)
        self.properties_text_embed.delete(1.0, tk.END)

        if tag_name not in self.tag_associations:
            self.tag_text_embed.insert(tk.END, "Tag not found.")
            self.tag_text_embed.config(state='disabled')
            self.properties_text_embed.config(state='disabled')
            return

        self.tag_text_embed.insert(tk.END, f"Tag: {tag_name}")
        nodes = self.tag_associations[tag_name].get('nodes', [])

        if not nodes:
            self.properties_text_embed.insert(tk.END, "Tag has no associated nodes.")
        else:
            node_list_str = '\n- '.join(nodes)
            self.properties_text_embed.insert(tk.END,
                                              f"Associated Nodes:\n- {node_list_str}\n\nProperties (from all nodes):\n")

        self.tag_text_embed.config(state='disabled')
        self.properties_text_embed.config(state='disabled')

        if not nodes: return

        self._update_status(f"Fetching properties for tag '{tag_name}'...")

        def task():
            all_properties = set()
            prefix = self.entry_prefix.get().strip()
            for node in nodes:
                query = f"""{prefix}
                SELECT DISTINCT ?p WHERE {{ 
                    ex:{node} ?p ?o .
                    FILTER (isLiteral(?o) || isBlank(?o))
                }}"""
                results = self._run_sparql_query(query)
                if results:
                    all_properties.update(p.split('#')[-1] for p in (res['p']['value'] for res in results) if
                                          p.split('#')[-1] not in ["hasValue", "hasUnit", "a", "type"])
            prop_text = "\n".join(f"- {p}" for p in sorted(list(all_properties))) or "(No direct properties found)"

            def update_text():
                self.properties_text_embed.config(state='normal')
                self.properties_text_embed.insert(tk.END, prop_text)
                self.properties_text_embed.config(state='disabled')
                self._update_status(f"Info loaded for tag '{tag_name}'.", 4000)

            self.after(0, update_text)

        threading.Thread(target=task, daemon=True).start()

    def _remove_file(self):
        if not self.file_listbox_manage.curselection():
            messagebox.showwarning("No Selection", "Please select a file to remove from the library.")
            return
        file_name = self.file_listbox_manage.get(self.file_listbox_manage.curselection())

        if self.current_xl_db and self.current_xl_db.name == file_name:
            self.current_xl_db.close()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

        if messagebox.askyesno("Confirm Removal",
                               f"Are you sure you want to permanently delete '{file_name}' from the library? This will also un-tag it from any associations."):
            try:
                os.remove(os.path.join(self.storage_dir, file_name))
                for tag in self.tag_associations:
                    if file_name in self.tag_associations[tag]['datasheets']:
                        self.tag_associations[tag]['datasheets'].remove(file_name)
                self._update_file_lists()
                self._on_tag_selected_in_editor_tab()
                messagebox.showinfo("Success", f"'{file_name}' has been removed from the library.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove file: {e}")

    def _add_tag(self):
        tag, node, datasheet = self.entry_tag.get().strip(), self.node_combobox.get().strip(), self.datasheet_combobox.get().strip()
        if not tag:
            messagebox.showwarning("Input Required", "Tag name cannot be empty.")
            return
        if tag not in self.tag_associations:
            self.tag_associations[tag] = {'nodes': [], 'datasheets': []}
        if node and node not in self.tag_associations[tag]['nodes']:
            self.tag_associations[tag]['nodes'].append(node)
        if datasheet and datasheet not in self.tag_associations[tag]['datasheets']:
            self.tag_associations[tag]['datasheets'].append(datasheet)
        messagebox.showinfo("Success", f"Tag '{tag}' created/updated.")
        self._update_tag_lists()
        self._show_tag_connections(tag_name=tag)

    def _show_tag_connections(self, event=None, tag_name=None):
        tag = tag_name or (
            self.tag_listbox.get(self.tag_listbox.curselection()) if self.tag_listbox.curselection() else None)
        if not tag: return
        self.nodes_display.delete(0, tk.END)
        self.datasheets_display.delete(0, tk.END)
        if tag in self.tag_associations:
            for node in self.tag_associations[tag].get('nodes', []):
                self.nodes_display.insert(tk.END, node)
            for datasheet in self.tag_associations[tag].get('datasheets', []):
                self.datasheets_display.insert(tk.END, datasheet)

    def _load_datasheet_from_functionalities_tab(self, event=None):
        if not event.widget.curselection(): return
        file_name = event.widget.get(event.widget.curselection())
        self.main_notebook.select(2)
        self._load_file_from_list(event)

    def _on_node_manual_select(self, event=None):
        if not self.node_listbox.curselection(): return
        node_name = self.node_listbox.get(self.node_listbox.curselection())
        self._set_active_node(node_name)

    def _set_active_node(self, node_name):
        self.active_node = node_name
        if node_name:
            self.active_node_display_label.config(text=node_name)
            try:
                idx = self.node_listbox.get(0, "end").index(node_name)
                self.node_listbox.selection_clear(0, tk.END)
                self.node_listbox.selection_set(idx)
                self.node_listbox.see(idx)
            except ValueError:
                pass
        else:
            self.active_node_display_label.config(text="None (Select a tag with one node)")
            self.node_listbox.selection_clear(0, tk.END)

    def _copy_to_clipboard(self, value_to_copy):
        if value_to_copy is None:
            messagebox.showinfo("No Value", "There is no value to copy.")
            return

        self.clipboard_clear()
        self.clipboard_append(str(value_to_copy))
        self._update_status(f"'{value_to_copy}' copied to clipboard.", 3000)


if __name__ == "__main__":
    # Create the application instance
    app = LeanDigitalTwin()

    # Only run the main event loop if the initialization was successful.
    # If it failed, the error message has already been shown.
    if app.initialization_ok:
        app.mainloop()