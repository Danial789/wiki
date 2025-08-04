import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- STABLE DEPENDENCIES ---
import pandas as pd
from pandastable import Table

import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from SPARQLWrapper import SPARQLWrapper, JSON

class LeanDigitalTwin(tk.Tk):
    """
    An application for interacting with a GraphDB repository, visualizing data models,
    and linking data to Excel datasheets, creating a "Lean Digital Twin."
    [VERSION: PANDASTABLE with Sheet Selection]
    """

    def __init__(self):
        super().__init__()
        self.title("LEAN Digital Twin")
        self.geometry("1400x800")

        self.properties = []
        self.current_excel_path = None
        self.tag_associations = {}
        self.graph = nx.DiGraph()
        self.pt_table = None
        self.current_workbook = None

        self.storage_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_files")
        os.makedirs(self.storage_dir, exist_ok=True)

        self._configure_styles()
        self._create_main_ui()
        
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        
        self.after(100, self._update_file_lists)
        self.after(100, self._update_tag_lists)

    def _configure_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('TLabel', font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        self.style.configure('TEntry', font=('Segoe UI', 10))
        self.style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=[10, 5])
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Info.TLabel', font=('Segoe UI', 9, 'italic'))

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
        parent.columnconfigure(1, weight=1)
        
        conn_frame = ttk.LabelFrame(parent, text="1. Connection & Node Selection", padding=10)
        conn_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        conn_frame.columnconfigure(1, weight=1)
        
        ttk.Label(conn_frame, text="Repo URL:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_repo = ttk.Entry(conn_frame)
        self.entry_repo.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        
        ttk.Label(conn_frame, text="SPARQL Prefix:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_prefix = ttk.Entry(conn_frame)
        self.entry_prefix.insert(0, "PREFIX ex: <http://example.org/pumps#>")
        self.entry_prefix.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        
        ttk.Button(conn_frame, text="Fetch All Nodes from Repo", command=self._fetch_nodes).grid(row=2, column=1, sticky="w", padx=5, pady=10)
        
        self.node_listbox = tk.Listbox(conn_frame, height=6, exportselection=False)
        self.node_listbox.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.node_listbox.bind("<Double-1>", lambda event: self._select_and_fetch_properties(event))
        conn_frame.rowconfigure(3, weight=1)
        
        prop_frame = ttk.LabelFrame(parent, text="2. Property Mapping", padding=10)
        prop_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=10)
        prop_frame.columnconfigure(1, weight=1)
        
        ttk.Label(prop_frame, text="Properties for selected node (Double-click to add):").grid(row=0, column=0, columnspan=2, sticky="w", padx=5)
        self.prop_listbox_selected = tk.Listbox(prop_frame, height=5, exportselection=False)
        self.prop_listbox_selected.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.prop_listbox_selected.bind("<Double-1>", lambda event: self._select_property(event))
        
        ttk.Separator(prop_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, sticky='ew', pady=10)
        
        ttk.Label(prop_frame, text="Property:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.entry_prop = ttk.Entry(prop_frame)
        self.entry_prop.grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Label(prop_frame, text="Value Cell (e.g., B2):").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.entry_val = ttk.Entry(prop_frame)
        self.entry_val.grid(row=4, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Label(prop_frame, text="Unit Cell (optional):").grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.entry_unit = ttk.Entry(prop_frame)
        self.entry_unit.grid(row=5, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Button(prop_frame, text="Add Property Mapping", command=self._add_property).grid(row=6, column=1, sticky="w", padx=5, pady=10)
        
        self.prop_listbox = tk.Listbox(prop_frame, height=5)
        self.prop_listbox.grid(row=7, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        out_frame = ttk.LabelFrame(parent, text="3. Excel Output", padding=10)
        out_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        out_frame.columnconfigure(1, weight=1)
        
        ttk.Label(out_frame, text="Target Datasheet:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_xlsx = ttk.Entry(out_frame, state='readonly')
        self.entry_xlsx.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Button(out_frame, text="Select from Library…", command=self._browse_xlsx_from_library).grid(row=0, column=2, padx=5)
        
        ttk.Label(out_frame, text="Sheet Name:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_sheet = ttk.Entry(out_frame)
        self.entry_sheet.insert(0, "Sheet1")
        self.entry_sheet.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Button(out_frame, text="Write to Excel", command=self._write_to_excel).grid(row=2, column=1, sticky="w", padx=5, pady=10)

    def _build_datasheet_editor_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(1, weight=1)
        
        left_pane = ttk.Frame(parent, padding=5)
        left_pane.grid(row=0, column=0, rowspan=2, sticky="ns", pady=5)
        left_pane.rowconfigure(3, weight=1)
        
        ttk.Label(left_pane, text="Select a Tag to View:", style='Header.TLabel').grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.tag_selector_combobox = ttk.Combobox(left_pane, state="readonly")
        self.tag_selector_combobox.grid(row=1, column=0, sticky="ew")
        self.tag_selector_combobox.bind("<<ComboboxSelected>>", lambda event: self._on_tag_selected_in_editor_tab(event))
        
        ttk.Label(left_pane, text="Datasheets for Selected Tag:", style='Header.TLabel').grid(row=2, column=0, sticky="w", pady=(10, 5))
        
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
        
        ttk.Label(sheet_controls_frame, text="Select Sheet:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.sheet_selector_combobox = ttk.Combobox(sheet_controls_frame, state="readonly")
        self.sheet_selector_combobox.pack(side=tk.LEFT, fill="x", expand=True)
        self.sheet_selector_combobox.bind("<<ComboboxSelected>>", lambda event: self._on_sheet_selected(event))
        
        self.excel_frame = ttk.Frame(table_container, relief=tk.SUNKEN, borderwidth=1)
        self.excel_frame.grid(row=1, column=0, sticky="nsew")
        
        info_frame = ttk.Frame(main_pane, padding=10)
        info_frame.columnconfigure(0, weight=1)
        main_pane.add(info_frame, weight=1)
        
        ttk.Label(info_frame, text="Tag Information", style='Header.TLabel').grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.tag_text_embed = tk.Text(info_frame, height=5, width=30, font=('Segoe UI', 10), relief=tk.SOLID, borderwidth=1)
        self.tag_text_embed.grid(row=1, column=0, sticky="ew")
        
        ttk.Label(info_frame, text="Associated Node Properties:", style='Header.TLabel').grid(row=2, column=0, sticky="w", pady=(10, 5))
        
        self.properties_text_embed = tk.Text(info_frame, height=15, width=30, font=('Segoe UI', 10), relief=tk.SOLID, borderwidth=1)
        self.properties_text_embed.grid(row=3, column=0, sticky="nsew")
        info_frame.rowconfigure(3, weight=1)
        
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=1, sticky="ew", pady=(10, 0), padx=10)
        
        ttk.Button(button_frame, text="Import Datasheet to Library...", command=self._import_datasheet_to_library).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Save Current Datasheet", command=self._save_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save As...", command=self._save_excel_as).pack(side=tk.LEFT, padx=5)

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
        
        ttk.Label(parent, text="Datasheet Library", style='Header.TLabel').grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Label(parent, text="This list contains all datasheets imported into the application's library.", style='Info.TLabel').grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        self.file_listbox_manage = tk.Listbox(parent, height=10)
        self.file_listbox_manage.grid(row=2, column=0, sticky="nsew", pady=5)
        
        ttk.Button(parent, text="Remove Selected File from Library", command=self._remove_file).grid(row=3, column=0, sticky="w", pady=10)

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
        
        ttk.Button(create_frame, text="Create/Update Tag", command=self._add_tag).grid(row=3, column=1, sticky="e", padx=5, pady=10)
        
        view_frame = ttk.LabelFrame(parent, text="View Tag Associations", padding=10)
        view_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        view_frame.columnconfigure(1, weight=1)
        view_frame.rowconfigure(1, weight=1)
        
        ttk.Label(view_frame, text="Existing Tags (Double-click to view)").grid(row=0, column=0, columnspan=2, sticky="w", padx=5)
        self.tag_listbox = tk.Listbox(view_frame, height=5, exportselection=False)
        self.tag_listbox.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tag_listbox.bind("<Double-1>", lambda event: self._show_tag_connections(event))
        
        ttk.Label(view_frame, text="Associated Nodes").grid(row=2, column=0, sticky="w", padx=5, pady=(10,0))
        self.nodes_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.nodes_display.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        view_frame.rowconfigure(3, weight=1)
        
        ttk.Label(view_frame, text="Associated Datasheets").grid(row=2, column=1, sticky="w", padx=5, pady=(10,0))
        self.datasheets_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.datasheets_display.grid(row=3, column=1, sticky="nsew", padx=(5,5), pady=5)
        self.datasheets_display.bind("<Double-1>", lambda event: self._load_datasheet_from_functionalities_tab(event))

    # --- MISSING METHODS - Added implementations ---
    
    def _update_data_model(self, event=None):
        """Update the graphical data model visualization"""
        try:
            # Clear existing visualization
            for widget in self.graph_frame.winfo_children():
                widget.destroy()
            
            # Get selected nodes
            selected_indices = self.node_listbox_display.curselection()
            if not selected_indices:
                # If no nodes selected, show a message
                label = ttk.Label(self.graph_frame, text="Select one or more nodes to visualize the data model.", 
                                font=('Segoe UI', 12), foreground='gray')
                label.pack(expand=True)
                return
            
            selected_nodes = [self.node_listbox_display.get(i) for i in selected_indices]
            
            # Create a new graph
            self.graph.clear()
            
            # Add nodes and query for relationships
            prefix = self.entry_prefix.get().strip()
            if prefix:
                self._fetch_graph_relationships(selected_nodes, prefix)
            
            # Create matplotlib figure
            fig, ax = plt.subplots(figsize=(8, 6))
            fig.patch.set_facecolor('white')
            
            if self.graph.nodes():
                pos = nx.spring_layout(self.graph, k=2, iterations=50)
                
                # Draw nodes
                nx.draw_networkx_nodes(self.graph, pos, ax=ax, 
                                     node_color='lightblue', 
                                     node_size=1000, 
                                     alpha=0.8)
                
                # Draw edges
                nx.draw_networkx_edges(self.graph, pos, ax=ax, 
                                     edge_color='gray', 
                                     arrows=True, 
                                     arrowsize=20,
                                     alpha=0.6)
                
                # Draw labels
                nx.draw_networkx_labels(self.graph, pos, ax=ax, 
                                      font_size=8, 
                                      font_weight='bold')
                
                ax.set_title(f"Data Model for: {', '.join(selected_nodes)}", 
                           fontsize=12, fontweight='bold')
            else:
                ax.text(0.5, 0.5, 'No relationships found for selected nodes', 
                       ha='center', va='center', transform=ax.transAxes, 
                       fontsize=12, color='gray')
                ax.set_title("Data Model Visualization", fontsize=12, fontweight='bold')
            
            ax.axis('off')
            
            # Embed in tkinter
            canvas = FigureCanvasTkAgg(fig, self.graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self._update_status(f"Model updated for {len(selected_nodes)} node(s).", 4000)
            
        except Exception as e:
            messagebox.showerror("Visualization Error", f"Failed to update data model: {e}")
    
    def _fetch_graph_relationships(self, nodes, prefix):
        """Fetch relationships between nodes for visualization"""
        try:
            for node in nodes:
                self.graph.add_node(node)
                
                # Query for outgoing relationships
                query = f"{prefix}\nSELECT DISTINCT ?predicate ?object WHERE {{ ex:{node} ?predicate ?object . FILTER(!isBlank(?object) && ISIRI(?object)) }}"
                results = self._run_sparql_query(query)
                
                if results:
                    for result in results:
                        obj_uri = result["object"]["value"]
                        # Extract the local name from URI
                        obj_name = obj_uri.split('#')[-1].split('/')[-1]
                        pred_name = result["predicate"]["value"].split('#')[-1].split('/')[-1]
                        
                        if obj_name and obj_name != node:
                            self.graph.add_edge(node, obj_name, label=pred_name)
                
                # Query for incoming relationships
                query = f"{prefix}\nSELECT DISTINCT ?subject ?predicate WHERE {{ ?subject ?predicate ex:{node} . FILTER(!isBlank(?subject) && ISIRI(?subject)) }}"
                results = self._run_sparql_query(query)
                
                if results:
                    for result in results:
                        subj_uri = result["subject"]["value"]
                        subj_name = subj_uri.split('#')[-1].split('/')[-1]
                        pred_name = result["predicate"]["value"].split('#')[-1].split('/')[-1]
                        
                        if subj_name and subj_name != node:
                            self.graph.add_edge(subj_name, node, label=pred_name)
                            
        except Exception as e:
            print(f"Error fetching graph relationships: {e}")
    
    def _clear_graph(self):
        """Clear the graph visualization"""
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        
        self.graph.clear()
        
        label = ttk.Label(self.graph_frame, text="Graph cleared. Select nodes to visualize.", 
                         font=('Segoe UI', 12), foreground='gray')
        label.pack(expand=True)
        
        self._update_status("Graph cleared.", 3000)
    
    # --- EXISTING METHODS ---
    
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
            messagebox.showerror("Invalid Prefix", "Please provide a valid SPARQL Prefix (e.g., PREFIX ex: <http://example.org#>)")
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
        
    def _select_and_fetch_properties(self, event=None):
        if not self.node_listbox.curselection(): return
        selected_node = self.node_listbox.get(self.node_listbox.curselection())
        self._update_status(f"Fetching properties for {selected_node}...")
        def task():
            query = f"{self.entry_prefix.get().strip()}\nSELECT DISTINCT ?predicate WHERE {{ ex:{selected_node} ?predicate ?object . }}"
            results = self._run_sparql_query(query)
            if results is None: return
            properties = sorted([p.split('#')[-1] for p in (res["predicate"]["value"] for res in results) if p.split('#')[-1] not in ["hasValue", "hasUnit", "a", "type"]])
            self.after(0, self.prop_listbox_selected.delete(0, tk.END))
            self.after(0, lambda: [self.prop_listbox_selected.insert(tk.END, p) for p in properties])
            self.after(0, lambda: self._update_status(f"Properties for {selected_node} loaded.", 4000))
        threading.Thread(target=task, daemon=True).start()

    def _select_property(self, event=None):
        if not self.prop_listbox_selected.curselection(): return
        self.entry_prop.delete(0, tk.END)
        self.entry_prop.insert(0, self.prop_listbox_selected.get(self.prop_listbox_selected.curselection()))
        self.entry_val.focus_set()

    def _add_property(self):
        prop, val, unit = self.entry_prop.get().strip(), self.entry_val.get().strip().upper(), self.entry_unit.get().strip().upper()
        if not prop or not val:
            messagebox.showwarning("Input Required", "Property and Value Cell are required.")
            return
        self.properties.append((prop, val, unit))
        self.prop_listbox.insert(tk.END, f"{prop} → Val: {val}, Unit: {unit or 'N/A'}")
        for entry in [self.entry_prop, self.entry_val, self.entry_unit]:
            entry.delete(0, tk.END)
        self.entry_prop.focus_set()

    def _browse_xlsx_from_library(self):
        files = [f for f in os.listdir(self.storage_dir) if f.endswith(('.xlsx', '.xls'))]
        if not files:
            messagebox.showinfo("Library Empty", "No datasheets have been imported. Import one from the 'Datasheet Editor' tab first.")
            return
        top = tk.Toplevel(self)
        top.title("Select Datasheet from Library")
        top.geometry("300x250")
        listbox = tk.Listbox(top)
        listbox.pack(expand=True, fill="both", padx=10, pady=5)
        for f in files:
            listbox.insert(tk.END, f)
        def on_select():
            if not listbox.curselection(): return
            filename = listbox.get(listbox.curselection())
            self.current_excel_path = os.path.join(self.storage_dir, filename)
            self.entry_xlsx.config(state='normal')
            self.entry_xlsx.delete(0, tk.END)
            self.entry_xlsx.insert(0, filename)
            self.entry_xlsx.config(state='readonly')
            top.destroy()
        ttk.Button(top, text="Select", command=on_select).pack(pady=5)
        
    def _cell_to_indices(self, cell_str):
        import re
        match = re.match(r"([A-Z]+)([0-9]+)", cell_str.upper())
        if not match:
            raise ValueError(f"Invalid cell format: '{cell_str}'")
        col_str, row_str = match.groups()
        row = int(row_str) - 1
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - ord('A') + 1)
        return row, col - 1

    def _write_to_excel(self):
        if not self.current_excel_path or not self.properties or not self.node_listbox.curselection():
            messagebox.showerror("Missing Data", "Ensure a node is selected, a target datasheet is chosen, and properties are mapped.")
            return
        self._update_status("Writing data to Excel...")
        node = self.node_listbox.get(self.node_listbox.curselection())
        def task():
            try:
                abs_path = os.path.abspath(self.current_excel_path)
                sheet_name = self.entry_sheet.get().strip() or "Sheet1"
                df = pd.read_excel(abs_path, sheet_name=sheet_name, header=None)
                prefix = self.entry_prefix.get().strip()
                for prop, val_cell, unit_cell in self.properties:
                    query_bnode = f"{prefix} SELECT ?value ?unit WHERE {{ ex:{node} ex:{prop} [ ex:hasValue ?value ; ex:hasUnit ?unit ] . }}"
                    res_bnode = self._run_sparql_query(query_bnode)
                    val, uni = (res_bnode[0]["value"]["value"], res_bnode[0]["unit"]["value"]) if res_bnode else (None, None)
                    if val is None:
                        query_direct = f"{prefix} SELECT ?value WHERE {{ ex:{node} ex:{prop} ?value . FILTER(!isBlank(?value)) }}"
                        res_direct = self._run_sparql_query(query_direct)
                        if res_direct:
                            val = res_direct[0]["value"]["value"]
                    if val is not None:
                        row, col = self._cell_to_indices(val_cell)
                        df.iat[row, col] = val
                    if uni is not None and unit_cell: 
                        row, col = self._cell_to_indices(unit_cell)
                        df.iat[row, col] = uni
                df.to_excel(abs_path, sheet_name=sheet_name, index=False, header=False)
                self.after(0, lambda: self._refresh_open_datasheet(abs_path))
                self.after(0, lambda: messagebox.showinfo("Success", "Data written to datasheet and saved."))
                self.after(0, lambda: self._update_status("Successfully wrote data to Excel.", 4000))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Write Error", f"An error occurred: {e}"))
                self.after(0, lambda: self._update_status("Error writing to Excel.", 4000))
        threading.Thread(target=task, daemon=True).start()

    def _update_file_lists(self):
        try:
            files = sorted([f for f in os.listdir(self.storage_dir) if f.endswith(('.xlsx', '.xls'))])
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
        paths = filedialog.askopenfilenames(title="Select Datasheet(s) to Import", filetypes=[("Excel files", "*.xlsx *.xls")])
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

    def _load_excel_file(self, file_path):
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                messagebox.showerror("File Not Found", f"The file could not be found at:\n{abs_path}")
                return
            
            self.current_workbook = pd.ExcelFile(abs_path)
            sheet_names = self.current_workbook.sheet_names
            
            self.sheet_selector_combobox['values'] = sheet_names
            if sheet_names:
                self.sheet_selector_combobox.set(sheet_names[0])
                self._display_sheet(sheet_names[0])

            self.current_excel_path = abs_path
            self._update_status(f"Opened '{os.path.basename(file_path)}'.", 4000)
        except Exception as e:
            messagebox.showerror("File Load Error", f"Failed to load the Excel file.\n\nError: {e}")
            self.current_workbook = None

    def _on_sheet_selected(self, event=None):
        selected_sheet = self.sheet_selector_combobox.get()
        if selected_sheet:
            self._display_sheet(selected_sheet)

    def _display_sheet(self, sheet_name):
        if not self.current_workbook: return
        try:
            df = pd.read_excel(self.current_workbook, sheet_name=sheet_name, engine='openpyxl')
            for widget in self.excel_frame.winfo_children():
                widget.destroy()
            self.pt_table = Table(self.excel_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
            self.pt_table.show()
        except Exception as e:
            messagebox.showerror("Sheet Load Error", f"Could not load sheet '{sheet_name}'.\n\nError: {e}")

    def _refresh_open_datasheet(self, file_path):
        if self.current_excel_path and os.path.normpath(file_path).lower() == os.path.normpath(self.current_excel_path).lower():
            self._load_excel_file(self.current_excel_path)

    def _save_excel(self):
        if not self.current_excel_path or not self.pt_table:
            messagebox.showwarning("No File", "No datasheet is currently open to save.")
            return
        try:
            df = self.pt_table.model.df
            df.to_excel(self.current_excel_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", "Datasheet saved successfully.")
            self._update_status("File saved.", 4000)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save file: {e}")
            
    def _save_excel_as(self):
        if not self.current_excel_path or not self.pt_table:
            messagebox.showwarning("No File", "A datasheet must be open to use 'Save As'.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save Datasheet As")
        if not save_path: return
        try:
            abs_save_path = os.path.abspath(save_path)
            df = self.pt_table.model.df
            df.to_excel(abs_save_path, index=False, engine='openpyxl')
            filename = os.path.basename(abs_save_path)
            internal_path = os.path.join(self.storage_dir, filename)
            if os.path.normpath(abs_save_path).lower() != os.path.normpath(internal_path).lower():
                shutil.copyfile(abs_save_path, internal_path)
            self.current_excel_path = internal_path
            self._update_file_lists()
            messagebox.showinfo("Success", f"File exported as '{filename}' and a new copy was added to the application library.")
        except Exception as e:
            messagebox.showerror("Save As Error", f"Failed to save the file: {e}")
            
    def _load_file_from_list(self, event=None):
        if not event.widget.curselection(): return
        file_name = event.widget.get(event.widget.curselection())
        self._load_excel_file(os.path.join(self.storage_dir, file_name))

    def _on_tag_selected_in_editor_tab(self, event=None):
        tag_name = self.tag_selector_combobox.get()
        if not tag_name: return
        self.datasheet_listbox_for_tag.delete(0, tk.END)
        if tag_name in self.tag_associations:
            for datasheet in self.tag_associations[tag_name].get('datasheets', []):
                self.datasheet_listbox_for_tag.insert(tk.END, datasheet)
        self._display_tag_info_in_editor_view(tag_name)

    def _display_tag_info_in_editor_view(self, tag_name):
        self.tag_text_embed.delete(1.0, tk.END)
        self.properties_text_embed.delete(1.0, tk.END)
        if tag_name not in self.tag_associations:
            self.tag_text_embed.insert(tk.END, "Tag not found.")
            return
        self.tag_text_embed.insert(tk.END, f"Tag: {tag_name}")
        nodes = self.tag_associations[tag_name].get('nodes', [])
        if not nodes:
            self.properties_text_embed.insert(tk.END, "Tag has no associated nodes.")
            return
        self.properties_text_embed.insert(tk.END, f"Associated Nodes:\n- {'\n- '.join(nodes)}\n\nProperties:\n")
        self._update_status(f"Fetching properties for tag '{tag_name}'...")
        def task():
            all_properties = set()
            prefix = self.entry_prefix.get().strip()
            for node in nodes:
                query = f"{prefix}\nSELECT DISTINCT ?p WHERE {{ ex:{node} ?p ?o . FILTER(!isBlank(?o)) }}"
                results = self._run_sparql_query(query)
                if results:
                    all_properties.update(p.split('#')[-1] for p in (res['p']['value'] for res in results) if p.split('#')[-1] not in ["hasValue", "hasUnit", "a", "type"])
            prop_text = "\n".join(sorted(list(all_properties))) or "(No direct properties found)"
            self.after(0, lambda: self.properties_text_embed.insert(tk.END, prop_text))
            self.after(0, lambda: self._update_status(f"Info loaded for tag '{tag_name}'.", 4000))
        threading.Thread(target=task, daemon=True).start()

    def _remove_file(self):
        if not self.file_listbox_manage.curselection():
            messagebox.showwarning("No Selection", "Please select a file to remove from the library.")
            return
        file_name = self.file_listbox_manage.get(self.file_listbox_manage.curselection())
        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to permanently delete '{file_name}' from the library? This will also un-tag it from any associations."):
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
        tag = tag_name or (self.tag_listbox.get(self.tag_listbox.curselection()) if self.tag_listbox.curselection() else None)
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
        self._load_excel_file(os.path.join(self.storage_dir, file_name))

if __name__ == "__main__":
    app = LeanDigitalTwin()
    app.mainloop()