import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, font
from tkcalendar import DateEntry
from datetime import datetime
import webbrowser
import os
import tempfile
import subprocess
import time
from MdToClipboard import MdToClipboard


config_file_path = "./config.conf"

with open(config_file_path, "r") as f:
    db_path = f.readline().strip()



# Custom DateEntry class to handle empty values and time
class CustomDateEntry(ttk.Frame):
    def __init__(self, master=None, **kw):
        super().__init__(master)
        self.date_entry = DateEntry(self, **kw)
        self.date_entry.pack(side=tk.LEFT)

        self.hour_spinbox = tk.Spinbox(self, from_=0, to=23, width=2, format="%02.0f")
        self.hour_spinbox.pack(side=tk.LEFT)
        self.hour_spinbox.insert(0, "00")

        self.minute_spinbox = tk.Spinbox(self, from_=0, to=59, width=2, format="%02.0f")
        self.minute_spinbox.pack(side=tk.LEFT)
        self.minute_spinbox.insert(0, "00")

        self._date = None

    def get_date(self):
        date_str = self.date_entry.get()
        hour = self.hour_spinbox.get()[0:2]
        minute = self.minute_spinbox.get()[0:2]
        if date_str and hour and minute:
            try:
                date_ret = datetime.strptime(f"{date_str} {hour}:{minute}", "%Y-%m-%d %H:%M")
                return date_ret.isoformat().replace(' ', 'T')
            except ValueError:
                messagebox.showwarning("Warning", f"Invalid date or time format '%Y-%m-%d %H:%M': {date_str} {hour}:{minute}")
                return None
        return None

    def set_date(self, date_str=None):
        if date_str:
            # Replace 'T' with a space to handle the ISO format
            date_str = date_str.replace('T', ' ')
            # Strip the seconds part if present
            if ':' in date_str.split()[1]:
                date_str = date_str.split(':')[0] + ":" + date_str.split(':')[1]
            date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M")
            self.date_entry.set_date(date_obj.date())
            self.hour_spinbox.delete(0, tk.END)
            self.hour_spinbox.insert(0, date_obj.strftime("%H"))
            self.minute_spinbox.delete(0, tk.END)
            self.minute_spinbox.insert(0, date_obj.strftime("%M"))
            self._date = date_obj
        else:
            self.date_entry.delete(0, tk.END)
            self.hour_spinbox.delete(0, tk.END)
            self.hour_spinbox.insert(0, "00")
            self.minute_spinbox.delete(0, tk.END)
            self.minute_spinbox.insert(0, "00")
            self._date = None

class CustomTreeview(ttk.Treeview):
    def __init__(self, parent, columns, show="headings", height=4, row_offset=0):
        super().__init__(parent, columns=columns, show=show, height=height)

        # Create scrollbar
        self.scrollbar = ttk.Scrollbar(parent, orient='vertical', command=self.yview)
        self.configure(yscrollcommand=self.scrollbar.set)

        # Grid the treeview and scrollbar
        self.grid(column=0, row=row_offset, sticky="we")
        self.scrollbar.grid(column=1, row=row_offset, sticky="ns")

class Database:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()
        self.connexion = self.Connexion(db_path)

    # Database setup
    def init_db(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.executescript('''
        CREATE TABLE IF NOT EXISTS task (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer TEXT,
            name TEXT,
            description TEXT,
            started_at TEXT,
            finished_at TEXT,
            task_id INTEGER,
            FOREIGN KEY(task_id) REFERENCES task(id)
        );

        CREATE TABLE IF NOT EXISTS delivery (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version TEXT,
            server TEXT,
            environment TEXT,
            delivery_date_time TEXT
        );

        CREATE TABLE IF NOT EXISTS task_delivery (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            delivery_id INTEGER,
            task_id INTEGER,
            FOREIGN KEY(delivery_id) REFERENCES delivery(id),
            FOREIGN KEY(task_id) REFERENCES task(id)
        );

        CREATE TABLE IF NOT EXISTS link (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT,
            raw_link TEXT
        );

        CREATE TABLE IF NOT EXISTS task_link (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER,
            link_id INTEGER,
            FOREIGN KEY(task_id) REFERENCES task(id),
            FOREIGN KEY(link_id) REFERENCES link(id)
        );

        CREATE TABLE IF NOT EXISTS tag (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT,
            keywords TEXT
        );

        CREATE TABLE IF NOT EXISTS tag_task (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER,
            tag_id INTEGER,
            FOREIGN KEY(task_id) REFERENCES task(id),
            FOREIGN KEY(tag_id) REFERENCES tag(id)
        );

        CREATE TABLE IF NOT EXISTS tag_link (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            link_id INTEGER,
            tag_id INTEGER,
            FOREIGN KEY(link_id) REFERENCES link(id),
            FOREIGN KEY(tag_id) REFERENCES tag(id)
        );

        CREATE TABLE IF NOT EXISTS origin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            type TEXT,
            raw_link TEXT
        );

        CREATE TABLE IF NOT EXISTS task_origin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER,
            origin_id INTEGER,
            FOREIGN KEY(task_id) REFERENCES task(id),
            FOREIGN KEY(origin_id) REFERENCES origin(id)
        );

        CREATE TABLE IF NOT EXISTS booking (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            description TEXT,
            started_at TEXT,
            ended_at TEXT,
            duration TEXT,
            task_id INTEGER,
            origin_id INTEGER,
            FOREIGN KEY(task_id) REFERENCES task(id),
            FOREIGN KEY(origin_id) REFERENCES origin(id)
        );

        CREATE TABLE IF NOT EXISTS note (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER,
            content TEXT,
            FOREIGN KEY(task_id) REFERENCES task(id)
        );

        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key TEXT UNIQUE,
            value TEXT
        );
        ''')
        conn.commit()
        conn.close()

    def execute_query(self, query, params=None):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return cursor

    def fetch_all(self, query, params=None):
        cursor = self.execute_query(query, params)
        return cursor.fetchall()

    def fetch_one(self, query, params=None):
        cursor = self.execute_query(query, params)
        return cursor.fetchone()

    def save_theme(self, theme):
        self.execute_query("INSERT OR REPLACE INTO settings (key, value) VALUES ('theme', ?)", (theme,))

    def get_theme(self):
        result = self.fetch_one("SELECT value FROM settings WHERE key = 'theme'")
        return result[0] if result else None
    
    def connect(self):
        self.conn = sqlite3.connect(self.db_path)

    def disconnect(self):
        self.conn.close()
    
    def fetch_one_connected(self, query, params=None):
        cursor = self.execute_query(query, params)
        return cursor.fetchone()

    class Connexion:
        def __init__(self, db_path):
            self.db_path = db_path
            self.conn = None
            self._is_connected = False

        def connect(self):
            if self.is_connected():
                raise RuntimeError("Trying to connect while already connected to {self.db_path}.")
            self.conn = sqlite3.connect(self.db_path)
            self._is_connected = True

        def disconnect(self):
            if not self.is_connected():
                raise RuntimeError("Trying to disconnect while not connected to {self.db_path}.")
            self.conn.close()
            self._is_connected = False

        def status(self):
            return "connected" if self._is_connected else "disconnected"
        
        def is_connected(self):
            return self._is_connected
        
        def fetch_all(self, query, params=None):
            cursor = self._execute_query(query, params)
            return cursor.fetchall()

        def fetch_one(self, query, params=None):
            cursor = self._execute_query(query, params)
            return cursor.fetchone()
        
        def _execute_query(self, query, params=None):
            cursor = self.conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            self.conn.commit()
            return cursor

# App class
class TaskManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Task Manager")
        self.db = Database(db_path)
        self.setup_ui()
        self.load_tasks()
        self.temp_file_path = None  # To store the path of the temporary file
        self.selected_related_id = None
        self.related_to_select_next = None  # For selecting related data when clicking on a search result

        # Load the saved theme
        self.load_theme()

    def load_theme(self):
        result = self.db.fetch_one("SELECT value FROM settings WHERE key = 'theme'")

        if result:
            theme = result[0]
            self.change_theme(theme)
        else:
            # Default to normal theme if no saved theme is found
            self.change_theme("normal")

    def setup_ui(self):

        # Define a constant width font (Consolas)
        self.constant_width_font = font.Font(family="Consolas", size=10)

        # Add a menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Create a Theme menu
        theme_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Theme", menu=theme_menu)
        theme_menu.add_command(label="Normal Theme", command=lambda: self.change_theme("normal"))
        theme_menu.add_command(label="Dark Theme", command=lambda: self.change_theme("dark"))
        theme_menu.add_command(label="Old Book Theme", command=lambda: self.change_theme("old_book"))
        theme_menu.add_command(label="Gray and Red Theme", command=lambda: self.change_theme("gray_red"))
        theme_menu.add_command(label="Orange and Blue Theme", command=lambda: self.change_theme("orange_blue"))

        Treeview_height = 3
        row_offset = 0
        col_offset = 0
        self.tree = ttk.Treeview(self.root, columns=("indicator", "customer", "name", "description", "started_at", "finished_at"), show="headings", height=Treeview_height)
        self.tree.heading("indicator", text=" ")
        self.tree.column("indicator", width=45)
        for col in ("customer", "name", "description", "started_at", "finished_at"):
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview(_col, False))
            self.tree.column(col, width=100)
        self.tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.tree.tag_configure("constant_width", font=self.constant_width_font)    # Configure the font for the Treeview
        self.tree.bind("<<TreeviewSelect>>", self.on_task_select)

        btn_frame = tk.Frame(self.root)
        btn_frame.grid(column=0, row=row_offset)
        row_offset += 1

        # Report button
        report_btn_frame = tk.Frame(self.root)
        report_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(report_btn_frame, text="Generate Report", command=self.generate_report).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Button(btn_frame, text="Add Task", command=lambda: self.task_form()).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(btn_frame, text="Add Subtask", command=lambda: self.task_form(subtask=True)).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(btn_frame, text="Edit Task", command=self.edit_task).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(btn_frame, text="Delete Task", command=self.delete_task).grid(column=col_offset, row=0)
        col_offset = 0


        tk.Label(self.root, text="Deliveries").grid(column=0, row=row_offset)
        row_offset += 1
        self.delivery_tree = CustomTreeview(self.root, columns=("version", "server", "environment", "delivery_date_time"), show="headings", height=Treeview_height, row_offset=row_offset)
        self.delivery_tree.heading("version", text="Version")
        self.delivery_tree.heading("server", text="Server")
        self.delivery_tree.heading("environment", text="Environment")
        self.delivery_tree.heading("delivery_date_time", text="Delivery Date Time")
        self.delivery_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.delivery_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "delivery"))

        delivery_btn_frame = tk.Frame(self.root)
        delivery_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(delivery_btn_frame, text="Add Delivery", command=self.add_delivery).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(delivery_btn_frame, text="Edit Delivery", command=self.edit_delivery).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(delivery_btn_frame, text="Delete Delivery", command=self.delete_delivery).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Label(self.root, text="Links").grid(column=0, row=row_offset)
        row_offset += 1
        self.link_tree = CustomTreeview(self.root, columns=("type", "raw_link"), show="headings", height=Treeview_height, row_offset=row_offset)
        self.link_tree.heading("type", text="Type")
        self.link_tree.heading("raw_link", text="Raw Link")
        self.link_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.link_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "link"))
        self.link_tree.bind("<Double-1>", self.open_link)

        link_btn_frame = tk.Frame(self.root)
        link_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(link_btn_frame, text="Add Link", command=self.add_link).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(link_btn_frame, text="Edit Link", command=self.edit_link).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(link_btn_frame, text="Delete Link", command=self.delete_link).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Label(self.root, text="Tags").grid(column=0, row=row_offset)
        row_offset += 1
        self.tag_tree = CustomTreeview(self.root, columns=("type", "keywords"), show="headings", height=Treeview_height, row_offset=row_offset)
        self.tag_tree.heading("type", text="Type")
        self.tag_tree.heading("keywords", text="Keywords")
        self.tag_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.tag_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "tag"))

        tag_btn_frame = tk.Frame(self.root)
        tag_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(tag_btn_frame, text="Add Tag", command=self.add_tag).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(tag_btn_frame, text="Edit Tag", command=self.edit_tag).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(tag_btn_frame, text="Delete Tag", command=self.delete_tag).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Label(self.root, text="Origin").grid(column=0, row=row_offset)
        row_offset += 1
        self.origin_tree = CustomTreeview(self.root, columns=("name", "type", "raw_link"), show="headings", height=Treeview_height, row_offset=row_offset)
        self.origin_tree.heading("name", text="Name")
        self.origin_tree.heading("type", text="Type")
        self.origin_tree.heading("raw_link", text="Raw Link")
        self.origin_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.origin_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "origin"))
        self.origin_tree.bind("<Double-1>", self.open_origin)

        origin_btn_frame = tk.Frame(self.root)
        origin_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(origin_btn_frame, text="Add Origin", command=self.add_origin).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(origin_btn_frame, text="Edit Origin", command=self.edit_origin).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(origin_btn_frame, text="Delete Origin", command=self.delete_origin).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Label(self.root, text="Bookings").grid(column=0, row=row_offset)
        row_offset += 1
        self.booking_tree = CustomTreeview(self.root, columns=("description", "started_at", "ended_at", "duration"), show="headings", height=Treeview_height, row_offset=row_offset)
        self.booking_tree.heading("description", text="Description")
        self.booking_tree.heading("started_at", text="Started At")
        self.booking_tree.heading("ended_at", text="Ended At")
        self.booking_tree.heading("duration", text="Duration")
        self.booking_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.booking_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "booking"))

        booking_btn_frame = tk.Frame(self.root)
        booking_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Button(booking_btn_frame, text="Add Booking", command=self.add_booking).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(booking_btn_frame, text="Edit Booking", command=self.edit_booking).grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(booking_btn_frame, text="Delete Booking", command=self.delete_booking).grid(column=col_offset, row=0)
        col_offset += 1

        tk.Label(self.root, text="Notes").grid(column=0, row=row_offset)
        row_offset += 1
        self.note_tree = CustomTreeview(self.root, columns=("content",), show="headings", height=Treeview_height, row_offset=row_offset)
        self.note_tree.heading("content", text="Content")
        self.note_tree.grid(column=0, row=row_offset)
        row_offset += 1
        self.note_tree.bind("<<TreeviewSelect>>", lambda event: self.on_related_select(event, "note"))

        note_btn_frame = tk.Frame(self.root)
        note_btn_frame.grid(column=0, row=row_offset)
        row_offset += 1
        self.add_note_button = tk.Button(note_btn_frame, text="Add Note", command=lambda: self.add_note(self.tree.selection()[0]), state=tk.NORMAL)
        self.add_note_button.grid(column=col_offset, row=0)
        # print('relief: ' + self.add_note_button['relief'])
        col_offset += 1
        self.edit_note_button = tk.Button(note_btn_frame, text="Edit Note", command=lambda: self.edit_note(self.tree.selection()[0]), state=tk.NORMAL)
        self.edit_note_button.grid(column=col_offset, row=0)
        col_offset += 1
        tk.Button(note_btn_frame, text="Delete Note", command=self.delete_note).grid(column=col_offset, row=0)
        col_offset += 1

        # Add search functionality
        tk.Label(self.root, text="Search").grid(column=0, row=row_offset)
        row_offset += 1
        search_frame = tk.Frame(self.root)
        search_frame.grid(column=0, row=row_offset)
        row_offset += 1
        tk.Label(search_frame, text="Table:").grid(column=0, row=0)
        self.table_var = tk.StringVar()
        table_dropdown = ttk.Combobox(search_frame, textvariable=self.table_var)
        table_dropdown['values'] = ["task", "delivery", "link", "tag", "origin", "booking", "note"]
        table_dropdown.grid(column=1, row=0)
        table_dropdown.bind("<<ComboboxSelected>>", self.update_field_dropdown)
        tk.Label(search_frame, text="Field:").grid(column=2, row=0)
        self.field_var = tk.StringVar()
        self.field_dropdown = ttk.Combobox(search_frame, textvariable=self.field_var)
        self.field_dropdown.grid(column=3, row=0)
        tk.Label(search_frame, text="Operator:").grid(column=4, row=0)
        self.operator_var = tk.StringVar()
        operator_dropdown = ttk.Combobox(search_frame, textvariable=self.operator_var)
        operator_dropdown['values'] = ["LIKE", "=", "!=", "<", ">", "<=", ">="]
        operator_dropdown.grid(column=5, row=0)
        operator_dropdown.current(0)  # Set default operator to LIKE
        tk.Label(search_frame, text="Value:").grid(column=6, row=0)
        self.value_entry = tk.Entry(search_frame)
        self.value_entry.grid(column=7, row=0)
        tk.Button(search_frame, text="Search", command=self.search_data).grid(column=8, row=0)

    def change_theme(self, theme):
        if theme == "normal":
            self.root.config(bg="white")
            style = ttk.Style()
            style.theme_use("default")
        elif theme == "dark":
            self.root.config(bg="#1e1e1e")
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TLabel", background="#1e1e1e", foreground="white")
            style.configure("TButton", background="#3c3f41", foreground="white")
            style.configure("TFrame", background="#1e1e1e")
            style.configure("Treeview", background="#1e1e1e", foreground="white", fieldbackground="#1e1e1e")
        elif theme == "old_book":
            self.root.config(bg="#f5f5dc")
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TLabel", background="#f5f5dc", foreground="#333333")
            style.configure("TButton", background="#c2b280", foreground="#333333")
            style.configure("TFrame", background="#f5f5dc")
            style.configure("Treeview", background="#f5f5dc", foreground="#333333", fieldbackground="#f5f5dc")
        elif theme == "gray_red":
            self.root.config(bg="#cccccc")
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TLabel", background="#cccccc", foreground="#333333")
            style.configure("TButton", background="#ff6347", foreground="white")
            style.configure("TFrame", background="#cccccc")
            style.configure("Treeview", background="#cccccc", foreground="#333333", fieldbackground="#cccccc")
        elif theme == "orange_blue":
            self.root.config(bg="#ffcc99")
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TLabel", background="#ffcc99", foreground="#003366")
            style.configure("TButton", background="#003366", foreground="white")
            style.configure("TFrame", background="#ffcc99")
            style.configure("Treeview", background="#ffcc99", foreground="#003366", fieldbackground="#ffcc99")

        # Save the selected theme to the database
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('theme', ?)", (theme,))
        conn.commit()
        conn.close()

    def generate_report(self):
        selected_items = self.tree.selection()
        selected_items = [int(item) for item in selected_items]
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a task to generate a report")
            return

        # Get the hierarchy of selected tasks
        task_hierarchy = self.get_task_hierarchy(selected_items)

        report_lines = []
        visited_tasks = set()

        def move_matching_row_to_end(array, col_nb = 2, pattern = "PROD"):
            # Find rows where the second column is "PROD"
            prod_rows = [row for row in array if row[col_nb] == pattern]
            
            # Remove these rows from the original array
            filtered_array = [row for row in array if row[col_nb] != pattern]
            
            # Append the "PROD" rows to the end
            result = filtered_array + prod_rows
            
            return result
        

        def sort_array_by_cols(list_of_lists, col_nbs, direction="asc"):
            """
            Sorts a list of lists by one or more column indices.
            
            :param list_of_lists: List of rows (each a list).
            :param col_nbs: A single column index (int) or list of column indices to sort by.
            :param direction: 'asc' for ascending, 'desc' for descending.
            :return: Sorted list of lists.
            """
            if isinstance(col_nbs, int):
                col_nbs = [col_nbs]

            reverse = direction.lower() == "desc"

            return sorted(
                list_of_lists,
                key=lambda x: tuple(x[col] for col in col_nbs),
                reverse=reverse
            )


        def add_task_to_report(task_id, indent_level=0):
            if task_id in visited_tasks:
                return
            visited_tasks.add(task_id)

            self.db.connexion.connect()

            # Get task details
            task = self.db.connexion.fetch_one("SELECT id, customer, name, description, started_at, finished_at FROM task WHERE id = ?", (task_id,))

            if task:
                one = 1
                indent = "    " * indent_level
                sub_indent = "    " * one
                if task[1] != "":
                    report_lines.append(f"{indent}<p><strong>{task[1]}</strong>: <code>{task[2]}</code><p><ul>")
                    # report_lines.append(f"{indent}<p>{task[1]}</p><ul>")
                    report_lines.append(f"{indent}{sub_indent}<li>Description: {task[3]}</li>")
                else:
                    report_lines.append(f"{indent}<li>Sub-task: <code>{task[2]}</code></li><ul>")
                    report_lines.append(f"{indent}{sub_indent}<li>Description: {task[3]}</li>")

                # Get related deliveries
                deliveries = self.db.connexion.fetch_all("SELECT d.version, d.server, d.environment, d.delivery_date_time FROM delivery d JOIN task_delivery td ON d.id = td.delivery_id WHERE td.task_id = ?", (task_id,))
                if deliveries:
                    report_lines.append(f"{indent}{sub_indent}<li>Deliveries:</li><ul>")

                    # Order deliveries
                    deliveries = sort_array_by_cols(deliveries, [0, 2])
                    deliveries = move_matching_row_to_end(deliveries, 2, "PROD")

                    last_version_and_server = None
                    delivery_date = ""
                    close_list_tag = ""
                    for delivery in deliveries:
                        version_and_server = f"V {delivery[0]}, {delivery[1]}"
                        if version_and_server != last_version_and_server :
                            report_lines.append(f"{close_list_tag}{indent}{sub_indent}{sub_indent}<li>{version_and_server}:</li><ul>") # new version or server
                            close_list_tag = "</ul>"
                        delivery_date = delivery[3][:-3].replace("T", " ").replace("-", ".").replace(":", "h")
                        report_lines.append(f"{indent}{sub_indent}{sub_indent}{sub_indent}<li>[x] {delivery[2]}, {delivery_date}")
                        last_version_and_server = version_and_server

                    report_lines.append("</ul></ul>")

                # Get related origins
                origins = self.db.connexion.fetch_all("SELECT o.name, o.type, o.raw_link FROM origin o JOIN task_origin t_o ON o.id = t_o.origin_id WHERE t_o.task_id = ?", (task_id,))
                if origins:
                    for origin in origins:
                        report_lines.append(f"{indent}{sub_indent}<li><a href=\"{origin[2]}\">BCS: {origin[0]}</a></li>")
                    
            self.db.connexion.disconnect()

            # Recursively add child tasks
            if task_id in task_hierarchy:
                for child_id in task_hierarchy[task_id]:
                    add_task_to_report(child_id, indent_level + 1)

        # Generate the report with hierarchy
        for task_id in selected_items:
            add_task_to_report(task_id)
            report_lines.append("</ul>")

        report_text = "\n".join(report_lines)

        MdToClipboard.html_to_clipboard_for_onenote(report_text)
        messagebox.showinfo("Info", "Report copied to clipboard")

    def get_task_hierarchy(self, task_ids):
        self.db.connexion.connect()
        task_hierarchy = {task_id: [] for task_id in task_ids}
        subtasks = []
        for task_id in task_ids:
            children = self.db.connexion.fetch_all("SELECT id FROM task WHERE task_id = ?", (task_id,))
            filtered_children = [child[0] for child in children if child[0] in task_ids]
            if len(filtered_children) > 0:
                for filtered_child in filtered_children:
                    subtasks.append(filtered_child)
            if task_id in subtasks:
                del task_hierarchy[task_id]
            else:
                task_hierarchy[task_id].extend(filtered_children)

        self.db.connexion.disconnect()
        return task_hierarchy



    def update_field_dropdown(self, event):
        table = self.table_var.get()
        if table:
            fields = self.db.fetch_all(f"PRAGMA table_info({table})")
            field_names = [field[1] for field in fields]
            self.field_dropdown['values'] = field_names
            if field_names:
                self.field_var.set(field_names[0])


    def load_tasks(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Load parent tasks
        parent_tasks = self.db.fetch_all("SELECT id, customer, name, description, started_at, finished_at, task_id FROM task WHERE task_id IS NULL")

        self.tree.delete(*self.tree.get_children())

        # Insert parent tasks into the tree
        for task in parent_tasks:
            self.tree.insert("", "end", iid=task[0], values=("─────", task[1], task[2], task[3], task[4], task[5]), tags=("constant_width",))
            self.load_children_tasks(task[0], "")


    def load_children_tasks(self, parent_id, parent_iid):
        # Load children tasks for the given parent_id
        children_tasks = self.db.fetch_all("SELECT id, customer, name, description, started_at, finished_at, task_id FROM task WHERE task_id = ?", (parent_id,))

        for task in children_tasks:
            indicator = "  └──"
            task_iid = self.tree.insert(parent_iid, "end", iid=task[0], values=(indicator, task[1], task[2], task[3], task[4], task[5]), tags=("constant_width",))
            self.load_children_tasks(task[0], task_iid)


    def on_task_select(self, event):
        selected_items = self.tree.selection()
        if selected_items:
            self.load_related_data(selected_items)
            if self.related_to_select_next:
                # Select the related data in the corresponding treeview
                table, related_id = self.related_to_select_next
                if table == "delivery":
                    self.delivery_tree.selection_set(related_id)
                    self.delivery_tree.focus(related_id)
                    self.delivery_tree.see(related_id)
                elif table == "link":
                    self.link_tree.selection_set(related_id)
                    self.link_tree.focus(related_id)
                    self.link_tree.see(related_id)
                elif table == "tag":
                    self.tag_tree.selection_set(related_id)
                    self.tag_tree.focus(related_id)
                    self.tag_tree.see(related_id)
                elif table == "origin":
                    self.origin_tree.selection_set(related_id)
                    self.origin_tree.focus(related_id)
                    self.origin_tree.see(related_id)
                elif table == "booking":
                    self.booking_tree.selection_set(related_id)
                    self.booking_tree.focus(related_id)
                    self.booking_tree.see(related_id)
                elif table == "note":
                    self.note_tree.selection_set(related_id)
                    self.note_tree.focus(related_id)
                    self.note_tree.see(related_id)
                self.related_to_select_next = None  # Reset after use

    def on_related_select(self, event, related_type):
        selected_item = event.widget.selection()
        if selected_item:
            self.selected_related_id = selected_item[0]
            self.selected_related_type = related_type

    def add_task(self):
        self.task_form()

    def edit_task(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to edit")
            return
        task_id = selected_item[0]
        task = self.db.fetch_one("SELECT * FROM task WHERE id = ?", (task_id,))
        if task:
            self.task_form(task)

    def delete_task(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to delete")
            return
        task_id = selected_item[0]
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM task WHERE id = ?", (task_id,))
        conn.commit()
        conn.close()
        self.load_tasks()


    def task_form(self, task=None, subtask=False):
        form = tk.Toplevel(self.root)
        form.title("Task Form")
        tk.Label(form, text="Customer:").grid(row=0, column=0)

        # Use a Combobox for the customer field
        self.customer_var = tk.StringVar()
        self.customer_combobox = ttk.Combobox(form, textvariable=self.customer_var)
        self.customer_combobox.grid(row=0, column=1)
        self.customer_combobox.bind("<KeyRelease>", self.update_customer_combobox)

        tk.Label(form, text="Name:").grid(row=1, column=0)
        self.name_var = tk.StringVar()
        self.name_combobox = ttk.Combobox(form, textvariable=self.name_var)
        self.name_combobox.grid(row=1, column=1)
        self.name_combobox.bind("<KeyRelease>", self.update_name_combobox)

        tk.Label(form, text="Description:").grid(row=2, column=0)
        self.desc_var = tk.StringVar()
        self.desc_combobox = ttk.Combobox(form, textvariable=self.desc_var)
        self.desc_combobox.grid(row=2, column=1)
        self.desc_combobox.bind("<KeyRelease>", self.update_desc_combobox)

        tk.Label(form, text="Started At:").grid(row=3, column=0)
        started_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
        started_entry.grid(row=3, column=1)
        tk.Label(form, text="Finished At:").grid(row=4, column=0)
        finished_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
        finished_entry.grid(row=4, column=1)

        parent_task_id = None
        if subtask:
            selected_item = self.tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Please select a parent task for the subtask")
                form.destroy()
                return
            parent_task_id = selected_item[0]

        if task:
            self.customer_var.set(task[1])
            self.name_var.set(task[2])
            self.desc_var.set(task[3])
            if task[4]:
                started_entry.set_date(task[4])
            if task[5]:
                finished_entry.set_date(task[5])

        def save_task():
            customer = self.customer_var.get()
            name = self.name_var.get()
            desc = self.desc_var.get()
            started = started_entry.get_date()
            finished = finished_entry.get_date()
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            if task:
                cursor.execute("UPDATE task SET customer=?, name=?, description=?, started_at=?, finished_at=? WHERE id=?", (customer, name, desc, started, finished, task[0]))
            else:
                cursor.execute("INSERT INTO task (customer, name, description, started_at, finished_at, task_id) VALUES (?, ?, ?, ?, ?, ?)", (customer, name, desc, started, finished, parent_task_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_tasks()

        tk.Button(form, text="Save", command=save_task).grid(row=5, column=0, columnspan=2)

        # Populate the comboboxes with existing values
        self.update_customer_combobox()
        self.update_name_combobox()
        self.update_desc_combobox()

    def update_customer_combobox(self, event=None):
        customer_input = self.customer_var.get()
        customers = self.db.fetch_all("SELECT DISTINCT customer FROM task WHERE customer LIKE ?", ('%' + customer_input + '%',))

        # Update the combobox values
        self.customer_combobox['values'] = [customer[0] for customer in customers]

    def update_name_combobox(self, event=None):
        name_input = self.name_var.get()
        names = self.db.fetch_all("SELECT DISTINCT name FROM task WHERE name LIKE ?", ('%' + name_input + '%',))

        # Update the combobox values
        self.name_combobox['values'] = [name[0] for name in names]

    def update_desc_combobox(self, event=None):
        desc_input = self.desc_var.get()
        descriptions = self.db.fetch_all("SELECT DISTINCT description FROM task WHERE description LIKE ?", ('%' + desc_input + '%',))


        # Update the combobox values
        self.desc_combobox['values'] = [desc[0] for desc in descriptions]


    def sort_treeview(self, col, descending):
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]
        data.sort(reverse=descending)
        for index, (val, child) in enumerate(data):
            self.tree.move(child, "", index)
        self.tree.heading(col, command=lambda: self.sort_treeview(col, not descending))

    def load_related_data(self, task_ids):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        deliveries = set()
        links = set()
        tags = set()
        origins = set()
        bookings = set()
        notes = set()

        for task_id in task_ids:
            delivery_rel = self.db.fetch_all("""
            SELECT d.id, d.version, d.server, d.environment, d.delivery_date_time
            FROM delivery d
            JOIN task_delivery td ON d.id = td.delivery_id
            WHERE td.task_id = ?""", (task_id,))
            deliveries.update(delivery_rel)

            link_rel = self.db.fetch_all("""
            SELECT l.id, l.type, l.raw_link
            FROM link l
            JOIN task_link tl ON l.id = tl.link_id
            WHERE tl.task_id = ?""", (task_id,))
            links.update(link_rel)

            tag_rel = self.db.fetch_all("""
            SELECT t.id, t.type, t.keywords
            FROM tag t
            JOIN tag_task tt ON t.id = tt.tag_id
            WHERE tt.task_id = ?""", (task_id,))
            tags.update(tag_rel)

            origin_rel = self.db.fetch_all("""
            SELECT o.id, o.name, o.type, o.raw_link
            FROM origin o
            JOIN task_origin t_o ON o.id = t_o.origin_id
            WHERE t_o.task_id = ?""", (task_id,))
            origins.update(origin_rel)

            booking_rel = self.db.fetch_all("""
            SELECT b.id, b.description, b.started_at, b.ended_at, b.duration
            FROM booking b
            WHERE b.task_id = ?""", (task_id,))
            bookings.update(booking_rel)
 
            note_rel = self.db.fetch_all("""
            SELECT n.id, n.content
            FROM note n
            WHERE n.task_id = ?""", (task_id,))
            notes.update(note_rel)


        self.delivery_tree.delete(*self.delivery_tree.get_children())
        for delivery in deliveries:
            self.delivery_tree.insert("", "end", iid=delivery[0], values=(delivery[1], delivery[2], delivery[3], delivery[4]))

        self.link_tree.delete(*self.link_tree.get_children())
        for link in links:
            self.link_tree.insert("", "end", iid=link[0], values=(link[1], link[2]))

        self.tag_tree.delete(*self.tag_tree.get_children())
        for tag in tags:
            self.tag_tree.insert("", "end", iid=tag[0], values=(tag[1], tag[2]))

        self.origin_tree.delete(*self.origin_tree.get_children())
        for origin in origins:
            self.origin_tree.insert("", "end", iid=origin[0], values=(origin[1], origin[2], origin[3]))

        self.booking_tree.delete(*self.booking_tree.get_children())
        for booking in bookings:
            self.booking_tree.insert("", "end", iid=booking[0], values=(booking[1], booking[2], booking[3], booking[4]))

        self.note_tree.delete(*self.note_tree.get_children())
        note_lines = None
        note_title = ''
        for note in notes:
            if note and len(note) > 1:
                note_lines = note[1].splitlines()
                note_title = ''
                if len(note[1]) > 1 and len(note_lines) >= 1:
                    note_title = note_lines[0]
                self.note_tree.insert("", "end", iid=note[0], values=(note_title,))


    def add_delivery(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to add a delivery")
            return
        task_id = selected_item[0]
        form = tk.Toplevel(self.root)
        form.title("Add Delivery")
        tk.Label(form, text="Version:").grid(row=0, column=0)
        self.version_var = tk.StringVar()
        self.version_combobox = ttk.Combobox(form, textvariable=self.version_var)
        self.version_combobox.grid(row=0, column=1)
        self.version_combobox.bind("<KeyRelease>", self.update_version_combobox)

        tk.Label(form, text="Server:").grid(row=1, column=0)
        self.server_var = tk.StringVar()
        self.server_combobox = ttk.Combobox(form, textvariable=self.server_var)
        self.server_combobox.grid(row=1, column=1)
        self.server_combobox.bind("<KeyRelease>", self.update_server_combobox)

        tk.Label(form, text="Environment:").grid(row=2, column=0)
        self.environment_var = tk.StringVar()
        self.environment_combobox = ttk.Combobox(form, textvariable=self.environment_var)
        self.environment_combobox.grid(row=2, column=1)
        self.environment_combobox.bind("<KeyRelease>", self.update_environment_combobox)

        tk.Label(form, text="Delivery Date Time:").grid(row=3, column=0)
        delivery_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
        delivery_entry.grid(row=3, column=1)

        def save_delivery():
            version = self.version_var.get()
            server = self.server_var.get()
            environment = self.environment_var.get()
            delivery_date_time = delivery_entry.get_date()
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO delivery (version, server, environment, delivery_date_time) VALUES (?, ?, ?, ?)", (version, server, environment, delivery_date_time))
            delivery_id = cursor.lastrowid
            cursor.execute("INSERT INTO task_delivery (delivery_id, task_id) VALUES (?, ?)", (delivery_id, task_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_related_data(task_id)

        tk.Button(form, text="Save", command=save_delivery).grid(row=4, column=0, columnspan=2)

        # Populate the comboboxes with existing values
        self.update_version_combobox()
        self.update_server_combobox()
        self.update_environment_combobox()

    def update_version_combobox(self, event=None):
        version_input = self.version_var.get()
        versions = self.db.fetch_all("SELECT DISTINCT version FROM delivery WHERE version LIKE ?", ('%' + version_input + '%',))
        self.version_combobox['values'] = [version[0] for version in versions]

    def update_server_combobox(self, event=None):
        server_input = self.server_var.get()
        servers = self.db.fetch_all("SELECT DISTINCT server FROM delivery WHERE server LIKE ?", ('%' + server_input + '%',))
        self.server_combobox['values'] = [server[0] for server in servers]

    def update_environment_combobox(self, event=None):
        environment_input = self.environment_var.get()
        environments = self.db.fetch_all("SELECT DISTINCT environment FROM delivery WHERE environment LIKE ?", ('%' + environment_input + '%',))
        self.environment_combobox['values'] = [environment[0] for environment in environments]


    def edit_delivery(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "delivery":
            messagebox.showwarning("Warning", "Please select a delivery to edit")
            return
        delivery_id = self.selected_related_id
        delivery = self.db.fetch_one("SELECT * FROM delivery WHERE id = ?", (delivery_id,))
        if delivery:
            form = tk.Toplevel(self.root)
            form.title("Edit Delivery")
            tk.Label(form, text="Version:").grid(row=0, column=0)
            self.version_var = tk.StringVar(value=delivery[1])
            self.version_combobox = ttk.Combobox(form, textvariable=self.version_var)
            self.version_combobox.grid(row=0, column=1)
            self.version_combobox.bind("<KeyRelease>", self.update_version_combobox)

            tk.Label(form, text="Server:").grid(row=1, column=0)
            self.server_var = tk.StringVar(value=delivery[2])
            self.server_combobox = ttk.Combobox(form, textvariable=self.server_var)
            self.server_combobox.grid(row=1, column=1)
            self.server_combobox.bind("<KeyRelease>", self.update_server_combobox)

            tk.Label(form, text="Environment:").grid(row=2, column=0)
            self.environment_var = tk.StringVar(value=delivery[3])
            self.environment_combobox = ttk.Combobox(form, textvariable=self.environment_var)
            self.environment_combobox.grid(row=2, column=1)
            self.environment_combobox.bind("<KeyRelease>", self.update_environment_combobox)

            tk.Label(form, text="Delivery Date Time:").grid(row=3, column=0)
            delivery_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
            delivery_entry.grid(row=3, column=1)
            if delivery[4]:
                delivery_entry.set_date(delivery[4])
            else:
                delivery_entry.set_date()

            def save_delivery():
                version = self.version_var.get()
                server = self.server_var.get()
                environment = self.environment_var.get()
                delivery_date_time = delivery_entry.get_date()
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE delivery SET version=?, server=?, environment=?, delivery_date_time=? WHERE id=?", (version, server, environment, delivery_date_time, delivery_id))
                conn.commit()
                conn.close()
                form.destroy()
                self.load_related_data(self.tree.selection()[0])

            tk.Button(form, text="Save", command=save_delivery).grid(row=4, column=0, columnspan=2)

            # Populate the comboboxes with existing values
            self.update_version_combobox()
            self.update_server_combobox()
            self.update_environment_combobox()

    def delete_delivery(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "delivery":
            messagebox.showwarning("Warning", "Please select a delivery to delete")
            return
        delivery_id = self.selected_related_id
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM delivery WHERE id = ?", (delivery_id,))
        conn.commit()
        conn.close()
        self.load_related_data(self.tree.selection()[0])

    def add_link(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to add a link")
            return
        task_id = selected_item[0]
        form = tk.Toplevel(self.root)
        form.title("Add Link")
        tk.Label(form, text="Type:").grid(row=0, column=0)
        self.link_type_var = tk.StringVar()
        self.link_type_combobox = ttk.Combobox(form, textvariable=self.link_type_var)
        self.link_type_combobox.grid(row=0, column=1)
        self.link_type_combobox.bind("<KeyRelease>", self.update_link_type_combobox)

        tk.Label(form, text="Raw Link:").grid(row=1, column=0)
        link_entry = tk.Entry(form)
        link_entry.grid(row=1, column=1)

        def save_link():
            link_type = self.link_type_var.get()
            raw_link = link_entry.get()
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO link (type, raw_link) VALUES (?, ?)", (link_type, raw_link))
            link_id = cursor.lastrowid
            cursor.execute("INSERT INTO task_link (link_id, task_id) VALUES (?, ?)", (link_id, task_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_related_data(task_id)

        tk.Button(form, text="Save", command=save_link).grid(row=2, column=0, columnspan=2)

        # Populate the combobox with existing values
        self.update_link_type_combobox()

    def update_link_type_combobox(self, event=None):
        link_type_input = self.link_type_var.get()
        link_types = self.db.fetch_all("SELECT DISTINCT type FROM link WHERE type LIKE ?", ('%' + link_type_input + '%',))
        self.link_type_combobox['values'] = [link_type[0] for link_type in link_types]


    def edit_link(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "link":
            messagebox.showwarning("Warning", "Please select a link to edit")
            return
        link_id = self.selected_related_id
        link = self.db.fetch_one("SELECT * FROM link WHERE id = ?", (link_id,))
        if link:
            form = tk.Toplevel(self.root)
            form.title("Edit Link")
            tk.Label(form, text="Type:").grid(row=0, column=0)
            self.link_type_var = tk.StringVar(value=link[1])
            self.link_type_combobox = ttk.Combobox(form, textvariable=self.link_type_var)
            self.link_type_combobox.grid(row=0, column=1)
            self.link_type_combobox.bind("<KeyRelease>", self.update_link_type_combobox)

            tk.Label(form, text="Raw Link:").grid(row=1, column=0)
            link_entry = tk.Entry(form)
            link_entry.grid(row=1, column=1)
            link_entry.insert(0, link[2])

            def save_link():
                link_type = self.link_type_var.get()
                raw_link = link_entry.get()
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE link SET type=?, raw_link=? WHERE id=?", (link_type, raw_link, link_id))
                conn.commit()
                conn.close()
                form.destroy()
                self.load_related_data(self.tree.selection()[0])

            tk.Button(form, text="Save", command=save_link).grid(row=2, column=0, columnspan=2)

            # Populate the combobox with existing values
            self.update_link_type_combobox()


    def delete_link(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "link":
            messagebox.showwarning("Warning", "Please select a link to delete")
            return
        link_id = self.selected_related_id
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM link WHERE id = ?", (link_id,))
        conn.commit()
        conn.close()
        self.load_related_data(self.tree.selection()[0])

    def add_tag(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to add a tag")
            return
        task_id = selected_item[0]
        form = tk.Toplevel(self.root)
        form.title("Add Tag")
        tk.Label(form, text="Type:").grid(row=0, column=0)
        self.tag_type_var = tk.StringVar()
        self.tag_type_combobox = ttk.Combobox(form, textvariable=self.tag_type_var)
        self.tag_type_combobox.grid(row=0, column=1)
        self.tag_type_combobox.bind("<KeyRelease>", self.update_tag_type_combobox)

        tk.Label(form, text="Keywords:").grid(row=1, column=0)
        self.keywords_var = tk.StringVar()
        self.keywords_combobox = ttk.Combobox(form, textvariable=self.keywords_var)
        self.keywords_combobox.grid(row=1, column=1)
        self.keywords_combobox.bind("<KeyRelease>", self.update_keywords_combobox)

        def save_tag():
            tag_type = self.tag_type_var.get()
            keywords = self.keywords_var.get()
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO tag (type, keywords) VALUES (?, ?)", (tag_type, keywords))
            tag_id = cursor.lastrowid
            cursor.execute("INSERT INTO tag_task (tag_id, task_id) VALUES (?, ?)", (tag_id, task_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_related_data(task_id)

        tk.Button(form, text="Save", command=save_tag).grid(row=2, column=0, columnspan=2)

        # Populate the comboboxes with existing values
        self.update_tag_type_combobox()
        self.update_keywords_combobox()

    def update_tag_type_combobox(self, event=None):
        tag_type_input = self.tag_type_var.get()
        tag_types = self.db.fetch_all("SELECT DISTINCT type FROM tag WHERE type LIKE ?", ('%' + tag_type_input + '%',))
        self.tag_type_combobox['values'] = [tag_type[0] for tag_type in tag_types]

    def update_keywords_combobox(self, event=None):
        keywords_input = self.keywords_var.get()
        keywords = self.db.fetch_all("SELECT DISTINCT keywords FROM tag WHERE keywords LIKE ?", ('%' + keywords_input + '%',))
        self.keywords_combobox['values'] = [keyword[0] for keyword in keywords]


    def edit_tag(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "tag":
            messagebox.showwarning("Warning", "Please select a tag to edit")
            return
        tag_id = self.selected_related_id
        tag = self.db.fetch_one("SELECT * FROM tag WHERE id = ?", (tag_id,))
        if tag:
            form = tk.Toplevel(self.root)
            form.title("Edit Tag")
            tk.Label(form, text="Type:").grid(row=0, column=0)
            self.tag_type_var = tk.StringVar(value=tag[1])
            self.tag_type_combobox = ttk.Combobox(form, textvariable=self.tag_type_var)
            self.tag_type_combobox.grid(row=0, column=1)
            self.tag_type_combobox.bind("<KeyRelease>", self.update_tag_type_combobox)

            tk.Label(form, text="Keywords:").grid(row=1, column=0)
            self.keywords_var = tk.StringVar(value=tag[2])
            self.keywords_combobox = ttk.Combobox(form, textvariable=self.keywords_var)
            self.keywords_combobox.grid(row=1, column=1)
            self.keywords_combobox.bind("<KeyRelease>", self.update_keywords_combobox)

            def save_tag():
                tag_type = self.tag_type_var.get()
                keywords = self.keywords_var.get()
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE tag SET type=?, keywords=? WHERE id=?", (tag_type, keywords, tag_id))
                conn.commit()
                conn.close()
                form.destroy()
                self.load_related_data(self.tree.selection()[0])

            tk.Button(form, text="Save", command=save_tag).grid(row=2, column=0, columnspan=2)

            # Populate the comboboxes with existing values
            self.update_tag_type_combobox()
            self.update_keywords_combobox()

    def delete_tag(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "tag":
            messagebox.showwarning("Warning", "Please select a tag to delete")
            return
        tag_id = self.selected_related_id
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM tag WHERE id = ?", (tag_id,))
        conn.commit()
        conn.close()
        self.load_related_data(self.tree.selection()[0])

    def add_origin(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to add an origin")
            return
        task_id = selected_item[0]
        form = tk.Toplevel(self.root)
        form.title("Add Origin")
        tk.Label(form, text="Name:").grid(row=0, column=0)
        self.origin_name_var = tk.StringVar()
        self.origin_name_combobox = ttk.Combobox(form, textvariable=self.origin_name_var)
        self.origin_name_combobox.grid(row=0, column=1)
        self.origin_name_combobox.bind("<KeyRelease>", self.update_origin_name_combobox)

        tk.Label(form, text="Type:").grid(row=1, column=0)
        self.origin_type_var = tk.StringVar()
        self.origin_type_combobox = ttk.Combobox(form, textvariable=self.origin_type_var)
        self.origin_type_combobox.grid(row=1, column=1)
        self.origin_type_combobox.bind("<KeyRelease>", self.update_origin_type_combobox)

        tk.Label(form, text="Raw Link:").grid(row=2, column=0)
        link_entry = tk.Entry(form)
        link_entry.grid(row=2, column=1)

        def save_origin():
            origin_name = self.origin_name_var.get()
            origin_type = self.origin_type_var.get()
            raw_link = link_entry.get()
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO origin (name, type, raw_link) VALUES (?, ?, ?)", (origin_name, origin_type, raw_link))
            origin_id = cursor.lastrowid
            cursor.execute("INSERT INTO task_origin (origin_id, task_id) VALUES (?, ?)", (origin_id, task_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_related_data(task_id)

        tk.Button(form, text="Save", command=save_origin).grid(row=3, column=0, columnspan=2)

        # Populate the comboboxes with existing values
        self.update_origin_name_combobox()
        self.update_origin_type_combobox()

    def update_origin_name_combobox(self, event=None):
        origin_name_input = self.origin_name_var.get()
        origin_names = self.db.fetch_all("SELECT DISTINCT name FROM origin WHERE name LIKE ?", ('%' + origin_name_input + '%',))
        self.origin_name_combobox['values'] = [origin_name[0] for origin_name in origin_names]

    def update_origin_type_combobox(self, event=None):
        origin_type_input = self.origin_type_var.get()
        origin_types = self.db.fetch_all("SELECT DISTINCT type FROM origin WHERE type LIKE ?", ('%' + origin_type_input + '%',))
        self.origin_type_combobox['values'] = [origin_type[0] for origin_type in origin_types]


    def edit_origin(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "origin":
            messagebox.showwarning("Warning", "Please select an origin to edit")
            return
        origin_id = self.selected_related_id
        origin = self.db.fetch_one("SELECT * FROM origin WHERE id = ?", (origin_id,))
        if origin:
            form = tk.Toplevel(self.root)
            form.title("Edit Origin")
            tk.Label(form, text="Name:").grid(row=0, column=0)
            self.origin_name_var = tk.StringVar(value=origin[1])
            self.origin_name_combobox = ttk.Combobox(form, textvariable=self.origin_name_var)
            self.origin_name_combobox.grid(row=0, column=1)
            self.origin_name_combobox.bind("<KeyRelease>", self.update_origin_name_combobox)

            tk.Label(form, text="Type:").grid(row=1, column=0)
            self.origin_type_var = tk.StringVar(value=origin[2])
            self.origin_type_combobox = ttk.Combobox(form, textvariable=self.origin_type_var)
            self.origin_type_combobox.grid(row=1, column=1)
            self.origin_type_combobox.bind("<KeyRelease>", self.update_origin_type_combobox)

            tk.Label(form, text="Raw Link:").grid(row=2, column=0)
            link_entry = tk.Entry(form)
            link_entry.grid(row=2, column=1)
            link_entry.insert(0, origin[3])

            def save_origin():
                origin_name = self.origin_name_var.get()
                origin_type = self.origin_type_var.get()
                raw_link = link_entry.get()
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE origin SET name=?, type=?, raw_link=? WHERE id=?", (origin_name, origin_type, raw_link, origin_id))
                conn.commit()
                conn.close()
                form.destroy()
                self.load_related_data(self.tree.selection()[0])

            tk.Button(form, text="Save", command=save_origin).grid(row=3, column=0, columnspan=2)

            # Populate the comboboxes with existing values
            self.update_origin_name_combobox()
            self.update_origin_type_combobox()

    def delete_origin(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "origin":
            messagebox.showwarning("Warning", "Please select an origin to delete")
            return
        origin_id = self.selected_related_id
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM origin WHERE id = ?", (origin_id,))
        conn.commit()
        conn.close()
        self.load_related_data(self.tree.selection()[0])

    def open_link(self, event):
        selected_item = self.link_tree.selection()
        if selected_item:
            link_id = selected_item[0]
            link = self.db.fetch_one("SELECT raw_link FROM link WHERE id = ?", (link_id,))
            if link:
                webbrowser.open_new_tab(link[0])

    def open_origin(self, event):
        selected_item = self.origin_tree.selection()
        if selected_item:
            origin_id = selected_item[0]
            origin = self.db.fetch_one("SELECT raw_link FROM origin WHERE id = ?", (origin_id,))
            if origin:
                webbrowser.open_new_tab(origin[0])

    def add_booking(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a task to add a booking")
            return
        task_id = selected_item[0]
        form = tk.Toplevel(self.root)
        form.title("Add Booking")
        tk.Label(form, text="Description:").grid(row=0, column=0)
        self.booking_desc_var = tk.StringVar()
        self.booking_desc_combobox = ttk.Combobox(form, textvariable=self.booking_desc_var)
        self.booking_desc_combobox.grid(row=0, column=1)
        self.booking_desc_combobox.bind("<KeyRelease>", self.update_booking_desc_combobox)

        tk.Label(form, text="Started At:").grid(row=1, column=0)
        started_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
        started_entry.grid(row=1, column=1)
        tk.Label(form, text="Ended At:").grid(row=2, column=0)
        ended_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
        ended_entry.grid(row=2, column=1)
        tk.Label(form, text="Duration:").grid(row=3, column=0)
        duration_entry = tk.Entry(form)
        duration_entry.grid(row=3, column=1)

        origins = self.db.fetch_all("SELECT id, name FROM origin")

        tk.Label(form, text="Origin:").grid(row=4, column=0)
        origin_var = tk.StringVar()
        origin_dropdown = ttk.Combobox(form, textvariable=origin_var)
        origin_dropdown['values'] = [origin[1] for origin in origins]
        origin_dropdown.grid(row=4, column=1)

        def save_booking():
            description = self.booking_desc_var.get()
            started_at = started_entry.get_date()
            ended_at = ended_entry.get_date()
            duration = duration_entry.get()
            origin_name = origin_var.get()

            if not origin_name:
                messagebox.showwarning("Warning", "Please select an origin")
                return

            if not started_at or not ended_at:
                messagebox.showwarning("Warning", "Please enter valid start and end times")
                return

            origin_id = next((origin[0] for origin in origins if origin[1] == origin_name), None)
            if origin_id is None:
                messagebox.showwarning("Warning", "Selected origin does not exist")
                return

            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO booking (description, started_at, ended_at, duration, task_id, origin_id) VALUES (?, ?, ?, ?, ?, ?)", (description, started_at, ended_at, duration, task_id, origin_id))
            conn.commit()
            conn.close()
            form.destroy()
            self.load_related_data(task_id)

        tk.Button(form, text="Save", command=save_booking).grid(row=5, column=0, columnspan=2)

        # Populate the combobox with existing values
        self.update_booking_desc_combobox()

    def update_booking_desc_combobox(self, event=None):
        booking_desc_input = self.booking_desc_var.get()
        booking_descs = self.db.fetch_all("SELECT DISTINCT description FROM booking WHERE description LIKE ?", ('%' + booking_desc_input + '%',))
        self.booking_desc_combobox['values'] = [booking_desc[0] for booking_desc in booking_descs]


    def edit_booking(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "booking":
            messagebox.showwarning("Warning", "Please select a booking to edit")
            return
        booking_id = self.selected_related_id
        booking = self.db.fetch_one("SELECT * FROM booking WHERE id = ?", (booking_id,))
        if booking:
            form = tk.Toplevel(self.root)
            form.title("Edit Booking")
            tk.Label(form, text="Description:").grid(row=0, column=0)
            self.booking_desc_var = tk.StringVar(value=booking[1])
            self.booking_desc_combobox = ttk.Combobox(form, textvariable=self.booking_desc_var)
            self.booking_desc_combobox.grid(row=0, column=1)
            self.booking_desc_combobox.bind("<KeyRelease>", self.update_booking_desc_combobox)

            tk.Label(form, text="Started At:").grid(row=1, column=0)
            started_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
            started_entry.grid(row=1, column=1)
            if booking[2]:
                started_entry.set_date(booking[2])
            else:
                started_entry.set_date()

            tk.Label(form, text="Ended At:").grid(row=2, column=0)
            ended_entry = CustomDateEntry(form, date_pattern="yyyy-mm-dd")
            ended_entry.grid(row=2, column=1)
            if booking[3]:
                ended_entry.set_date(booking[3])
            else:
                ended_entry.set_date()

            tk.Label(form, text="Duration:").grid(row=3, column=0)
            duration_entry = tk.Entry(form)
            duration_entry.grid(row=3, column=1)
            duration_entry.insert(0, booking[4])

            origins = self.db.fetch_all("SELECT id, name FROM origin")

            tk.Label(form, text="Origin:").grid(row=4, column=0)
            origin_var = tk.StringVar()
            origin_dropdown = ttk.Combobox(form, textvariable=origin_var)
            origin_dropdown['values'] = [origin[1] for origin in origins]
            origin_dropdown.grid(row=4, column=1)
            origin_var.set(next((origin[1] for origin in origins if origin[0] == booking[6]), ""))

            def save_booking():
                description = self.booking_desc_var.get()
                started_at = started_entry.get_date()
                ended_at = ended_entry.get_date()
                duration = duration_entry.get()
                origin_name = origin_var.get()

                if not origin_name:
                    messagebox.showwarning("Warning", "Please select an origin")
                    return

                origin_id = next((origin[0] for origin in origins if origin[1] == origin_name), None)
                if origin_id is None:
                    messagebox.showwarning("Warning", "Selected origin does not exist")
                    return

                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("UPDATE booking SET description=?, started_at=?, ended_at=?, duration=?, origin_id=? WHERE id=?", (description, started_at, ended_at, duration, origin_id, booking_id))
                conn.commit()
                conn.close()
                form.destroy()
                self.load_related_data(self.tree.selection()[0])

            tk.Button(form, text="Save", command=save_booking).grid(row=5, column=0, columnspan=2)

            # Populate the combobox with existing values
            self.update_booking_desc_combobox()

    def delete_booking(self):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "booking":
            messagebox.showwarning("Warning", "Please select a booking to delete")
            return
        booking_id = self.selected_related_id
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM booking WHERE id = ?", (booking_id,))
        conn.commit()
        conn.close()
        self.load_related_data(self.tree.selection()[0])

    def add_note(self, task_id):
        self.add_note_button.config(relief=tk.RIDGE)
        self.add_note_button.config(command=lambda: self.save_note_content(task_id, None, self.add_note_button, self.add_note, 'add'))

        if not task_id:
            messagebox.showwarning("Warning", "Please select a task to add a note")
            return

        self.load_note_and_open_text_editor(None)

    def edit_note(self, task_id):
        if not hasattr(self, 'selected_related_id') or self.selected_related_type != "note":
            messagebox.showwarning("Warning", "Please select a note to edit")
            return
        notes_id = self.selected_related_id
        self.edit_note_button.config(relief=tk.RIDGE)
        self.edit_note_button.config(command=lambda: self.save_note_content(task_id, notes_id, self.edit_note_button, self.edit_note, 'edit'))

        self.load_note_and_open_text_editor(notes_id)

    def load_note_and_open_text_editor(self, notes_id):
        # task_id = self.tree.selection()[0]
        note = None

        if notes_id:
            note = self.db.fetch_one("SELECT * FROM note WHERE note.id = ?", (notes_id,))

        # temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".md")
        self.temp_file_path = temp_file.name

        if note:
            with open(self.temp_file_path, "w") as file:
                file.write(note[2])
        else:
            open(self.temp_file_path, "w").close()

        # Open the default editor
        if os.name == 'nt':  # For Windows
            os.startfile(self.temp_file_path)
        else:  # For macOS and Linux
            subprocess.call(["xdg-open" if os.name != 'darwin' else "open", self.temp_file_path])

    def save_note_content(self, task_id, notes_id, button, cmd, behavior):
        if self.temp_file_path:
            with open(self.temp_file_path, "r") as file:
                content = file.read()

            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            if behavior == 'edit':
                cursor.execute("SELECT id, content FROM note WHERE note.id = ?", (notes_id,))
                note = cursor.fetchone()

                cursor.execute("UPDATE note SET content = ? WHERE id = ?", (content, note[0]))
            else:
                cursor.execute("INSERT INTO note (task_id, content) VALUES (?, ?)", (task_id, content))
            conn.commit()
            conn.close()

            os.remove(self.temp_file_path)
            self.temp_file_path = None
            self.load_related_data([task_id])

        button.config(command=lambda: cmd(task_id), relief=tk.RAISED)

    def delete_note(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a task to delete a note")
            return
        task_id = selected_items[0]
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM note WHERE task_id = ?", (task_id,))
        conn.commit()
        conn.close()
        self.load_related_data([task_id])

    def is_file_open(self, file_path):
        try:
            os.rename(file_path, file_path + "_renamed")
            time.sleep(0.2)
            os.rename(file_path + "_renamed", file_path)
            return False
        except OSError as e:
            # Log or print the error for diagnostics
            print(f"Error checking file status for {file_path}: {e}")
            return True  # File is likely open

    def search_data(self):
        table = self.table_var.get()
        field = self.field_var.get()
        operator = self.operator_var.get()
        value = self.value_entry.get()

        if not table or not field or not value:
            messagebox.showwarning("Warning", "Please enter all search criteria")
            return

        if operator == "LIKE":
            query = f"SELECT * FROM {table} WHERE {field} LIKE ?"
            value = '%' + value + '%'
        else:
            query = f"SELECT * FROM {table} WHERE {field} {operator} ?"

        results = self.db.fetch_all(query, (value,))

        if results:
            search_result_window = tk.Toplevel(self.root)
            search_result_window.title("Search Results")
            result_tree = ttk.Treeview(search_result_window, columns=[desc[0] for desc in cursor.description], show="headings")
            for col, desc in zip(result_tree["columns"], cursor.description):
                result_tree.heading(col, text=desc[0])
                result_tree.column(col, width=100)
            result_tree.pack(expand=True, fill=tk.BOTH)

            for row in results:
                result_tree.insert("", "end", values=row, tags=("clickable",))

            result_tree.tag_bind("clickable", "<ButtonRelease-1>", lambda event: self.on_search_result_click(event, result_tree, table))
        else:
            messagebox.showinfo("Info", "No results found")


    def on_search_result_click(self, event, result_tree, table):
        selected_item = result_tree.selection()[0]
        item_values = result_tree.item(selected_item, "values")
        related_id = item_values[0]  # Assuming the first column is the ID of the searched table

        if table == "task":
            task_id = related_id
        elif table == "delivery":
            task_id = self.db.fetch_one("SELECT task_id FROM task_delivery WHERE delivery_id = ?", (related_id,))[0]
        elif table == "link":
            task_id = self.db.fetch_one("SELECT task_id FROM task_link WHERE link_id = ?", (related_id,))[0]
        elif table == "tag":
            task_id = self.db.fetch_one("SELECT task_id FROM tag_task WHERE tag_id = ?", (related_id,))[0]
        elif table == "origin":
            task_id = self.db.fetch_one("SELECT task_id FROM task_origin WHERE origin_id = ?", (related_id,))[0]
        elif table == "booking":
            task_id = self.db.fetch_one("SELECT task_id FROM booking WHERE id = ?", (related_id,))[0]
        elif table == "note":
            task_id = self.db.fetch_one("SELECT task_id FROM note WHERE id = ?", (related_id,))[0]
        else:
            messagebox.showwarning("Warning", "Unsupported table for search results")
            return

        # Select the related task in the main tree
        self.tree.selection_set(task_id)
        self.tree.focus(task_id)
        self.tree.see(task_id)

        # Load related data for the selected task
        self.load_related_data([task_id])

        self.related_to_select_next = (table, related_id)


if __name__ == "__main__":
    root = tk.Tk()
    app = TaskManagerApp(root)
    root.mainloop()
