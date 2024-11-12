# --- Imports ---
import os
import sys
import tkinter as tk
from tkinter import messagebox, Toplevel, simpledialog
from tkinter import ttk
from datetime import datetime
from tksheet import Sheet
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
from matplotlib.figure import Figure
from mpl_toolkits.mplot3d import Axes3D
import keyboard
import time

# --- Constants and Configuration ---
# Determine the directory of the running executable/script
if getattr(sys, 'frozen', False):
    current_directory = os.path.dirname(sys.executable)
else:
    current_directory = os.path.dirname(os.path.abspath(__file__))

# Define the base path for accessing resources
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # This is the temporary folder where PyInstaller unpacks files
else:
    base_path = current_directory  # Use the current directory when not frozen

# Define the SQLite database file
DB_FILE = os.path.join(current_directory, 'warehouse_inventory.db')

# Define paths for images
crane1_image_path = os.path.join(base_path, 'crane1_image.png')
crane2_image_path = os.path.join(base_path, 'crane2_image.png')

# Warehouse layout constants
VALID_COLUMNS = [1, 2, 3, 4]
VALID_BAYS = [1, 2, 3, 4, 5, 6, 7, 8]
VALID_LEVELS = [1, 2, 3, 4, 5, 6, 7, 8]

BLOCKED_POSITIONS_CRANE_1 = [
    (3, 1, 1), (3, 2, 1), (3, 1, 2), (3, 2, 2),
    (1, 1, 1), (1, 2, 1), (1, 1, 2), (1, 2, 2),
    (4, 1, 1), (4, 2, 1), (4, 1, 2), (4, 2, 2),
    (2, 1, 1), (2, 2, 1), (2, 1, 2), (2, 2, 2)
]

BLOCKED_POSITIONS_CRANE_2 = [
    (3, 1, 1), (3, 2, 1), (3, 1, 2), (3, 2, 2),
    (1, 1, 1), (1, 2, 1), (1, 1, 2), (1, 2, 2),
    (4, 1, 1), (4, 2, 1), (4, 1, 2), (4, 2, 2),
    (2, 1, 1), (2, 2, 1), (2, 1, 2), (2, 2, 2)
]

# --- Database Functions ---
def create_database():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS warehouse_inventory (
        material_id TEXT,
        description TEXT,
        quantity INTEGER,
        unit TEXT,
        batch TEXT,
        handling_unit TEXT,
        stock_type TEXT,
        sap_bin TEXT,
        date_in TEXT,
        date_out TEXT,
        column INTEGER,
        bay INTEGER,
        level INTEGER,
        crane TEXT,
        status TEXT,
        destination TEXT
    )
    ''')
    
    # Add handling_unit column if it doesn't exist
    try:
        cursor.execute("ALTER TABLE warehouse_inventory ADD COLUMN handling_unit TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists
        
    conn.commit()
    conn.close()

# --- Material Management Functions ---
def store_material():
    try:
        # Get values from entries
        material_id = entry_material_id.get()
        description = entry_description.get()
        quantity = int(entry_quantity.get())
        unit = entry_unit.get()
        batch = entry_batch.get()
        handling_unit = entry_handling_unit.get()
        stock_type = entry_stock_type.get()
        sap_bin = entry_sap_bin.get()
        column = int(entry_column.get())
        bay = int(entry_bay.get())
        level = int(entry_level.get())
        crane = selected_crane.get()
        
        # Validate inputs
        if not all([material_id, description, quantity, unit, batch, handling_unit, stock_type, sap_bin]):
            messagebox.showerror("Error", "All fields must be filled")
            return
            
        if column not in VALID_COLUMNS or bay not in VALID_BAYS or level not in VALID_LEVELS:
            messagebox.showerror("Error", "Invalid position")
            return
            
        # Check if position is blocked for the specific crane
        position = (column, bay, level)
        blocked_positions = BLOCKED_POSITIONS_CRANE_1 if crane == "Crane 1" else BLOCKED_POSITIONS_CRANE_2
        if position in blocked_positions:
            messagebox.showerror("Error", "This position is blocked")
            return

        # Check if position is occupied for this specific crane
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
        SELECT * FROM warehouse_inventory 
        WHERE column = ? 
        AND bay = ? 
        AND level = ? 
        AND crane = ?
        AND status = 'Stored'
        ''', (column, bay, level, crane))
        
        if cursor.fetchone():
            messagebox.showerror("Error", f"This position is already occupied for {crane}")
            conn.close()
            return
            
        # Check for height limitation at level 4
        if level == 4:
            messagebox.showinfo("Height Limitation", "Level 4 has a height limitation of 120 cm at maximum.")
        
        # Modified confirmation message
        if messagebox.askyesno("Confirm Position", 
                             f"Storing in position Column {column}, Bay {bay}, Level {level} with {crane}\n" +
                             "Do you want to continue?"):
            # Store in database
            cursor.execute('''
            INSERT INTO warehouse_inventory 
            (material_id, description, quantity, unit, batch, handling_unit, stock_type, sap_bin, 
             date_in, column, bay, level, crane, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (material_id, description, quantity, unit, batch, handling_unit, stock_type, sap_bin,
                  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  column, bay, level, crane, "Stored"))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", "Material stored successfully")
            clear_entries()
            update_sheet()
            
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {str(e)}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def retrieve_material():
    try:
        material_id = entry_material_id.get()
        batch = entry_batch.get()
        column = entry_column.get()
        bay = entry_bay.get()
        level = entry_level.get()
        crane = selected_crane.get()  # Get the selected crane
        
        # Verify all required fields are filled
        if not all([material_id, batch, column, bay, level]):
            messagebox.showerror("Error", "Please fill in Material ID, Batch, Column, Bay, and Level")
            return
        
        # Get the sheet data
        data = sheet.get_sheet_data()
        
        # Find the exact match for position (including crane), material ID, and batch
        found = False
        for row_index, row in enumerate(data):
            if (row[0] == material_id and      # Material ID match
                row[4] == batch and            # Batch match
                row[10] == column and          # Column match
                row[11] == bay and             # Bay match
                row[12] == level and           # Level match
                row[13] == crane and           # Crane match
                row[14] == "Stored"):          # Status check
                
                # Found the exact pallet we want to retrieve
                found = True
                # Update the status and date_out
                current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                sheet.set_cell_data(row_index, 9, current_date)  # date_out
                sheet.set_cell_data(row_index, 14, "Retrieved")  # status
                
                # Update database and save changes
                update_database()
                messagebox.showinfo("Success", "Material retrieved successfully")
                break
        
        if not found:
            messagebox.showerror("Error", "No matching material found at specified position")
            
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def store_material_wrapper():
    store_material()

def retrieve_material_wrapper():
    retrieve_material()

def search_material():
    search_window = Toplevel(root)
    search_window.title("Search Material")
    search_window.geometry("300x250")

    # Input fields
    tk.Label(search_window, text="Search by Material ID:").pack(pady=5)
    entry_material_id = tk.Entry(search_window)
    entry_material_id.pack(pady=5)

    tk.Label(search_window, text="Search by Batch:").pack(pady=5)
    entry_batch = tk.Entry(search_window)
    entry_batch.pack(pady=5)

    def perform_search():
        material_id = entry_material_id.get()
        batch = entry_batch.get()

        # Modified query to include batch
        query = """
        SELECT * FROM warehouse_inventory 
        WHERE material_id LIKE ? 
        AND batch LIKE ? 
        AND status = 'Stored'
        """
        params = [f'%{material_id}%', f'%{batch}%']

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()

        if rows:
            show_search_results(rows)
        else:
            messagebox.showinfo("Search Results", "No stored materials found.")

    # Search button
    tk.Button(search_window, text="Search", command=perform_search).pack(pady=20)

    # Clear button to close the search window
    tk.Button(search_window, text="Close", command=search_window.destroy).pack(pady=5)

def show_search_results(results):
    search_window = Toplevel()
    search_window.title("Search Results")
    search_window.geometry("800x400")
    
    # Create Treeview
    tree = ttk.Treeview(search_window)
    tree["columns"] = ("ID", "Description", "Quantity", "Location", "Status", "Crane")
    
    # Format columns
    tree.column("#0", width=0, stretch=tk.NO)
    tree.column("ID", width=100)
    tree.column("Description", width=200)
    tree.column("Quantity", width=100)
    tree.column("Location", width=200)
    tree.column("Status", width=100)
    tree.column("Crane", width=100)  # Add column for Crane
    
    # Create headings
    tree.heading("ID", text="Material ID")
    tree.heading("Description", text="Description")
    tree.heading("Quantity", text="Quantity")
    tree.heading("Location", text="Location")
    tree.heading("Status", text="Status")
    tree.heading("Crane", text="Crane")  # Heading for Crane
    
    # Add data
    for item in results:
        location = f"Column {item[9]}, Bay {item[10]}, Level {item[11]}"  # Adjust indices based on your schema
        tree.insert("", tk.END, values=(item[0], item[1], item[2], location, item[13], item[12]))  # item[12] is the crane
    
    tree.pack(expand=True, fill="both")

# --- Utility Functions ---
def clear_entries():
    for entry in [entry_material_id, entry_description, entry_quantity, entry_unit,
                 entry_batch, entry_handling_unit, entry_stock_type, entry_sap_bin,
                 entry_column, entry_bay, entry_level]:
        entry.delete(0, tk.END)

def update_sheet():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM warehouse_inventory')
    rows = cursor.fetchall()
    conn.close()

    sheet.set_sheet_data([list(row) for row in rows])
    
    for row_index, row in enumerate(rows):
        if row[13] == "Retrieved":
            sheet.highlight_rows(row_index, bg="#FFCCCC")
    
    # Update 3D views
    for widget in crane1_tab.winfo_children():
        widget.destroy()
    for widget in crane2_tab.winfo_children():
        widget.destroy()
    create_3d_plot(crane1_tab, 'Crane 1')
    create_3d_plot(crane2_tab, 'Crane 2')

def update_crane_image(*args):
    if not hasattr(root, 'winfo_width'):
        return
        
    frame_width = image_frame.winfo_width()
    frame_height = image_frame.winfo_height()
    
    if frame_width <= 1 or frame_height <= 1:
        return
    
    current_crane = selected_crane.get()
    image_filename = 'crane1_image.png' if current_crane == "Crane 1" else 'crane2_image.png'
    crane_image_path = os.path.join(current_directory, image_filename)
    
    if os.path.exists(crane_image_path):
        try:
            pil_image = Image.open(crane_image_path)
            desired_width = max(int(frame_width * 0.95), 800)
            desired_height = max(int(frame_height * 0.95), 1000)
            
            original_width, original_height = pil_image.size
            aspect_ratio = original_width / original_height
            
            if desired_width / desired_height > aspect_ratio:
                new_width = int(desired_height * aspect_ratio)
                new_height = desired_height
            else:
                new_width = desired_width
                new_height = int(desired_width / aspect_ratio)
            
            resized_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            photo_image = ImageTk.PhotoImage(resized_image)
            
            image_label.config(image=photo_image)
            image_label.image = photo_image
            
        except Exception as e:
            print(f"Error resizing image: {e}")
    else:
        print(f"Image not found: {crane_image_path}")

def export_to_excel():
    try:
        conn = sqlite3.connect(DB_FILE)
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Query for today's activities (both storing and retrieving)
        query = '''
        SELECT * FROM warehouse_inventory 
        WHERE date(date_in) = ? 
        OR date(date_out) = ?
        '''
        
        df = pd.read_sql_query(query, conn, params=[today, today])
        
        if df.empty:
            messagebox.showinfo("Export", "No activities found for today")
            return
            
        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(current_directory, f'warehouse_daily_report_{timestamp}.xlsx')
        
        # Create Excel writer object
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        
        # Write data to Excel
        df.to_excel(writer, sheet_name='Daily Activities', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Daily Activities']
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1
        })
        
        # Format the header row
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
            
        writer.close()
        messagebox.showinfo("Export", f"Daily report exported to {filename}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Export failed: {str(e)}")
    finally:
        conn.close()

def create_3d_plot(parent, view_type):
    fig = Figure(figsize=(8, 6))
    ax = fig.add_subplot(111, projection='3d')
    
    # Create the basic warehouse structure with rearranged columns
    X, Y, Z = [], [], []
    column_mapping = {
        1: 1,  # Column 1 stays at position 1
        2: 3,  # Column 2 moves to position 3
        3: 0,  # Column 3 moves to position 0
        4: 4   # Column 4 stays at position 4
    }
    
    for col in VALID_COLUMNS:
        mapped_x = column_mapping[col]
        for bay in VALID_BAYS:
            for level in VALID_LEVELS:
                X.append(mapped_x)
                Y.append(bay)
                Z.append(level)
    
    # Plot the structure
    ax.scatter(X, Y, Z, c='blue', marker='s', alpha=0.3)
    
    # Add crane position (red dashed diagonal line)
    # The crane is at column 2 (mapped_x = 2), spanning from bay 1, level 1 to bay 8, level 8
    crane_x = 2  # Mapped position for "Crane"
    ax.plot([crane_x, crane_x], [1, 8], [1, 8], 'r--', linewidth=2, label='Crane Position')
    
    # Plot blocked positions with mapped coordinates
    blocked_positions = BLOCKED_POSITIONS_CRANE_1 if view_type == 'Crane 1' else BLOCKED_POSITIONS_CRANE_2
    X_blocked = [column_mapping[pos[0]] for pos in blocked_positions]
    Y_blocked = [pos[1] for pos in blocked_positions]
    Z_blocked = [pos[2] for pos in blocked_positions]
    ax.scatter(X_blocked, Y_blocked, Z_blocked, c='red', marker='x', s=100, label='Blocked')
    
    # Get occupied positions from database and map their coordinates
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT column, bay, level FROM warehouse_inventory WHERE status = "Stored" AND crane = ?', (view_type,))
    occupied = cursor.fetchall()
    conn.close()
    
    if occupied:
        X_occ = [column_mapping[pos[0]] for pos in occupied]
        Y_occ = [pos[1] for pos in occupied]
        Z_occ = [pos[2] for pos in occupied]
        ax.scatter(X_occ, Y_occ, Z_occ, c='green', marker='o', s=100, label='Occupied')
    
    # Customize the plot
    ax.set_xlabel('Columns')
    ax.set_ylabel('Bays')
    ax.set_zlabel('Levels')
    ax.set_title(f'{view_type} Warehouse View')
    
    # Set custom x-axis ticks to show actual column numbers
    ax.set_xticks([0, 1, 2, 3, 4])
    ax.set_xticklabels(['3', '1', 'Crane', '2', '4'])
    
    # Adjust the view angle for better visualization
    ax.view_init(elev=20, azim=45)
    
    # Add legend
    ax.legend()
    
    canvas = FigureCanvasTkAgg(fig, parent)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

def create_entries(parent_frame, *entry_widgets):
    entry_labels = ["Material ID", "Description/Notes", "Quantity", "Unit of Measure", 
                   "Batch", "Handling Unit", "Stock Type", "SAP Bin", "Column", "Bay", "Level"]
    
    parent_frame.grid_columnconfigure(1, weight=1)
    
    for i, (label_text, entry_widget) in enumerate(zip(entry_labels, entry_widgets)):
        tk.Label(parent_frame, text=label_text).grid(row=i, column=0, sticky="e", padx=5, pady=2)
        entry_widget.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
        entry_widget.config(width=20)

def create_buttons(parent_frame, store_func, retrieve_func, clear_func, 
                  update_func, export_func, search_func):
    button_configs = [
        ("Store Material", store_func),
        ("Retrieve Material", retrieve_func),
        ("Clear Entries", clear_func),
        ("Refresh View", update_func),
        ("Export Daily Data", export_func),
        ("Export Complete Log", export_complete_log),
        ("Search Material", search_func),
        ("Reverse Last Entry", reverse_last_entry),
        ("Undo Last Operation", undo_last_operation)
    ]
    
    parent_frame.grid_columnconfigure(0, weight=1)
    
    for i, (text, command) in enumerate(button_configs):
        tk.Button(parent_frame, text=text, command=command, width=15).grid(
            row=i, column=0, pady=5, padx=10, sticky="ew")

def export_complete_log():
    try:
        conn = sqlite3.connect(DB_FILE)
        
        # Get all data from database, including both stored and retrieved materials
        query = '''
        SELECT * FROM warehouse_inventory 
        ORDER BY 
            CASE 
                WHEN date_out IS NULL THEN date_in 
                ELSE date_out 
            END DESC
        '''
        df = pd.read_sql_query(query, conn)
        
        if df.empty:
            messagebox.showinfo("Export", "No data in database")
            return
            
        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(current_directory, f'warehouse_complete_log_{timestamp}.xlsx')
        
        # Create Excel writer object
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        
        # Write data to Excel
        df.to_excel(writer, sheet_name='Complete Log', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Complete Log']
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1
        })
        
        # Format the header row
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
            
        writer.close()
        messagebox.showinfo("Export", f"Complete log exported to {filename}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Export failed: {str(e)}")
    finally:
        conn.close()

# Add at the top with other global variables
operation_history = []  # Will store the last few operations for undo purposes

def reverse_last_entry():
    if not messagebox.askyesno("Confirm Reverse", "Are you sure you want to reverse the last entry?"):
        return
        
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Get the last entry from the database
        query = '''
        SELECT * FROM warehouse_inventory 
        ORDER BY 
            CASE 
                WHEN date_out IS NULL THEN date_in 
                ELSE date_out 
            END DESC 
        LIMIT 1
        '''
        
        cursor.execute(query)
        last_entry = cursor.fetchone()
        
        if not last_entry:
            messagebox.showinfo("Reverse Entry", "No entries to reverse")
            return
            
        # Store the entry and operation type for potential undo
        operation_data = {
            'entry': last_entry,
            'operation': 'reverse',
            'original_status': 'Stored' if last_entry[9] is None else 'Retrieved'
        }
            
        # Check if it was a storage or retrieval operation
        if last_entry[9] is None:  # date_out is NULL, meaning it was a storage operation
            # Delete the last storage entry
            delete_query = '''
            DELETE FROM warehouse_inventory 
            WHERE material_id = ? 
            AND batch = ? 
            AND column = ? 
            AND bay = ? 
            AND level = ? 
            AND crane = ?
            AND date_in = ?
            '''
            cursor.execute(delete_query, (
                last_entry[0],  # material_id
                last_entry[4],  # batch
                last_entry[10], # column
                last_entry[11], # bay
                last_entry[12], # level
                last_entry[13], # crane
                last_entry[8]   # date_in
            ))
            operation_type = "Storage"
            
        else:  # It was a retrieval operation
            # Update the entry to mark it as stored again
            update_query = '''
            UPDATE warehouse_inventory 
            SET date_out = NULL, status = 'Stored'
            WHERE material_id = ? 
            AND batch = ? 
            AND column = ? 
            AND bay = ? 
            AND level = ? 
            AND crane = ?
            AND date_out = ?
            '''
            cursor.execute(update_query, (
                last_entry[0],  # material_id
                last_entry[4],  # batch
                last_entry[10], # column
                last_entry[11], # bay
                last_entry[12], # level
                last_entry[13], # crane
                last_entry[9]   # date_out
            ))
            operation_type = "Retrieval"
            
        conn.commit()
        
        # Add to operation history
        operation_history.append(operation_data)
        
        messagebox.showinfo("Success", f"Last {operation_type} operation has been reversed")
        
        # Update the view
        update_sheet()
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to reverse last entry: {str(e)}")
    finally:
        conn.close()

def undo_last_operation():
    if not operation_history:
        messagebox.showinfo("Undo", "No operations to undo")
        return
        
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        last_operation = operation_history.pop()
        entry = last_operation['entry']
        
        if last_operation['original_status'] == 'Stored':
            # Re-insert the deleted entry
            insert_query = '''
            INSERT INTO warehouse_inventory 
            (material_id, description, quantity, unit, batch, handling_unit, 
             stock_type, sap_bin, date_in, date_out, column, bay, level, 
             crane, status, destination)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            '''
            cursor.execute(insert_query, entry)
            
        else:  # Was Retrieved
            # Update back to retrieved status
            update_query = '''
            UPDATE warehouse_inventory 
            SET date_out = ?, status = 'Retrieved'
            WHERE material_id = ? 
            AND batch = ? 
            AND column = ? 
            AND bay = ? 
            AND level = ? 
            AND crane = ?
            '''
            cursor.execute(update_query, (
                entry[9],      # original date_out
                entry[0],      # material_id
                entry[4],      # batch
                entry[10],     # column
                entry[11],     # bay
                entry[12],     # level
                entry[13]      # crane
            ))
            
        conn.commit()
        messagebox.showinfo("Success", "Last operation has been undone")
        
        # Update the view
        update_sheet()
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to undo operation: {str(e)}")
    finally:
        conn.close()

# --- Main Application ---
if __name__ == "__main__":
    # Initialize database
    create_database()

    # Create main window
    root = tk.Tk()
    root.title("Warehouse Inventory Management")
    root.minsize(1200, 800)

    # Add after the root = tk.Tk() line
    default_font = ('Arial', 12)  # or any other size/font you prefer
    root.option_add('*Font', default_font)

    # Create Tkinter variables
    selected_crane = tk.StringVar(value="Crane 1")
    destination_var = tk.StringVar(value="Into Production")

    # Create notebook for tabs
    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    # Create main interface tab
    main_frame = tk.Frame(notebook)
    notebook.add(main_frame, text='Main Interface')

    # Create crane 1 3D view tab
    crane1_tab = tk.Frame(notebook)
    notebook.add(crane1_tab, text='Crane 1 - 3D View')

    # Create crane 2 3D view tab
    crane2_tab = tk.Frame(notebook)
    notebook.add(crane2_tab, text='Crane 2 - 3D View')

    # Create sub-frames in main interface
    entries_frame = tk.Frame(main_frame, relief="groove", borderwidth=1)
    image_frame = tk.Frame(main_frame, relief="groove", borderwidth=1)
    buttons_frame = tk.Frame(main_frame, relief="groove", borderwidth=1)
    sheet_frame = tk.Frame(main_frame, relief="groove", borderwidth=1)

    # Create entry widgets
    entry_material_id = tk.Entry(entries_frame)
    entry_description = tk.Entry(entries_frame)
    entry_quantity = tk.Entry(entries_frame)
    entry_unit = tk.Entry(entries_frame)
    entry_batch = tk.Entry(entries_frame)
    entry_handling_unit = tk.Entry(entries_frame)
    entry_stock_type = tk.Entry(entries_frame)
    entry_sap_bin = tk.Entry(entries_frame)
    entry_column = tk.Entry(entries_frame)
    entry_bay = tk.Entry(entries_frame)
    entry_level = tk.Entry(entries_frame)

    # Create image label in image frame
    image_label = tk.Label(image_frame)
    image_label.grid(row=0, column=0, sticky="nsew")
    image_frame.grid_rowconfigure(0, weight=1)
    image_frame.grid_columnconfigure(0, weight=1)

    # Create 3D views in respective tabs
    create_3d_plot(crane1_tab, 'Crane 1')
    create_3d_plot(crane2_tab, 'Crane 2')

    # Configure main frame grid
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=3)
    main_frame.grid_columnconfigure(2, weight=1)
    for i in range(13):
        main_frame.grid_rowconfigure(i, weight=1)
    main_frame.grid_rowconfigure(12, weight=2)

    # Place frames in main interface
    entries_frame.grid(row=0, column=0, rowspan=12, sticky="nsew", padx=5, pady=5)
    image_frame.grid(row=0, column=1, rowspan=12, sticky="nsew", padx=5, pady=5)
    buttons_frame.grid(row=0, column=2, rowspan=12, sticky="nsew", padx=5, pady=5)
    sheet_frame.grid(row=12, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

    # Create entries and buttons
    create_entries(entries_frame, 
                  entry_material_id, 
                  entry_description, 
                  entry_quantity,
                  entry_unit, 
                  entry_batch, 
                  entry_handling_unit,
                  entry_stock_type, 
                  entry_sap_bin,
                  entry_column, 
                  entry_bay, 
                  entry_level)

    # Add some spacing
    tk.Label(entries_frame, text="").grid(row=12, column=0, columnspan=2, pady=10)  # Empty row for spacing

    # Add crane selection lower
    crane_frame = tk.Frame(entries_frame)
    crane_frame.grid(row=13, column=0, columnspan=2, pady=5)
    tk.Label(crane_frame, text="Crane Selection:").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(crane_frame, text="Crane 1", variable=selected_crane, 
                   value="Crane 1", command=update_crane_image).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(crane_frame, text="Crane 2", variable=selected_crane, 
                   value="Crane 2", command=update_crane_image).pack(side=tk.LEFT, padx=5)

    # Add another spacing
    tk.Label(entries_frame, text="").grid(row=14, column=0, columnspan=2, pady=5)  # Empty row for spacing

    # Add destination selection even lower
    destination_frame = tk.Frame(entries_frame)
    destination_frame.grid(row=15, column=0, columnspan=2, pady=5)
    tk.Label(destination_frame, text="Destination:").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(destination_frame, text="Into Production", 
                   variable=destination_var,  # Use the renamed variable
                   value="Into Production").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(destination_frame, text="Back to Warehouse", 
                   variable=destination_var,  # Use the renamed variable
                   value="Back to Warehouse").pack(side=tk.LEFT, padx=5)

    create_buttons(buttons_frame, 
        store_material_wrapper,
        retrieve_material_wrapper,
        clear_entries, 
        update_sheet, 
        export_to_excel,
        search_material
    )

    # Create sheet
    sheet = Sheet(sheet_frame)
    sheet.enable_bindings()
    sheet.grid(row=0, column=0, sticky="nsew")
    sheet_frame.grid_rowconfigure(0, weight=1)
    sheet_frame.grid_columnconfigure(0, weight=1)

    # Set up sheet headers
    sheet.headers([
        "Material ID", "Description", "Quantity", "Unit", "Batch", 
        "Stock Type", "SAP Bin", "Handling Unit", "Date In", "Date Out", 
        "Column", "Bay", "Level", "Crane", "Status", "Destination"
    ])

    # Initial setup
    update_crane_image()
    update_sheet()

    # Add to imports
    import keyboard

    # Add this global variable at the start of your script
    scanned_data = ""

    class ScannerState:
        def __init__(self):
            self.data = ""
            self.last_scan_time = 0
            self.timeout = 0.5
            self.scanning = False

    scanner_state = ScannerState()

    def handle_scanner_input(event):
        """
        Gets the currently focused widget and inserts the scanned data into it
        """
        global scanner_state
        
        current_time = time.time()
        focused_widget = root.focus_get()
        
        if isinstance(focused_widget, tk.Entry):
            # Reset if timeout occurred
            if current_time - scanner_state.last_scan_time > scanner_state.timeout:
                scanner_state.data = ""
                scanner_state.scanning = False
                
            scanner_state.last_scan_time = current_time
            
            if event.keysym == 'Return':
                focused_widget.delete(0, tk.END)
                focused_widget.insert(0, scanner_state.data.replace('+', '-'))
                scanner_state.data = ""
                scanner_state.scanning = False
            elif event.keysym.isprintable():
                if not scanner_state.scanning:
                    scanner_state.data = event.char
                    scanner_state.scanning = True
                else:
                    scanner_state.data += event.char

    # Add near the end of the main block, before root.mainloop()
    # Bind both key press and return events
    root.bind('<Key>', handle_scanner_input)

    # Start the application
    root.mainloop()