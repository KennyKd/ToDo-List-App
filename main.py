import os
import pyodbc as dbc
import datetime as dt
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

db = None
dir_path = os.path.dirname(os.path.realpath(__file__))
try:
    conn = dbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+dir_path+'/DB_Memory.accdb;')
    db = conn.cursor()
except Exception as e:
    messagebox.showerror("Database Error", f"Could not connect to database: {e}")
    exit()

selected_index = None
rows = []

def load_tasks():
    """Load tasks from the Excel file."""
    tasks_listbox.delete(0, tk.END) # Clear current tasks

    db.execute("SELECT * FROM List")
    table_rows = db.fetchall()
    for per_row in table_rows:
        rows.append(per_row)
        select_id, task, status, reminder = per_row
        display_text = (
            f"{task} - {'Complete' if status == 'Complete' else 'Incomplete'}"
            f" (Reminder: {reminder})"
        )
        tasks_listbox.insert(tk.END, display_text)

def add_task():
    """Add a new task to the to-do list."""
    task = task_entry.get().strip()
    reminder = reminder_entry.get().strip()

    if task in (None, ''):
        messagebox.showwarning("Warning", "Task cannot be empty.")

    if reminder in (None, ''):
        messagebox.showwarning("Warning", "Reminder date and time cannot be empty.")

    try:
        dt.datetime.strptime(reminder, "%Y-%m-%d %H:%M")  # Validate datetime format
    except ValueError:
        messagebox.showerror("Error", "Invalid date or time format.")
        return

    db.execute("""
            INSERT INTO List (Task, Status, Reminder)
            VALUES (?,?,?)
            """,
            (task, "Incomplete", reminder)
            )

    db.commit()
    
    load_tasks()
    task_entry.delete(0, tk.END)
    reminder_entry.delete(0, tk.END)

def get_selected_task_id():
    """Retrieve the ID of the selected task."""
    try:
        selected_index = tasks_listbox.curselection()[0]
        return rows[selected_index][0]  # Return the task ID
    except IndexError:
        messagebox.showwarning("Warning", "No task selected.")
        return None
    
def mark_complete():
    """Mark the selected task as complete."""
    task_id = get_selected_task_id()
    if not task_id:
        return

    db.execute("UPDATE List SET Status = ? WHERE ID = ?", ('Complete', "{"+task_id+"}"))
    conn.commit()

    load_tasks()
    rows = []

def mark_incomplete():
    """Mark the selected task as incomplete."""
    task_id = get_selected_task_id()
    if not task_id:
        return

    db.execute("UPDATE List SET Status = ? WHERE ID = ?", ('Incomplete', "{"+task_id+"}"))
    db.connection.commit()

    load_tasks()
    rows = []

def delete_task():
    """Delete the selected task."""
    task_id = get_selected_task_id()
    if not task_id:
        return

    db.execute("DELETE FROM List WHERE ID = ?", "{"+task_id+"}")
    conn.commit()
    
    load_tasks()
    rows = []

def confirm_pop(cmd, action):
    result = messagebox.askyesno(f"{action.capitalize()}", f"Are you sure you want to {action} this entry?")
    if result:
        cmd()
        
def search():
    category = category_var.get()
    if search_entry.get().strip() in (None, ''):
        db.execute(f"SELECT * FROM List")
    elif category in ('Reminder'):
        db.execute(f"SELECT * FROM List WHERE {category} LIKE '%{search_entry.get().strip()}%'")
    elif category in ('Task', 'Status'):
        db.execute(f"SELECT * FROM List WHERE {category} LIKE '{search_entry.get().strip()}%'")
        
    table_rows = db.fetchall()
    tasks_listbox.delete(0, tk.END)
    for per_row in table_rows:
        select_id, task, status, reminder = per_row
        display_text = (
            f"{task} - {'Complete' if status == 'Complete' else 'Incomplete'}"
            f" (Reminder: {reminder})"
        )
        tasks_listbox.insert(tk.END, display_text)

# GUI Setup
root = tk.Tk()
img = tk.PhotoImage(file=dir_path+'/images/note-icon.png')
root.iconphoto(False, img)
root.title("To-Do List")
bgcolor = '#2F2F2F'
entrycolor = '#6F6F6F'
root.config(bg=bgcolor)

frame = tk.Frame(root, bg=bgcolor)
frame.pack(pady=10)

task_label = tk.Label(frame, text="Task:", bg=bgcolor, fg='white')
task_label.grid(row=0, column=0, padx=5, sticky='e')

task_entry = tk.Entry(frame, width=30, bg=entrycolor, fg='white')
task_entry.grid(row=0, column=1, padx=5)

add_button = tk.Button(frame, text="Add Task", command=add_task, bg='#3CA64A', fg='white')
add_button.grid(row=0, column=4, rowspan=2, padx=5)

reminder_label = tk.Label(frame, text="Reminder (YYYY-MM-DD HH:MM):", bg=bgcolor, fg='white')
reminder_label.grid(row=1, column=0, padx=5, sticky='e')

reminder_entry = tk.Entry(frame, width=30, bg=entrycolor, fg='white')
reminder_entry.grid(row=1, column=1, padx=5)

category_var = tk.StringVar()
category_combobox = ttk.Combobox(frame, textvariable=category_var, state='readonly', width=20)
category_combobox['values'] = ("Task", "Status", "Reminder")
category_combobox.grid(row=2, column=0, padx=5, pady=5, sticky='e')
category_combobox.current(0)
style = ttk.Style()
style.map(category_combobox, selectbackground=[('readonly', 'red')])

search_entry = tk.Entry(frame, width=30, bg=entrycolor, fg='white')
search_entry.grid(row=2, column=1, padx=5, pady=5)

search_button = tk.Button(frame, text="Search", command=search, bg='#2562E6', fg='white')
search_button.grid(row=2, column=4, padx=5, pady=5)

tasks_listbox = tk.Listbox(root, width=70, height=15, bg=entrycolor, fg='white', activestyle='none')
tasks_listbox.pack(pady=10)

buttons_frame = tk.Frame(root, bg=bgcolor)
buttons_frame.pack(pady=10)

complete_button = tk.Button(buttons_frame, text="Mark Complete", command=mark_complete, bg='#216FED', fg='white')
complete_button.grid(row=0, column=0, padx=5)

incomplete_button = tk.Button(buttons_frame, text="Mark Incomplete", command=mark_incomplete, bg='#E0A500', fg='white')
incomplete_button.grid(row=0, column=1, padx=5)

delete_button = tk.Button(buttons_frame, text="Delete Task", command=lambda: confirm_pop(delete_task, "delete"), bg='#CF0000', fg='white')
delete_button.grid(row=0, column=2, padx=5)

exit_button = tk.Button(buttons_frame, text="Exit", command=root.quit, bg='#9E0505', fg='white')
exit_button.grid(row=0, column=3, padx=5)

# Load tasks on startup
load_tasks()

root.mainloop()

# Close the cursor and connection
db.close()
conn.close()