import os
import pyodbc as dbc
import datetime as dt
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

""" CHECK IF THE DATABASE FILE EXISTS, ELSE THE PROGRAM WILL NOT RUN """

db = None
# DIR_PATH IS USED DUE TO MY EDITOR REQUIRING ABSOLUTE PATH FOR FILE DIRECTORY
dir_path = os.path.dirname(os.path.realpath(__file__)) 

try:
    conn = dbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+dir_path+'/DB_Memory.accdb;')
    db = conn.cursor()
except Exception as e:
    messagebox.showerror("Database Error", f"Could not connect to database: {e}")
    exit()

""" BACKEND FUNCTIONALITIES """

rows = [] # TO SERVE AS A TEMPORARY MEMORY FOR ACCESSING ALL ENTRIES

def load_tasks():
    """LOAD ALL TASKS FROM DB_Memory.accdb FILE"""
    rows.clear()
    tasks_listbox.delete(0, tk.END)
    
    # FETCH ENTRIES FROM DB_Memory.accdb
    db.execute("SELECT * FROM List")
    table_rows = db.fetchall()
    
    # INSERT ALL ENTRIES INTO LISTBOX
    for per_row in table_rows:
        rows.append(per_row)
        select_id, task, status, reminder = per_row
        display_text = (
            f"{task} - {'Complete' if status == 'Complete' else 'Incomplete'}"
            f" (Reminder: {reminder})"
        )
        tasks_listbox.insert(tk.END, display_text)

def add_task():
    """ADD NEW TASKS TO THE TABLE List"""
    task = task_entry.get().strip()
    reminder = reminder_entry.get().strip()

    # ENSURE ENTRY FOR TASK NAME IS NOT EMPTY
    if task in (None, ''):
        messagebox.showwarning("Warning", "Task cannot be empty.")
        
    # ENSURE ENTRY FOR TASK REMINDER IS NOT EMPTY
    if reminder in (None, ''):
        messagebox.showwarning("Warning", "Reminder date and time cannot be empty.")
        
    # ENSURE ENTRY FOR TASK REMINDER FOLLOWS BY DATETIME FORMAT
    try:
        dt.datetime.strptime(reminder, "%Y-%m-%d %H:%M")
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
    """RETRIEVE ID OF THE SELECTED TASK FROM LISTBOX"""
    selected = tasks_listbox.curselection() # GET INDEX OF SELECTED TASK
    if selected in (None, ''):
        messagebox.showwarning("Warning", "No task selected.")
        return None

    selected_index = selected[0] # CLEAR TUPLE IN selected

    # ENSURE THE SELECTED TASK'S INDEX IS NOT OUT OF RANGE FROM rows
    if selected_index >= len(rows):
        messagebox.showerror("Error", "Task selection is out of sync. Please reload the tasks.")
        return None

    return rows[selected_index][0]  # RETURNS THE SELECTED TASK'S INDEX

def mark_complete():
    """MARKS SELECTED TASK AS COMPLETE"""
    task_id = get_selected_task_id()
    if task_id == None:
        # EXITS FUNCTION IF task_id RETURNS None
        return

    db.execute("UPDATE List SET Status = ? WHERE ID = ?", ('Complete', "{"+task_id+"}"))
    conn.commit()

    load_tasks()

def mark_incomplete():
    """MARKS SELECTED TASK AS INCOMPLETE"""
    task_id = get_selected_task_id()
    if task_id == None:
        # EXITS FUNCTION IF task_id RETURNS None
        return

    db.execute("UPDATE List SET Status = ? WHERE ID = ?", ('Incomplete', "{"+task_id+"}"))
    db.connection.commit()

    load_tasks()

def delete_task():
    """DELETES SELECTED TASK"""
    task_id = get_selected_task_id()
    if task_id == None:
        # EXITS FUNCTION IF task_id RETURNS None
        return

    db.execute("DELETE FROM List WHERE ID = ?", "{"+task_id+"}")
    conn.commit()
    
    load_tasks()

def confirm_pop(cmd, action):
    """GIVE AN ALERT FOR WHEN A BUTTON IS PRESSED"""
    result = messagebox.askyesno(f"{action.capitalize()}", f"Are you sure you want to {action} this entry?")
    if result:
        cmd()
        
def search():
    """FILTER DISPLAYED TASKS BASED ON CERTAIN PARTS OF ENTRIES"""
    category = category_var.get()
    
    # ADJUSTS KEYWORDS FILTER BASED ON SELECTED PART OF ENTRIES, BUT RETURNS ALL ENTRIES WHEN
    # INPUT ENTRY FROM search_entry DOES NOT HOLD ANY KEYWORD INPUT
    if search_entry.get().strip() in (None, ''):
        db.execute(f"SELECT * FROM List")
    elif category in ('Reminder'):
        db.execute(f"SELECT * FROM List WHERE {category} LIKE '%{search_entry.get().strip()}%'")
    elif category in ('Task', 'Status'):
        db.execute(f"SELECT * FROM List WHERE {category} LIKE '{search_entry.get().strip()}%'")
        
    table_rows = db.fetchall()
    tasks_listbox.delete(0, tk.END)
    
    # RE-DISPLAY TASKS IN LISTBOX
    for per_row in table_rows:
        select_id, task, status, reminder = per_row
        display_text = (
            f"{task} - {'Complete' if status == 'Complete' else 'Incomplete'}"
            f" (Reminder: {reminder})"
        )
        tasks_listbox.insert(tk.END, display_text)

""" FRONTEND DISPLAY """

root = tk.Tk()
img = tk.PhotoImage(file=dir_path+'/images/note-icon.png')
root.iconphoto(False, img)
root.title("To-Do List")
bgcolor = '#2F2F2F'
entrycolor = '#6F6F6F'
root.config(bg=bgcolor)

frame = tk.Frame(root, bg=bgcolor)
frame.pack(pady=10)

# UI FOR ADDING NEW TASKS
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

# UI FOR SEARCH FEATURE
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

# UI FOR DISPLAYING ALL TASKS
tasks_listbox = tk.Listbox(root, width=70, height=15, bg=entrycolor, fg='white', activestyle='none')
tasks_listbox.pack(pady=10)

# UI FOR OTHER BUTTON FUNCTIONALITIES
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

# LOAD ALL TASKS STORED IN DB_Memory.accdb FILE UPON STARTUP
load_tasks()

root.mainloop()

# CLOSES ALL CONNECTION WITH DB_Memory.accdb FILE
db.close()
conn.close()