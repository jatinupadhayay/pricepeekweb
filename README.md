import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# Function to read Excel files and return DataFrame with normalized column names
def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        # Normalize column names: strip spaces and convert to lowercase
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading file: {e}")
        return None

# Function to load and display subject data
def load_subject_data():
    file_path = filedialog.askopenfilename(title="Select Subject Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global subjects_df
        subjects_df = read_excel_file(file_path)
        if subjects_df is not None:
            print("Subjects DataFrame loaded successfully.")
            display_subjects()

# Function to display subject data in Treeview
def display_subjects():
    subjects_tree.delete(*subjects_tree.get_children())  # Clear previous entries
    print("Displaying subjects...")
    for year in subjects_df.columns:
        for subject in subjects_df[year].dropna().tolist():
            subjects_tree.insert("", "end", values=(year.title(), subject))
    print("Subjects displayed.")

# Function to load and display room data
def load_room_data():
    file_path = filedialog.askopenfilename(title="Select Room Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global rooms_df
        rooms_df = read_excel_file(file_path)
        if rooms_df is not None:
            print("Rooms DataFrame loaded successfully.")
            display_rooms()

# Function to display room data in Treeview
def display_rooms():
    rooms_tree.delete(*rooms_tree.get_children())  # Clear previous entries
    print("Displaying rooms...")
    for _, row in rooms_df.iterrows():
        building = row.get('building', 'N/A')
        room_number = row.get('room number', 'N/A')
        capacity = row.get('capacity', 'N/A')
        rooms_tree.insert("", "end", values=(building.title(), room_number, capacity))
    print("Rooms displayed.")

# Function to load and display faculty data
def load_faculty_data():
    file_path = filedialog.askopenfilename(title="Select Faculty Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global faculty_df
        faculty_df = read_excel_file(file_path)
        if faculty_df is not None:
            print("Faculty DataFrame loaded successfully.")
            display_faculty()

# Function to display faculty data in Treeview
def display_faculty():
    faculty_tree.delete(*faculty_tree.get_children())  # Clear previous entries
    print("Displaying faculty...")
    
    # Normalize required columns
    required_columns = ['faculty name', 'occupation', 'experience']
    missing_columns = [col for col in required_columns if col not in faculty_df.columns]
    
    if missing_columns:
        messagebox.showerror("Error", f"Missing columns in faculty data: {', '.join(missing_columns)}")
        print(f"Missing columns: {missing_columns}")
        return
    
    for _, row in faculty_df.iterrows():
        faculty_name = row.get('faculty name', 'N/A')
        occupation = row.get('occupation', 'N/A')
        experience = row.get('experience', 'N/A')
        faculty_tree.insert("", "end", values=(faculty_name, occupation, experience))
    print("Faculty displayed.")

# Function to generate and display the timetable
def generate_timetable():
    if subjects_df is None or rooms_df is None or faculty_df is None:
        messagebox.showwarning("Warning", "Please upload subject, room, and faculty files first.")
        print("Timetable generation aborted: Missing data.")
        return
    
    print("Generating timetable...")
    
    try:
        # Create a new window for timetable display
        timetable_window = tk.Toplevel(root)
        timetable_window.title("Generated Timetable")
        timetable_window.geometry("1000x600")
    
        # Create Treeview for displaying timetable
        columns = ("Year", "Subject", "Room Number", "Building", "Date & Time", "Faculty")
        timetable_tree = ttk.Treeview(timetable_window, columns=columns, show='headings')
        for col in columns:
            timetable_tree.heading(col, text=col)
            timetable_tree.column(col, width=150, anchor='center')
        timetable_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
        # Combine subjects, rooms, and faculty into a timetable
        for year in subjects_df.columns:
            subjects_list = subjects_df[year].dropna().tolist()
            for i, subject in enumerate(subjects_list):
                room_index = i % len(rooms_df)  # Simple allocation logic
                faculty_index = i % len(faculty_df)  # Simple allocation logic
                
                room = rooms_df.iloc[room_index]
                faculty = faculty_df.iloc[faculty_index]
    
                # Fetch data with safe defaults
                building = room.get('building', 'N/A').title()
                room_number = room.get('room number', 'N/A')
                faculty_name = faculty.get('faculty name', 'N/A')
    
                # Ask for the date and time input for each subject
                date_time = simpledialog.askstring("Input", f"Enter Date & Time for '{subject}' ({year.title()}):", parent=timetable_window)
                if not date_time:
                    date_time = "Not Assigned"
    
                # Insert the timetable data into the Treeview
                timetable_tree.insert("", "end", values=(year.title(), subject, room_number, building, date_time, faculty_name))
        
        print("Timetable generated and displayed successfully.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the timetable: {e}")
        print(f"Error during timetable generation: {e}")

# Initialize main application window
root = tk.Tk()
root.title("Subject, Room, and Faculty Details Viewer")
root.geometry("800x700")  # Increased size to accommodate all elements

# Create a frame for better layout management
frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# UI for uploading subject data
load_subject_button = tk.Button(frame, text="Upload Subject Excel File", command=load_subject_data, width=30)
load_subject_button.grid(row=0, column=0, pady=5, padx=5)

# Treeview for displaying subjects
subjects_tree = ttk.Treeview(frame, columns=("Year", "Subject"), show='headings', height=5)
subjects_tree.heading("Year", text="Year")
subjects_tree.heading("Subject", text="Subject")
subjects_tree.column("Year", width=100, anchor='center')
subjects_tree.column("Subject", width=200, anchor='w')
subjects_tree.grid(row=1, column=0, padx=5, pady=5, sticky='nsew')

# UI for uploading room data
load_room_button = tk.Button(frame, text="Upload Room Excel File", command=load_room_data, width=30)
load_room_button.grid(row=2, column=0, pady=5, padx=5)

# Treeview for displaying rooms
rooms_tree = ttk.Treeview(frame, columns=("Building", "Room Number", "Capacity"), show='headings', height=5)
rooms_tree.heading("Building", text="Building")
rooms_tree.heading("Room Number", text="Room Number")
rooms_tree.heading("Capacity", text="Capacity")
rooms_tree.column("Building", width=100, anchor='center')
rooms_tree.column("Room Number", width=100, anchor='center')
rooms_tree.column("Capacity", width=100, anchor='center')
rooms_tree.grid(row=3, column=0, padx=5, pady=5, sticky='nsew')

# UI for uploading faculty data
load_faculty_button = tk.Button(frame, text="Upload Faculty Excel File", command=load_faculty_data, width=30)
load_faculty_button.grid(row=4, column=0, pady=5, padx=5)

# Treeview for displaying faculties
faculty_tree = ttk.Treeview(frame, columns=("Faculty Name", "Occupation", "Experience"), show='headings', height=5)
faculty_tree.heading("Faculty Name", text="Faculty Name")
faculty_tree.heading("Occupation", text="Occupation")
faculty_tree.heading("Experience", text="Experience")
faculty_tree.column("Faculty Name", width=150, anchor='w')
faculty_tree.column("Occupation", width=100, anchor='center')
faculty_tree.column("Experience", width=100, anchor='center')
faculty_tree.grid(row=5, column=0, padx=5, pady=5, sticky='nsew')

# Configure grid weights for proper resizing
frame.rowconfigure(1, weight=1)
frame.rowconfigure(3, weight=1)
frame.rowconfigure(5, weight=1)
frame.columnconfigure(0, weight=1)

# Button to generate and display the timetable
generate_timetable_button = tk.Button(root, text="Generate Timetable", command=generate_timetable, bg="green", fg="white", font=("Helvetica", 12, "bold"))
generate_timetable_button.pack(pady=10, padx=10, fill=tk.X)

# Initialize data variables
subjects_df = None
rooms_df = None
faculty_df = None

# Run the main application loop
root.mainloop()
