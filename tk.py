


import tkinter as tk
from tkinter import filedialog

# create tkinter root window (it won't be shown)
root = tk.Tk()
root.withdraw()

# show file selection dialog for multiple files
selected_files = filedialog.askopenfilenames(title="Select DB files", filetypes=[("All Files", "*.xlsx")])

# check if user cancelled
if not selected_files:
    print("File selection cancelled.")
else:
    print(selected_files)