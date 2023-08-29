import tkinter as tk
from tkinter import filedialog
from difflib import ndiff
import os

def choose_file1():
    file_path = filedialog.askopenfilename(filetypes=[("XDM Files", "*.xdm")])
    if file_path:
        file1_entry.delete(0, tk.END)
        file1_entry.insert(tk.END, file_path)

def choose_file2():
    file_path = filedialog.askopenfilename(filetypes=[("XDM Files", "*.xdm")])
    if file_path:
        file2_entry.delete(0, tk.END)
        file2_entry.insert(tk.END, file_path)

def compare_files():
    file1_path = file1_entry.get()
    file2_path = file2_entry.get()

    if not file1_path or not file2_path:
        diff_text.delete("1.0", tk.END)
        diff_text.insert(tk.END, "Error: Please select both files.")
        return

    if not file1_path.lower().endswith(".xdm") or not file2_path.lower().endswith(".xdm"):
        diff_text.delete("1.0", tk.END)
        diff_text.insert(tk.END, "Error: Invalid file extension.")
        return

    try:
        file1_name = os.path.basename(file1_path)
        file2_name = os.path.basename(file2_path)

        file1_folder = os.path.basename(os.path.dirname(file1_path))
        file2_folder = os.path.basename(os.path.dirname(file2_path))

        with open(file1_path, 'r') as file1, open(file2_path, 'r') as file2:
            diff = ndiff(file1.readlines(), file2.readlines())

        diff_text.delete("1.0", tk.END)

        for line in diff:
            if line.startswith('?'):
                continue
            elif line.startswith('+'):
                diff_text.insert(tk.END, f'[{file2_folder}] ({file2_name}) {line[1:]}', 'added')
            elif line.startswith('-'):
                diff_text.insert(tk.END, f'[{file1_folder}] ({file1_name}) {line[1:]}', 'removed')
            else:
                diff_text.insert(tk.END, line)

    except FileNotFoundError:
        diff_text.delete("1.0", tk.END)
        diff_text.insert(tk.END, "Error: File not found.")
    except IOError:
        diff_text.delete("1.0", tk.END)
        diff_text.insert(tk.END, "Error: Unable to open file.")

root = tk.Tk()
root.title("File Compare")

# File 1 Section
file1_label = tk.Label(root, text="File 1:")
file1_label.grid(row=0, column=0, padx=10, pady=10, sticky="E")

file1_entry = tk.Entry(root, width=40)
file1_entry.grid(row=0, column=1, padx=10, pady=10, sticky="W")

file1_button = tk.Button(root, text="Choose", command=choose_file1)
file1_button.grid(row=0, column=2, padx=10, pady=10)

# File 2 Section
file2_label = tk.Label(root, text="File 2:")
file2_label.grid(row=1, column=0, padx=10, pady=10, sticky="E")

file2_entry = tk.Entry(root, width=40)
file2_entry.grid(row=1, column=1, padx=10, pady=10, sticky="W")

file2_button = tk.Button(root, text="Choose", command=choose_file2)
file2_button.grid(row=1, column=2, padx=10, pady=10)

# Compare Button
compare_button = tk.Button(root, text="Compare", command=compare_files)
compare_button.grid(row=2, column=0, columnspan=3, pady=10)

# Create a scrollable text display area
scrollbar = tk.Scrollbar(root)
scrollbar.grid(row=3, column=2, sticky="NS")

diff_text = tk.Text(root, width=100, height=30, yscrollcommand=scrollbar.set)
diff_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

scrollbar.config(command=diff_text.yview)

# Define tag configuration for colors
diff_text.tag_configure('added', foreground='green')
diff_text.tag_configure('removed', foreground='red')
root.resizable(False, False) 
root.mainloop()
