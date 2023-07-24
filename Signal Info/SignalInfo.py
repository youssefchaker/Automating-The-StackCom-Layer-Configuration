import tkinter as tk
from tkinter import ttk, filedialog
from lxml import etree

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("XDM Files", "*.xdm")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(tk.END, file_path)

def search_signal():
    file_path = file_entry.get()
    signal_name = signal_entry.get()

    if not file_path or not signal_name:
        result_label.config(text="Error: Please select a file and enter a signal name.")
        return

    if not file_path.lower().endswith(".xdm"):
        result_label.config(text="Error: Invalid file extension. Please select an XDM file.")
        return

    try:
        tree = etree.parse(file_path)
        root = tree.getroot()

        namespace = {'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}

        signals = root.xpath(f".//d:ctr[@name='{signal_name}']", namespaces=namespace)
        if not signals:
            result_label.config(text="Error: Signal not found.")
            return

        attributes = signals[0].xpath("d:var", namespaces=namespace)

        attribute_names = [attr.get("name") for attr in attributes]

        attribute_dropdown['values'] = attribute_names
        attribute_dropdown.current(0)
        attribute_dropdown['state'] = "readonly"
        result_label.config(text="")

        display_attribute(signal_name)  # Call display_attribute with signal_name

    except FileNotFoundError:
        result_label.config(text="Error: File not found.")
    except etree.XMLSyntaxError:
        result_label.config(text="Error: Invalid XML format.")

def display_attribute(signal_name):  # Accept signal_name as an argument
    selected_attribute = attribute_dropdown.get()

    result_text.delete("1.0", tk.END)

    file_path = file_entry.get()

    try:
        tree = etree.parse(file_path)
        root = tree.getroot()

        namespace = {'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}

        signals = root.xpath(f".//d:ctr[@name='{signal_name}']", namespaces=namespace)
        if not signals:
            result_label.config(text="Error: Signal not found.")
            return

        attributes = signals[0].xpath("d:var", namespaces=namespace)

        for attr in attributes:
            if attr.get("name") == selected_attribute:
                result_text.insert(tk.END, f"Attribute Name: {attr.get('name')}\n")
                result_text.insert(tk.END, f"Attribute Type: {attr.get('type')}\n")
                result_text.insert(tk.END, f"Attribute Value: {attr.get('value')}\n")
                result_text.insert(tk.END, f"\n")
                return

    except FileNotFoundError:
        result_label.config(text="Error: File not found.")
    except etree.XMLSyntaxError:
        result_label.config(text="Error: Invalid XML format.")

root = tk.Tk()
root.title("Signal Attribute Viewer")

# File Selection Section
file_label = tk.Label(root, text="Input File:")
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="E")

file_entry = tk.Entry(root, width=40)
file_entry.grid(row=0, column=1, padx=10, pady=10, sticky="W")

file_button = tk.Button(root, text="Choose", command=choose_file)
file_button.grid(row=0, column=2, padx=10, pady=10)

# Signal Search Section
signal_label = tk.Label(root, text="Signal Name:")
signal_label.grid(row=1, column=0, padx=10, pady=10, sticky="E")

signal_entry = tk.Entry(root, width=40)
signal_entry.grid(row=1, column=1, padx=10, pady=10, sticky="W")

search_button = tk.Button(root, text="Search", command=search_signal)
search_button.grid(row=1, column=2, padx=10, pady=10)

# Attribute Selection Section
attribute_label = tk.Label(root, text="Attribute:")
attribute_label.grid(row=2, column=0, padx=10, pady=10, sticky="E")

attribute_dropdown = ttk.Combobox(root, state="disabled")
attribute_dropdown.grid(row=2, column=1, padx=10, pady=10, sticky="W")

display_button = tk.Button(root, text="Display", command=lambda: display_attribute(signal_entry.get()))
display_button.grid(row=2, column=2, padx=10, pady=10)

# Result Section
result_label = tk.Label(root, fg="red")
result_label.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

result_text = tk.Text(root, width=50, height=10)
result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

root.resizable(False, False) 
root.mainloop()
