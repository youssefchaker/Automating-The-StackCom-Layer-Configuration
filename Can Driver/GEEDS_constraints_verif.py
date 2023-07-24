import tkinter as tk
from tkinter import filedialog
from lxml import etree

def is_xdm_ordered(xdm_file):
    try:
        # Parse the XML file
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)

        # Get all the "ctr" elements in the .xdm file
        ctr_elements = root.xpath(".//d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        # Extract the names of frames and their CanObjectType
        frames_data = [(ctr, ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})) for ctr in ctr_elements]

        # Check if the frames are ordered as per the requirement
        receive_frames = [frame for frame, can_type in frames_data if can_type == "RECEIVE"]
        transmit_frames = [frame for frame, can_type in frames_data if can_type == "TRANSMIT"]

        if receive_frames and transmit_frames:
            # Compare the positions of the last RECEIVE frame and the first TRANSMIT frame
            last_receive_frame_index = ctr_elements.index(receive_frames[-1])
            first_transmit_frame_index = ctr_elements.index(transmit_frames[0])

            # Check if there are any RECEIVE frames after the first TRANSMIT frame
            if last_receive_frame_index > first_transmit_frame_index:
                return False

            return True

        else:
            return False

    except Exception as e:
        print(f"Error occurred while processing the .xdm file: {e}")
        return False

def browse_xdm():
    xdm_file_path = filedialog.askopenfilename(filetypes=[("XDM files", "*.xdm")])
    if not xdm_file_path:
        return
    xdm_file_entry.delete(0, tk.END)
    xdm_file_entry.insert(tk.END, xdm_file_path)

def check_ordered():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return

    is_ordered = is_xdm_ordered(xdm_file_path)

    if is_ordered:
        result_label.config(text="The .xdm file is ordered.", fg="green")
    else:
        result_label.config(text="The .xdm file is not ordered.", fg="red")

# Create the GUI
root = tk.Tk()
root.title("XDM File Order Checker")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

xdm_file_label = tk.Label(frame, text="Select XDM File:")
xdm_file_label.grid(row=0, column=0, padx=5, pady=5)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=0, column=1, padx=5, pady=5)

xdm_file_button = tk.Button(frame, text="Browse", command=browse_xdm)
xdm_file_button.grid(row=0, column=2, padx=5, pady=5)

check_button = tk.Button(frame, text="Check Ordered", command=check_ordered)
check_button.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

result_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
result_label.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
