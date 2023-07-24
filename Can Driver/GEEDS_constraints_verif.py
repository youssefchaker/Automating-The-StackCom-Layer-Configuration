import tkinter as tk
from tkinter import filedialog
from lxml import etree

def ordered_by_RX_TX(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)

        # Get all "d:ctr" elements inside "d:lst" with name="CanHardwareObject"
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        # Find the last RECEIVE frame
        last_receive_frame_index = None
        for i, ctr in enumerate(ctr_elements):
            if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "RECEIVE":
                last_receive_frame_index = i

        # Find the first TRANSMIT frame
        first_transmit_frame_index = None
        for i, ctr in enumerate(ctr_elements):
            if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "TRANSMIT":
                first_transmit_frame_index = i
                break

        if last_receive_frame_index is not None and first_transmit_frame_index is not None:
            if last_receive_frame_index > first_transmit_frame_index:
                return False

            return True

        else:
            return False

    except Exception as e:
        print(f"Error occurred while processing the XDM file: {e}")
        return False

def ordered_by_CAN(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)

        # Get all "d:ctr" elements inside "d:lst" with name="CanHardwareObject"
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        # Create a list to store the order of appearance
        ordered_can_refs = []

        for ctr in ctr_elements:
            ref_value = ctr.xpath("string(d:ref[@name='CanControllerRef']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})
            if not ref_value:
                continue

            # Find the matching CanControllerRef value
            for ref in ordered_can_refs:
                if ref in ref_value:
                    break
            else:
                ordered_can_refs.append(ref_value)

        for i in range(1, len(ordered_can_refs)):
            if ordered_can_refs[i - 1] > ordered_can_refs[i]:
                return False

        return True

    except Exception as e:
        print(f"Error occurred while processing the XDM file: {e}")
        return False

def ordered_by_id(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)

        # Get all "d:ctr" elements inside "d:lst" with name="CanHardwareObject"
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        # Extract the names of frames and their CanObjectId values in the order they appear in the file
        frames_data = [(ctr, ctr.xpath("string(d:var[@name='CanObjectId']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})) for ctr in ctr_elements]

        frames_data = [(ctr, obj_id) for ctr, obj_id in frames_data if obj_id.strip()]

        # Check if the CanObjectId values are unique
        can_object_ids = [obj_id for _, obj_id in frames_data]
        if len(can_object_ids) != len(set(can_object_ids)):
            return False

        # Check if the frames are neatly ordered based on CanObjectId
        for i in range(1, len(frames_data)):
            if int(frames_data[i - 1][1]) > int(frames_data[i][1]):
                return False

        return True

    except Exception as e:
        print(f"Error occurred while processing the XDM file: {e}")
        return False



def browse_xml():
    xml_file_path = filedialog.askopenfilename(filetypes=[("XML files", "*XDM")])
    if not xml_file_path:
        return
    xml_file_entry.delete(0, tk.END)
    xml_file_entry.insert(tk.END, xml_file_path)

def check_ordered_by_RX_TX():
    xml_file_path = xml_file_entry.get()
    if not xml_file_path:
        return

    is_ordered = ordered_by_RX_TX(xml_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by RX-TX.", fg="green")
    else:
        result_label.config(text="The XDM file is not ordered by RX-TX.", fg="red")

def check_ordered_by_CAN():
    xml_file_path = xml_file_entry.get()
    if not xml_file_path:
        return

    is_ordered = ordered_by_CAN(xml_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CAN.", fg="green")
    else:
        result_label.config(text="The XDM file is not ordered by CAN.", fg="red")

def check_ordered_by_Id():
    xml_file_path = xml_file_entry.get()
    if not xml_file_path:
        return

    is_ordered = ordered_by_id(xml_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CanObjectId.", fg="green")
    else:
        result_label.config(text="The XDM file is not ordered by CanObjectId.", fg="red")

root = tk.Tk()
root.title("XML File Order Checker")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

xml_file_label = tk.Label(frame, text="Select XML File:")
xml_file_label.grid(row=0, column=0, padx=5, pady=5)

xml_file_entry = tk.Entry(frame)
xml_file_entry.grid(row=0, column=1, padx=5, pady=5)

xml_file_button = tk.Button(frame, text="Browse", command=browse_xml)
xml_file_button.grid(row=0, column=2, padx=5, pady=5)

check_receive_transmit_button = tk.Button(frame, text="Check Ordered by RX_TX", command=check_ordered_by_RX_TX)
check_receive_transmit_button.grid(row=1, column=0, padx=5, pady=5)

check_can_controller_ref_button = tk.Button(frame, text="Check Ordered by Can", command=check_ordered_by_CAN)
check_can_controller_ref_button.grid(row=1, column=1, padx=5, pady=5)

check_can_object_id_button = tk.Button(frame, text="Check Ordered by CanObjectId", command=check_ordered_by_Id)
check_can_object_id_button.grid(row=1, column=2, padx=5, pady=5)

result_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
result_label.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
