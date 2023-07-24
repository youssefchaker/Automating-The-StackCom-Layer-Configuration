import tkinter as tk
from tkinter import filedialog
from lxml import etree
import logging

# Configure logging
logging.basicConfig(filename='xdm_order_checker.log', level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

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
                frame_name = ctr_elements[first_transmit_frame_index].attrib['name']
                logging.error(f"Received frame '{frame_name}' comes after the first transmit frame.")
                return False

            return True

        else:
            return False

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
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
                frame_name = ctr_elements[i].attrib['name']
                logging.error(f"The frame '{frame_name}' is not in the correct order.")
                return False

        return True

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
        return False

def ordered_by_id(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)

        # Get all "d:ctr" elements inside "d:lst" with name="CanHardwareObject"
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        # Extract the names of frames and their CanObjectId values in the order they appear in the file
        frames_data = [(ctr.attrib['name'], ctr.xpath("string(d:var[@name='CanObjectId']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})) for ctr in ctr_elements]

        frames_data = [(name, obj_id) for name, obj_id in frames_data if obj_id.strip()]

        # Check if the first frame starts with CanObjectId 0
        first_can_object_id = int(frames_data[0][1])
        if first_can_object_id != 0:
            logging.error(f"The first frame's CanObjectId should be '0', but found '{first_can_object_id}'.")
            return False

        # Check if the CanObjectId values are unique and in order
        can_object_ids = [int(obj_id) for _, obj_id in frames_data]
        if len(can_object_ids) != len(set(can_object_ids)):
            duplicates = [frame_name for frame_name, obj_id in frames_data if can_object_ids.count(int(obj_id)) > 1]
            for frame_name in duplicates:
                logging.error(f"The frame '{frame_name}' has a duplicate CanObjectId.")
            return False

        # Check if the frames are neatly ordered based on CanObjectId
        last_can_object_id = int(frames_data[-1][1])
        total_frames = len(frames_data)
        if last_can_object_id != total_frames - 1:
            logging.error(f"The last frame's CanObjectId should be '{total_frames - 1}', but found '{last_can_object_id}'.")
            return False

        for i in range(1, len(frames_data)):
            if int(frames_data[i - 1][1]) > int(frames_data[i][1]):
                frame_name = frames_data[i][0]
                logging.error(f"The frame '{frame_name}' has a jump in CanObjectId.")
                return False

        return True

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
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
        logging.info("The XDM file is ordered by RX-TX.")
    else:
        result_label.config(text="The XDM file is not ordered by RX-TX.", fg="red")

def check_ordered_by_CAN():
    xml_file_path = xml_file_entry.get()
    if not xml_file_path:
        return

    is_ordered = ordered_by_CAN(xml_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CAN.", fg="green")
        logging.info("The XDM file is ordered by CAN.")
    else:
        result_label.config(text="The XDM file is not ordered by CAN.", fg="red")

def check_ordered_by_Id():
    xml_file_path = xml_file_entry.get()
    if not xml_file_path:
        return

    is_ordered = ordered_by_id(xml_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CanObjectId.", fg="green")
        logging.info("The XDM file is ordered by CanObjectId.")
    else:
        result_label.config(text="The XDM file is not ordered by CanObjectId.", fg="red")

def clear_log():
    # Clear the log file
    open('xdm_order_checker.log', 'w').close()
    logging.info("Log file cleared.")

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

check_can_object_id_button = tk.Button(frame, text="Check Ordered by Id", command=check_ordered_by_Id)
check_can_object_id_button.grid(row=1, column=2, padx=5, pady=5)

clear_log_button = tk.Button(frame, text="Clear Log", command=clear_log)
clear_log_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)  # Centered the button by spanning three columns

result_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
result_label.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()