import tkinter as tk
from tkinter import filedialog
from lxml import etree
import logging

#Preparing the log file
logging.basicConfig(filename='xdm_order_checker.log', level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

#Ordering by TRANSMIT and RECIEVE
def ordered_by_RX_TX(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        receive_indices = [i for i, ctr in enumerate(ctr_elements) if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "RECEIVE"]
        transmit_indices = [i for i, ctr in enumerate(ctr_elements) if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "TRANSMIT"]

        if receive_indices and transmit_indices:
            if receive_indices[-1] > transmit_indices[0]:
                frame_name = ctr_elements[transmit_indices[0]].attrib['name']
                logging.error(f"TRANSMIT frame '{frame_name}' comes before a RECEIVE frame.")
                return False

            if any(ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "RECEIVE" for ctr in ctr_elements[transmit_indices[-1] + 1:]):
                frame_name = ctr_elements[transmit_indices[-1]].attrib['name']
                logging.error(f"RECEIVE frame '{frame_name}' comes after a TRANSMIT frame.")
                return False

        return True

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
        return False

#Ordering by Cancontrollerref
def ordered_by_CAN_Ref(xdm_file):
    expected_order = ['CAN_2', 'CAN_1', 'CAN_DEVAID', 'CAN_3']
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        def check_order(frames):
            prev_index = None
            for ctr in frames:
                ref_value = ctr.xpath("string(d:ref[@name='CanControllerRef']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})
                index = next((i for i, val in enumerate(expected_order) if val in ref_value), None)
                if index is not None:
                    if prev_index is not None and index < prev_index:
                        frame_name = ctr.attrib['name']
                        logging.error(f"The frame '{frame_name}' has incorrect 'CanControllerRef' attribute order.")
                        return False
                    prev_index = index
                else:
                    frame_name = ctr.attrib['name']
                    logging.error(f"The frame '{frame_name}' has an invalid 'CanControllerRef' attribute.")
                    return False
            return True

        receive_frames = [ctr for ctr in ctr_elements if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "RECEIVE"]
        transmit_frames = [ctr for ctr in ctr_elements if ctr.xpath("string(d:var[@name='CanObjectType']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}) == "TRANSMIT"]

        receive_order_check = check_order(receive_frames)
        transmit_order_check = check_order(transmit_frames)

        return receive_order_check and transmit_order_check

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
        return False



#Ordering by CanObjectId
def ordered_by_id(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xml_content = file.read()

        root = etree.fromstring(xml_content)
        ctr_elements = root.xpath(".//d:lst[@name='CanHardwareObject']/d:ctr", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})

        frames_data = [(ctr.attrib['name'], ctr.xpath("string(d:var[@name='CanObjectId']/@value)", namespaces={'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'})) for ctr in ctr_elements]
        frames_data = [(name, obj_id) for name, obj_id in frames_data if obj_id.strip()]

        first_can_object_id = int(frames_data[0][1])
        if first_can_object_id != 0:
            logging.error(f"The first frame's CanObjectId should be '0', but found '{first_can_object_id}'.")
            return False

        can_object_ids = [int(obj_id) for _, obj_id in frames_data]
        if len(can_object_ids) != len(set(can_object_ids)):
            duplicates = [frame_name for frame_name, obj_id in frames_data if can_object_ids.count(int(obj_id)) > 1]
            for frame_name in duplicates:
                logging.error(f"The frame '{frame_name}' has a duplicate CanObjectId.")
            return False

        last_can_object_id = int(frames_data[-1][1])
        total_frames = len(frames_data)
        if last_can_object_id != total_frames - 1:
            logging.error(f"The last frame's CanObjectId should be '{total_frames - 1}', but found '{last_can_object_id}'.")
            return False

        if any(int(frames_data[i - 1][1]) > int(frames_data[i][1]) for i in range(1, len(frames_data))):
            frame_name = frames_data[next(i for i in range(1, len(frames_data)) if int(frames_data[i - 1][1]) > int(frames_data[i][1]))][0]
            logging.error(f"The frame '{frame_name}' has a jump in CanObjectId.")
            return False

        return True

    except Exception as e:
        logging.error(f"Error occurred while processing the XDM file: {e}")
        return False
#open the xdm file
def browse_xdm():
    xdm_file_path = filedialog.askopenfilename(filetypes=[("XDM files", "*XDM")])
    if not xdm_file_path:
        return
    xdm_file_entry.delete(0, tk.END)
    xdm_file_entry.insert(tk.END, xdm_file_path)

#check RX_TX
def check_ordered_by_RX_TX():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return

    is_ordered = ordered_by_RX_TX(xdm_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by RX-TX.", fg="green")
        logging.info("The XDM file is ordered by RX-TX.")
    else:
        result_label.config(text="The XDM file is not ordered by RX-TX.", fg="red")

# check by cancontrollerref
def check_ordered_by_CAN_Ref():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return

    is_ordered = ordered_by_CAN_Ref(xdm_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CAN_REF.", fg="green")
        logging.info("The XDM file is ordered by CAN.")
    else:
        result_label.config(text="The XDM file is not ordered by CAN_REF.", fg="red")

#check by CanObjectId
def check_ordered_by_Id():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return

    is_ordered = ordered_by_id(xdm_file_path)

    if is_ordered:
        result_label.config(text="The XDM file is ordered by CanObjectId.", fg="green")
        logging.info("The XDM file is ordered by CanObjectId.")
    else:
        result_label.config(text="The XDM file is not ordered by CanObjectId.", fg="red")

#clear the log file
def clear_log():
    open('xdm_order_checker.log', 'w').close()
    logging.info("Log file cleared.")

root = tk.Tk()
root.title("XDM File Order Checker")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

xml_file_label = tk.Label(frame, text="Select CAN File:")
xml_file_label.grid(row=0, column=0)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=0, column=1)

xml_file_button = tk.Button(frame, text="Browse", command=browse_xdm)
xml_file_button.grid(row=0, column=2)

check_receive_transmit_button = tk.Button(frame, text="Check Ordered by RX_TX", command=check_ordered_by_RX_TX)
check_receive_transmit_button.grid(row=1, column=0, pady=5)

check_can_controller_ref_button = tk.Button(frame, text="Check Ordered by Can-REF", command=check_ordered_by_CAN_Ref)
check_can_controller_ref_button.grid(row=1, column=1, pady=5)

check_can_object_id_button = tk.Button(frame, text="Check Ordered by Id", command=check_ordered_by_Id)
check_can_object_id_button.grid(row=1, column=2, pady=5)

clear_log_button = tk.Button(frame, text="Clear Log", command=clear_log)
clear_log_button.grid(row=2, column=0, columnspan=3, pady=5)

result_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
result_label.grid(row=3, column=0, columnspan=3, pady=5)

root.mainloop()
