import pandas as pd
import tkinter as tk
from tkinter import filedialog
from lxml import etree

expected_headers = {'FRAMES': ['Radical', 'Activation trame', 'Protocole_M', 'Identifiant_T', 'Taille_Max_T', 'Lmin_T', 'Mode_Transmission_T', 'Nature_Evenement_FR_T', 'Nature_Evenement_GB_T', 'Periode_T', 'UCE Emetteur', 'AEE10r3 Reseau_T']}

#clean the excel file from uncessary data
def cleanExcelData(excel_file):
    df = pd.read_excel(excel_file, sheet_name='FRAMES', header=0)
    headers = [col for col in df.columns if col in expected_headers['FRAMES']]
    return df[headers]

#extract the necessary attributes for the target frame
def extract_CanValues(xdm_file, frame_name):
    with open(xdm_file, 'r') as file:
        xml_content = file.read()

    root = etree.fromstring(xml_content)
    namespace = {'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd'}

    ctr_elements = root.xpath(".//d:ctr[contains(@name, $name)]", namespaces=namespace, name=frame_name)
    if ctr_elements:
        CanIdValue = int(ctr_elements[0].xpath("d:var[@name='CanIdValue']/@value", namespaces=namespace)[0])
        CanObjectType = ctr_elements[0].xpath("string(d:var[@name='CanObjectType']/@value)", namespaces=namespace)
        CanIdType = ctr_elements[0].xpath("string(d:var[@name='CanIdType']/@value)", namespaces=namespace)
        CanHandleType = ctr_elements[0].xpath("string(d:var[@name='CanHandleType']/@value)", namespaces=namespace)
        CanControllerRef = ctr_elements[0].xpath("string(d:ref[@name='CanControllerRef']/@value)", namespaces=namespace)
    else:
        CanIdValue, CanObjectType, CanIdType, CanHandleType, CanControllerRef = None, None, None, None, None

    return CanIdValue, CanObjectType, CanIdType, CanHandleType, CanControllerRef

#verify the frame attributes from the excel file with the attributes from the .xdm file
def verify_frame(excel_file_path, xdm_file_path, signal_name):
    CanIdValue, CanObjectType, CanIdType, CanHandleType, CanControllerRef = extract_CanValues(xdm_file_path, signal_name)
    if CanIdValue is None or CanObjectType is None or CanIdType is None or CanHandleType is None or CanControllerRef is None:
        result_label.config(text="Frame Not Found", fg="red")
        can_id_value_label.config(text="")
        return

    if CanIdType != "STANDARD":
        result_label.config(text="Fail (Incorrect CanIdType).", fg="red")
        can_id_value_label.config(text="")
        return

    if CanHandleType != "FULL":
        result_label.config(text="Fail (Incorrect CanHandleType).", fg="red")
        can_id_value_label.config(text="")
        return

    frames_data = cleanExcelData(excel_file_path)
    selected_frame = frames_data[frames_data['Radical'] == signal_name]

    if selected_frame.empty:
        result_label.config(text="Frame Id Mismatch.", fg="red")
        can_id_value_label.config(text="")
        return

    identifiant_t_hex = selected_frame["Identifiant_T"].values[0]
    try:
        identifiant_t_decimal = int(identifiant_t_hex, 16)

        if "/Can/Can/CanConfigSet_0/CAN_1" in CanControllerRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("HS1"):
            pass
        elif "/Can/Can/CanConfigSet_0/CAN_2" in CanControllerRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("HS2"):
            pass
        elif "/Can/Can/CanConfigSet_0/CAN_3" in CanControllerRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("E_CAN"):
            pass
        else:
            result_label.config(text="Fail (Incorrect CanControllerRef/AEE10r3 Reseau_T association).", fg="red")
            can_id_value_label.config(text="")
            return

        result_label.config(text="Confirmed", fg="green")
        can_id_value_label.config(text=f"CanObjectType : {CanObjectType}\n CanIdType : {CanIdType}\n CanHandleType : {CanHandleType}\n CanIdValue : {CanIdValue}\n CanControllerRef : {selected_frame['AEE10r3 Reseau_T'].values[0]}\n")

    except ValueError:
        result_label.config(text="Invalid Identifiant_T format.", fg="red")
        can_id_value_label.config(text="")

#select the excel file from the interface
def browse_excel():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_file_path:
        return
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(tk.END, excel_file_path)
#select the xdm file from the interface
def browse_xdm():
    xdm_file_path = filedialog.askopenfilename(filetypes=[("XDM files", "*.xdm")])
    if not xdm_file_path:
        return
    xdm_file_entry.delete(0, tk.END)
    xdm_file_entry.insert(tk.END, xdm_file_path)
#execute functionality on button clieck
def verify_button_click():
    excel_file_path = excel_file_entry.get()
    xdm_file_path = xdm_file_entry.get()
    frame_name = frame_entry.get()

    verify_frame(excel_file_path, xdm_file_path, frame_name)

# Create the GUI
root = tk.Tk()
root.title("Frame Verification")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=browse_excel)
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

xdm_file_label = tk.Label(frame, text="Select CAN.xdm File:")
xdm_file_label.grid(row=1, column=0, padx=5, pady=5)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=1, column=1, padx=5, pady=5)

xdm_file_button = tk.Button(frame, text="Browse", command=browse_xdm)
xdm_file_button.grid(row=1, column=2, padx=5, pady=5)

frame_label = tk.Label(frame, text="Enter Frame Name:")
frame_label.grid(row=2, column=0, padx=5, pady=5)

frame_entry = tk.Entry(frame)
frame_entry.grid(row=2, column=1, padx=5, pady=5)

verify_button = tk.Button(frame, text="Verify", command=verify_button_click)
verify_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

result_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
result_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

can_id_value_label = tk.Label(frame, text="", font=("Arial", 12, "bold"))
can_id_value_label.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
