import statfuncs
from statfuncs import *

sheet_name="CANIF_PDU_Messagerie_verif"

# Function to extract necessary attributes for the target frame from the .xdm file
def extract_CanifValues(xdm_file, frame_name,excel_file_path):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()
    
    frames_data = cleanExcelFrameData(excel_file_path)

    root = etree.fromstring(xdm_content)
    selected_frame = frames_data[frames_data['Radical'] == frame_name]
    if selected_frame.empty:
        result_data = {
            'Frame Name': [frame_name],
            'Passed?': ["Frame Not Found in Messagerie "],
            'Frame type':' ',
            'CanIfCanCtrlIdRef':' ',
            'AEE10r3 Reseau_T': ' ',
            'CanIfCanCtrlIdRef/AEE10r3 Reseau_T':' ',
            'CanIfCanHandleTypeRef':' ',
            'CanIfCanHandleTypeRef/Frame Name': ' ',
            'CanIfIdSymRef':' ',
            'CanIfIdSymRef/Frame Name':' ',
        }
        write_to_Excel(result_data,file_path,sheet_name)
        CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef= -1, -1, -1
        return CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef
    else:
        if( not selected_frame["UCE Emetteur"].str.endswith("E_VCU").any()):
            ctr_elements = root.xpath(".//d:lst[@name='CanIfHrhCfg']/d:ctr[contains(@name, $name)]", namespaces=namespace, name=frame_name)
            if ctr_elements:
                CanIfCanCtrlIdRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHrhCanCtrlIdRef']/@value)", namespaces=namespace)
                CanIfCanHandleTypeRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHrhCanHandleTypeRef']/@value)", namespaces=namespace)
                CanIfIdSymRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHrhIdSymRef']/@value)", namespaces=namespace)
            else:
                CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef= None, None, None
        else:
            ctr_elements = root.xpath(".//d:lst[@name='CanIfHthCfg']/d:ctr[contains(@name, $name)]", namespaces=namespace, name=frame_name)
            if ctr_elements:
                CanIfCanCtrlIdRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHthCanCtrlIdRef']/@value)", namespaces=namespace)
                CanIfCanHandleTypeRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHthCanHandleTypeRef']/@value)", namespaces=namespace)
                CanIfIdSymRef = ctr_elements[0].xpath("string(d:ref[@name='CanIfHthIdSymRef']/@value)", namespaces=namespace)
            else:
                CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef= None, None, None
        return CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef
           

def verify_frame(excel_file_path, xdm_file_path, frame_name):
    try:
        CanIfCanCtrlIdRef, CanIfCanHandleTypeRef, CanIfIdSymRef = extract_CanifValues(xdm_file_path, frame_name,excel_file_path)
        if CanIfCanCtrlIdRef== -1 and CanIfCanHandleTypeRef== -1 and CanIfIdSymRef== -1 :
            return
        elif CanIfCanCtrlIdRef is None and CanIfCanHandleTypeRef is None and CanIfIdSymRef is None :
            result_data = {
                    'Frame Name': [frame_name],
                    'Passed?':["Frame Not Found in CANIF"],
                    'Frame type':' ',
                    'CanIfCanCtrlIdRef':' ',
                    'AEE10r3 Reseau_T': ' ',
                    'CanIfCanCtrlIdRef/AEE10r3 Reseau_T':' ',
                    'CanIfCanHandleTypeRef':' ',
                    'CanIfCanHandleTypeRef/Frame Name': ' ',
                    'CanIfIdSymRef':' ',
                    'CanIfIdSymRef/Frame Name':' ',
                }
            write_to_Excel(result_data,file_path,sheet_name)
            return False
        else:
            frames_data = cleanExcelFrameData(excel_file_path)
            selected_frame = frames_data[frames_data['Radical'] == frame_name]
            CanIfCanCtrlIdReftst=CanIfCanHandleTypeReftst=CanIfIdSymReftst=True

            if "ASPath:/CanIf/CanIf/CanIfCtrlDrvCfg/CanIf_Controller_Can_InterSystem" == CanIfCanCtrlIdRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("HS1"):
                pass
            elif "ASPath:/CanIf/CanIf/CanIfCtrlDrvCfg/CanIf_Controller_Can_LAS" == CanIfCanCtrlIdRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("HS2"):
                pass
            elif "ASPath:/CanIf/CanIf/CanIfCtrlDrvCfg/CanIf_Controller_Can_eCAN" == CanIfCanCtrlIdRef and selected_frame["AEE10r3 Reseau_T"].values[0].startswith("E_CAN"):
                pass
            else:
                    CanIfCanCtrlIdReftst=False
            
            if(frame_name not in CanIfCanHandleTypeRef):
                CanIfCanHandleTypeReftst=False
            if(frame_name not in CanIfIdSymRef):
                CanIfIdSymReftst=False
            result_data = {
                'Frame Name': [frame_name],
                'Passed?':["NOK" if CanIfCanCtrlIdReftst ==False or CanIfIdSymReftst==False or CanIfCanHandleTypeReftst==False else "OK"],
                'Frame type':["TRANSMIT" if selected_frame["UCE Emetteur"].str.endswith("E_VCU").any() else "RECEIVE" ],
                'CanIfCanCtrlIdRef':[CanIfCanCtrlIdRef],
                'AEE10r3 Reseau_T': [selected_frame["AEE10r3 Reseau_T"].values[0]],
                'CanIfCanCtrlIdRef/AEE10r3 Reseau_T Errors': ['Error (CanIfCanCtrlIdRef Mismatch)' if CanIfCanCtrlIdReftst==False else "None"],
                'CanIfCanHandleTypeRef':[CanIfCanHandleTypeRef],
                'CanIfCanHandleTypeRef/Frame Name Errors': ['Error (Frame Name not present in CanIfCanHandleTypeRef)' if CanIfCanHandleTypeReftst==False else "None"],
                'CanIfIdSymRef':[CanIfIdSymRef],
                'CanIfIdSymRef/Frame Name Errors': ['Error (Frame Name not present in CanIfIdSymRef)' if CanIfIdSymReftst==False else "None"],
            }
            write_to_Excel(result_data,file_path,sheet_name)

    except Exception as e:
                print(f"Error occurred : {e}")
                return False     


#execute functionality on button click
def verify_button_click():
    excel_file_path = excel_file_entry.get()
    xdm_file_path = xdm_file_entry.get()
    frame_name = frame_entry.get()

    verify_frame(excel_file_path, xdm_file_path, frame_name)
    completion_label.config(text="Output Created", fg="green")


# tkinter Interface
root = tk.Tk()
root.title("CanIf Frame Info CANIF/Messagerie Verification")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

frame_label = tk.Label(frame, text="Enter Frame Name:")
frame_label.grid(row=2, column=0, padx=5, pady=5)

frame_entry = ttk.Combobox(frame)
frame_entry.grid(row=2, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=lambda:browse_excel_frames(excel_file_entry,frame_entry))
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

xdm_file_label = tk.Label(frame, text="Select Canif File:")
xdm_file_label.grid(row=1, column=0, padx=5, pady=5)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=1, column=1, padx=5, pady=5)

xdm_file_button = tk.Button(frame, text="Browse", command=lambda:browse_xdm(xdm_file_entry))
xdm_file_button.grid(row=1, column=2, padx=5, pady=5)

verify_button = tk.Button(frame, text="Verify", command=verify_button_click)
verify_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

clear_excel_button = tk.Button(frame, text="Clear Output", command=lambda:clear_excel(sheet_name,completion_label))
clear_excel_button.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
