import statfuncs
from statfuncs import *

sheet_name="COM_index_table_ComIpdu_order_verif"

#Order by SEND and RECIEVE check
def ordered_by_TX_RX(xdm_file):
    try:
        with open(xdm_file, 'r') as file:
            xdm_content = file.read()

        root = etree.fromstring(xdm_content)
        ctr_elements = root.xpath(".//d:lst[@name='ComIPdu']/d:ctr", namespaces=namespace)

        receive_indices = [i for i, ctr in enumerate(ctr_elements) if ctr.xpath("string(d:var[@name='ComIPduDirection']/@value)", namespaces=namespace) == "RECEIVE"]
        send_indices = [i for i, ctr in enumerate(ctr_elements) if ctr.xpath("string(d:var[@name='ComIPduDirection']/@value)", namespaces=namespace) == "SEND"]

        if send_indices and receive_indices:
            if send_indices[-1] > receive_indices[0]:
                frame_name = ctr_elements[receive_indices[-1]].attrib['name']
                return "SEND frames after RECEIVE frame("+frame_name+")"

            if any(ctr.xpath("string(d:var[@name='ComIPduDirection']/@value)", namespaces=namespace) == "SEND" for ctr in ctr_elements[receive_indices[-1] + 1:]):
                frame_name = ctr_elements[receive_indices[-1]].attrib['name']
                return "SEND frames after RECEIVE frame("+frame_name+")"

        return True

    except Exception as e:
        print(f"Error occurred : {e}")
        return False

#function to check for ascending order of Frame IDs
def ordered_by_id_Messagerie(excel_file_path):
    try:
        frames_data = getFullFrameData(excel_file_path)
        identifiers = frames_data["Identifiant_T"]
        identifiers = identifiers.apply(lambda x: int(x, 16))

        if identifiers.is_monotonic_increasing:
            return True
        else:
            disordered_index = identifiers[identifiers.diff() < 0].index[0]
            problematic_frame = frames_data.loc[disordered_index]
            frame_name=frames_data.loc[disordered_index]["Radical"]
            return f"Frames are not ordered. Problematic frame:({frame_name})"

    except Exception as e:
        print(f"Error occurred : {e}")
        return False

#check order for all critera at once
def check_order():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return
    excel_file_path = excel_file_entry.get()
    result_data = {
        'Passed?':["OK" if ordered_by_id_COM(xdm_file_path,"ComIPduHandleId","ComIPdu")==ordered_by_id_COM(xdm_file_path,"ComHandleId","ComSignal")==ordered_by_TX_RX(xdm_file_path) else "NOK"],
        'Order Frames by TX_RX Errors':["None" if ordered_by_TX_RX(xdm_file_path)==True else ordered_by_TX_RX(xdm_file_path)],
        'Order Frames by ComIPduHandleId Errors':["None" if ordered_by_id_COM(xdm_file_path,"ComIPduHandleId","ComIPdu")==True else ordered_by_id_COM(xdm_file_path,"ComIPduHandleId","ComIPdu")],
        'Order Signals by ComHandleId Errors':["None" if ordered_by_id_COM(xdm_file_path,"ComHandleId","ComSignal")==True else ordered_by_id_COM(xdm_file_path,"ComHandleId","ComSignal")],
        'Order Frames by Messagerie Identifiant_T Errors':["None" if ordered_by_id_Messagerie(excel_file_path)==True else ordered_by_id_Messagerie(excel_file_path)]

     }
    write_to_Excel(result_data,file_path,sheet_name)
    completion_label.config(text="Output Created", fg="green")


def browse_excel(excel_file_entry):
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not excel_file_path:
        return
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(tk.END, excel_file_path)

# tkinter Interface
root = tk.Tk()
root.title("Com File Order Checker")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=lambda:browse_excel(excel_file_entry))
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

xdm_file_label = tk.Label(frame, text="Select Com File:")
xdm_file_label.grid(row=1, column=0, padx=5, pady=5)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=1, column=1, padx=5, pady=5)

xdm_file_button = tk.Button(frame, text="Browse", command=lambda:browse_xdm(xdm_file_entry))
xdm_file_button.grid(row=1, column=2, padx=5, pady=5)

check_receive_transmit_button = tk.Button(frame, text="Check Order", command=check_order)
check_receive_transmit_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

clear_excel_button = tk.Button(frame, text="Clear Output", command=lambda:clear_excel(sheet_name,completion_label))
clear_excel_button.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
