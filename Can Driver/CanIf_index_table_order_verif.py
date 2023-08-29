import statfuncs
from statfuncs import *

sheet_name="CANIF_index_table_order_verif"


#function to check the order
def check_order():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return
    result_data = {
        'Passed?':["OK" if ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')==ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg') else "NOK"],
        'Order by CanIfRxPduId':["None" if ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg')==True else ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg')],
        'Order by CanIfTxPduId':["None" if ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')==True else ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')]
     }
    write_to_Excel(result_data,file_path,sheet_name)
    completion_label.config(text="Output Created", fg="green")


#tkinter Interface
root = tk.Tk()
root.title("CanIf Order Checker")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

xml_file_label = tk.Label(frame, text="Select CANIf File:")
xml_file_label.grid(row=0, column=0)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=0, column=1)

xdm_file_button = tk.Button(frame, text="Browse", command=lambda:browse_xdm(xdm_file_entry))
xdm_file_button.grid(row=0, column=2)

check_receive_transmit_button = tk.Button(frame, text="Check Order", command=check_order)
check_receive_transmit_button.grid(row=1, column=0, columnspan=3, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

clear_excel_button = tk.Button(frame, text="Clear Output", command=lambda:clear_excel(sheet_name,completion_label))
clear_excel_button.grid(row=2, column=0, columnspan=3, pady=5)

root.mainloop()
