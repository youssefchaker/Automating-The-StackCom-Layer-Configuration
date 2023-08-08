import statfuncs
from statfuncs import clear_excel,write_to_Excel,file_path,ordered_by_id_CanIf,tk,filedialog

sheet_name="CANIF_index_table_order_verif"

def check_order():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return
    result_data = {
        'Passed?':["X" if ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')==ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg') else " "],
        'Order by CanIfRxPduId':["None" if ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg')==True else ordered_by_id_CanIf(xdm_file_path,'CanIfRxPduId','CanIfRxPduCfg')],
        'Order by CanIfTxPduId':["None" if ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')==True else ordered_by_id_CanIf(xdm_file_path,'CanIfTxPduId','CanIfTxPduCfg')]
     }
    write_to_Excel(result_data,file_path,sheet_name)
    completion_label.config(text="Output Created", fg="green")

def clean_output(sheet_name):
    clear_excel(sheet_name)
    completion_label.config(text="Output File Cleared", fg="blue")
    
#open the xdm file
def browse_canif():
    xdm_file_path = filedialog.askopenfilename(filetypes=[("XDM files", "*XDM")])
    if not xdm_file_path:
        return
    xdm_file_entry.delete(0, tk.END)
    xdm_file_entry.insert(tk.END, xdm_file_path)

root = tk.Tk()
root.title("CanIf.xdm File Order Checker")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

xml_file_label = tk.Label(frame, text="Select CANIf File:")
xml_file_label.grid(row=0, column=0)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=0, column=1)

xdm_file_button = tk.Button(frame, text="Browse", command=browse_canif)
xdm_file_button.grid(row=0, column=2)

check_receive_transmit_button = tk.Button(frame, text="Check Order", command=check_order)
check_receive_transmit_button.grid(row=1, column=0, columnspan=3, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

clear_excel_button = tk.Button(frame, text="Clear Output", command=lambda:clean_output(sheet_name))
clear_excel_button.grid(row=2, column=0, columnspan=3, pady=5)

root.mainloop()
