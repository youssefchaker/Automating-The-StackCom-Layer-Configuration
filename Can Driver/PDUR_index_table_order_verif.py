import statfuncs
from statfuncs import clear_excel,write_to_Excel,file_path,ordered_by_id_PDUR,tk,filedialog

sheet_name="PDUR_index_table_order_verif"
nodes_src_Tx=['Com_PduRRoutingTable','PduRRoutingPath','PduRSrcPdu','PduRSourcePduHandleId']
nodes_src_Rx=['CanIf_PduRRoutingTable','PduRRoutingPath','PduRSrcPdu','PduRSourcePduHandleId']
nodes_dest_Tx=['Com_PduRRoutingTable','PduRRoutingPath','PduRDestPdu','PduRDestPduHandleId']
nodes_dest_Rx=['CanIf_PduRRoutingTable','PduRRoutingPath','PduRDestPdu','PduRDestPduHandleId']


def check_order():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return
    result_data = {
        'Passed?':["X" if ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)==True else " "],
        'Order by Tx_PduRSourcePduHandleId':[" " if ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)],
        'Order by Tx_PduRDestPduHandleId':[" " if ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)],
        'Order by Rx_PduRSourcePduHandleId':[" " if ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)],
        'Order by Rx_PduRDestPduHandleId':[" " if ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)]
        
     }
    write_to_Excel(result_data,file_path,sheet_name)
    completion_label.config(text="Output Created", fg="green")

    
#open the xdm file
def browse_pdur():
    xdm_file_path = filedialog.askopenfilename(filetypes=[("XDM files", "*XDM")])
    if not xdm_file_path:
        return
    xdm_file_entry.delete(0, tk.END)
    xdm_file_entry.insert(tk.END, xdm_file_path)

def clean_output(sheet_name):
    clear_excel(sheet_name)
    completion_label.config(text="Output File Cleared", fg="blue")

root = tk.Tk()
root.title("PDUR.xdm File Order Checker")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

xml_file_label = tk.Label(frame, text="Select PDUR File:")
xml_file_label.grid(row=0, column=0)

xdm_file_entry = tk.Entry(frame)
xdm_file_entry.grid(row=0, column=1)

xdm_file_button = tk.Button(frame, text="Browse", command=browse_pdur)
xdm_file_button.grid(row=0, column=2)

check_receive_transmit_button = tk.Button(frame, text="Check Order", command=check_order)
check_receive_transmit_button.grid(row=1, column=0, columnspan=3, pady=5)

clear_excel_button = tk.Button(frame, text="Clear Excel", command=lambda:clean_output(sheet_name))
clear_excel_button.grid(row=2, column=0, columnspan=3, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
