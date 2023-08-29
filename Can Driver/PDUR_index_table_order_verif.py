import statfuncs
from statfuncs import *

sheet_name="PDUR_index_table_order_verif"

#elements for every type of order
nodes_src_Tx=['Com_PduRRoutingTable','PduRRoutingPath','PduRSrcPdu','PduRSourcePduHandleId']
nodes_src_Rx=['CanIf_PduRRoutingTable','PduRRoutingPath','PduRSrcPdu','PduRSourcePduHandleId']
nodes_dest_Tx=['Com_PduRRoutingTable','PduRRoutingPath','PduRDestPdu','PduRDestPduHandleId']
nodes_dest_Rx=['CanIf_PduRRoutingTable','PduRRoutingPath','PduRDestPdu','PduRDestPduHandleId']

#check order for all 4 criteras
def check_order():
    xdm_file_path = xdm_file_entry.get()
    if not xdm_file_path:
        return
    result_data = {
        'Passed?':["OK" if ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)==True and ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)==True else "NOK"],
        'Order by Tx_PduRSourcePduHandleId':["None" if ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_src_Tx)],
        'Order by Tx_PduRDestPduHandleId':["None" if ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_dest_Tx)],
        'Order by Rx_PduRSourcePduHandleId':["None" if ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_src_Rx)],
        'Order by Rx_PduRDestPduHandleId':["None" if ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)==True else ordered_by_id_PDUR(xdm_file_path,nodes_dest_Rx)]
        
     }
    write_to_Excel(result_data,file_path,sheet_name)
    completion_label.config(text="Output Created", fg="green")

#tkinter Interface
root = tk.Tk()
root.title("PDUR Order Checker")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

xml_file_label = tk.Label(frame, text="Select PDUR File:")
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
