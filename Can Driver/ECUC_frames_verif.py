import statfuncs
from statfuncs import clear_excel,write_to_Excel,file_path,etree,cleanExcelFrameData,tk,filedialog,namespace

sheet_name="ECUC_frames_verif"

# Function to extract necessary attributes for the target frame from the .xdm file
def extract_ECUC_values(xdm_file, frame_name):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()

    root = etree.fromstring(xdm_content)
    elements = root.xpath(".//d:lst[@name='Pdu']/d:ctr[@name=$name]/d:var[@name='PduLength']", namespaces=namespace, name=frame_name)
    if elements:
        PduLength=elements[0].get("value")
        return PduLength
    else:
        return None

# Function to verify the frame attributes from the Excel file with the attributes from the .xdm file
def verify_frame(excel_file_path, xdm_file_path, frame_name):
    try:
        PduLength= extract_ECUC_values(xdm_file_path, frame_name)
        if PduLength is None:
            result_data = {
                'Frame Name': [frame_name],
                'Passed?': ["Frame Not Found in EcuC"],
                'PduLength':' ',
                'frame_size':' ', 
                'PduLength/frame_size':' '
            }
            write_to_Excel(result_data,file_path,sheet_name)
            return False
        else:
            
            frames_data = cleanExcelFrameData(excel_file_path)
            selected_frame = frames_data[frames_data['Radical'] == frame_name]

            if selected_frame.empty:
                result_data = {
                    'Frame Name': [frame_name],
                    'Passed?': ["Frame Not Found in Messagerie "],
                }
                write_to_Excel(result_data,file_path,sheet_name)
                return False
            else:
                PduLengthtst=True
                frame_size = selected_frame["Taille_Max_T"].values[0]
                
                if(int(PduLength)!=int(frame_size)):
                    PduLengthtst=False

                result_data = {
                    'Frame Name': [frame_name],
                    'Passed?':[" " if PduLengthtst==False else "X"],
                    'PduLength':[PduLength],
                    'frame_size':[frame_size],
                    'PduLength/frame_size':["Error(Frame Size Mismatch)" if PduLengthtst==False else "None"]
                }
                write_to_Excel(result_data,file_path,sheet_name)
    except Exception as e:
                print(f"Error occurred : {e}")
                return False   

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


#execute functionality on button click
def verify_button_click():
    excel_file_path = excel_file_entry.get()
    xdm_file_path = xdm_file_entry.get()
    frame_name = frame_entry.get()

    verify_frame(excel_file_path, xdm_file_path, frame_name)
    completion_label.config(text="Output Created", fg="green")

def clean_output(sheet_name):
    clear_excel(sheet_name)
    completion_label.config(text="Output File Cleared", fg="blue")

# Create the GUI
root = tk.Tk()
root.title("ECUC.xdm Frame Info Verification")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=browse_excel)
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

xdm_file_label = tk.Label(frame, text="Select Ecuc File:")
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

clear_excel_button = tk.Button(frame, text="Clear Output", command=lambda:clean_output(sheet_name))
clear_excel_button.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

completion_label = tk.Label(frame, text="", fg="green")
completion_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()