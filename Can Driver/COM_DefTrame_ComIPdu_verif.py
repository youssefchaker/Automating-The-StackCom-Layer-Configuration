import statfuncs
from statfuncs import clear_excel,write_to_Excel,file_path,cleanExcelSignalData,etree,tk,filedialog,namespace,Signal_Type,Signal_Position_inFrame,Signal_Type,cleanExcelFrameData,math

sheet_name="COM_DefTrame_ComIPdu_verif"

# Function to extract necessary attributes for the target signal from the .xdm file
def extract_ComValues(xdm_file, frame_name):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()

    root = etree.fromstring(xdm_content)
    elements = root.xpath(".//d:lst[@name='ComIPdu']/d:ctr[@name=$name]", namespaces=namespace, name=frame_name)
    if elements:
        ComIPduDirection=elements[0].xpath("d:var[@name='ComIPduDirection']/@value", namespaces=namespace)[0]
        ComIPduSignalProcessing = elements[0].xpath("d:var[@name='ComIPduSignalProcessing']/@value", namespaces=namespace)[0]
        ComIPduType = elements[0].xpath("d:var[@name='ComIPduType']/@value", namespaces=namespace)[0]
        ComPduIdRef = elements[0].xpath("d:ref[@name='ComPduIdRef']/@value", namespaces=namespace)[0]
        ComIPduSignalRef=elements[0].xpath("d:lst[@name='ComIPduSignalRef']/d:ref", namespaces=namespace)
        signals = [ref.attrib['value'] for ref in ComIPduSignalRef]
        signals = [signal.replace('ASPath:/Com/Com/ComConfig/', '') for signal in signals]
        ComIPduCallout = elements[0].xpath("d:var[@name='ComIPduCallout']/@value", namespaces=namespace)[0]
        ComIPduCounter = elements[0].xpath("d:ctr[@name='ComIPduCounter']/d:var[@name='ComIPduCounterErrorNotification']/@value", namespaces=namespace)
        ComIPduCounter = list(filter(None, ComIPduCounter))
        if(ComIPduCounter):
            ComIPduCounter=ComIPduCounter[0]
        ComIPduCounterSize = elements[0].xpath("d:ctr[@name='ComIPduCounter']/d:var[@name='ComIPduCounterSize']/@value", namespaces=namespace)
        ComIPduCounterSize = list(filter(None, ComIPduCounterSize))
        if(ComIPduCounterSize):
            ComIPduCounterSize=ComIPduCounterSize[0]
        ComIPduCounterStartPosition = elements[0].xpath("d:ctr[@name='ComIPduCounter']/d:var[@name='ComIPduCounterStartPosition']/@value", namespaces=namespace)
        ComIPduCounterStartPosition = list(filter(None, ComIPduCounterStartPosition))
        if(ComIPduCounterStartPosition):
            ComIPduCounterStartPosition=ComIPduCounterStartPosition[0]
        ComTxModeMode = elements[0].xpath("d:ctr[@name='ComTxMode']/d:var[@name='ComTxModeMode']/@value", namespaces=namespace)
        ComTxModeMode = list(filter(None, ComTxModeMode))
        if(ComTxModeMode):
            ComTxModeMode=ComTxModeMode[0]
        ComTxModeTimePeriod=elements[0].xpath("d:ctr[@name='ComTxMode']/d:var[@name='ComTxModeTimePeriod']/@value", namespaces=namespace)
        ComTxModeTimePeriod = list(filter(None, ComTxModeTimePeriod))
        if(ComTxModeTimePeriod):
            ComTxModeTimePeriod=ComTxModeTimePeriod[0]
    else:
        ComIPduDirection=ComIPduSignalProcessing=ComIPduType=ComPduIdRef=signals=ComIPduCallout=ComIPduCounter=ComIPduCounterSize=ComIPduCounterStartPosition=ComTxModeMode=ComTxModeTimePeriod= None

    return ComIPduDirection,ComIPduSignalProcessing,ComIPduType,ComPduIdRef,signals,ComIPduCallout,ComIPduCounter,ComIPduCounterSize,ComIPduCounterStartPosition,ComTxModeMode,ComTxModeTimePeriod

def Frame_Signals(xdm_file,frame_name):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()

    root = etree.fromstring(xdm_content)
    elements = root.xpath(".//d:lst[@name='ComSignal']/d:ctr[contains(@name, $name)]", namespaces=namespace, name=frame_name)
    if elements:
        signals = [ctr.attrib['name'] for ctr in elements]
        signals.reverse()
    else:
        signals= None

    return signals

def verify_frame(excel_file_path,xdm_file_path, frame_name):
    try:
        ComIPduDirection,ComIPduSignalProcessing,ComIPduType,ComPduIdRef,signals,ComIPduCallout,ComIPduCounter,ComIPduCounterSize,ComIPduCounterStartPosition,ComTxModeMode,ComTxModeTimePeriod= extract_ComValues(xdm_file_path, frame_name)
        ComSignal_signals=Frame_Signals(xdm_file_path,frame_name)
        if ComIPduDirection==None and ComIPduSignalProcessing==None and ComIPduType==None and ComPduIdRef==None and ComIPduCallout==None and ComIPduCounter==None and ComIPduCounterSize==None and ComIPduCounterStartPosition==None and ComTxModeMode==None and ComTxModeTimePeriod==None:
            result_data = {
                'Signal Name': [frame_name],
                'Passed?':["Signal Not Found in COM"],
            }
            write_to_Excel(result_data,file_path,sheet_name)
            return False
        else:
            frames_data = cleanExcelFrameData(excel_file_path)
            signals_data = cleanExcelSignalData(excel_file_path)
            selected_frame = frames_data[frames_data['Radical'] == frame_name]
            if selected_frame.empty:
                result_data = {
                    'Signal Name': [frame_name],
                    'Passed?':["Signal Not Found in Messagerie"],
                }
                write_to_Excel(result_data,file_path,sheet_name)
            else:
                ComIPduDirectiontst=ComIPduSignalProcessingtst=ComIPduTypetst=ComPduIdReftst=signalstst=ComIPduCallouttst=ComIPduCountertst=ComIPduCounterSizetst=ComIPduCounterStartPositiontst=ComTxModeModetst=ComTxModeTimePeriodtst=True
                selected_signals = signals_data[signals_data['Radical_T'] == frame_name]
                if(ComIPduDirection!="RECEIVE" or ComIPduDirection!="SEND"):
                    ComIPduDirectiontst=False
                    
                if(ComIPduSignalProcessing!="DEFERRED"):
                    ComIPduSignalProcessingtst=False
                
                if(ComIPduType!="NORMAL"):
                    ComIPduTypetst=False
                
                if(frame_name not in ComPduIdRef):
                    ComPduIdReftst=False
                
                if(ComSignal_signals==signals):
                    signalstst=False

                chk_exist=cpt_exist=False
                signal_cpt=None
                for sig in selected_signals:
                    if sig["Nécessité de sécurisation par checksum et compteur de process"].values[0]=="CHK":
                        chk_exist=True
                    if sig["Nécessité de sécurisation par checksum et compteur de process"].values[0]=="CPT":
                        cpt_exist=True
                        signal_cpt=sig
                frame_id=selected_frame["Identifiant_T"].values[0]

                if(chk_exist):
                    if(ComIPduDirection=="SEND" and ComIPduCallout=="ISCAN_EveTxF"+ str(frame_id)+"_Callout"):
                        pass
                    elif(ComIPduDirection=="RECEIVE" and ComIPduCallout=="ISCAN_EveRxF"+ str(frame_id)+"_Callout"):
                        pass
                    else:
                        ComIPduCallouttst=False
                else:
                    ComIPduCallouttst=None

                if(cpt_exist):
                    if(ComIPduDirection=="SEND" and ComIPduCounter=="ISCAN_EveTxF"+ str(frame_id)+"_Callout"):
                        pass
                    elif(ComIPduDirection=="RECEIVE" and ComIPduCounter=="ISCAN_EveRxF"+ str(frame_id)+"_Callout"):
                        pass
                    else:
                        ComIPduCountertst=False
                    
                    if(ComIPduCounterSize!=signal_cpt["Taille_Max_S"].values[0]):
                        ComIPduCounterSizetst=False

                    pos_oct_excel=selected_signal["Position_octet_S"].values[0]
                    pos_bit_excel=selected_signal["Position_bit_S"].values[0]
                    pos_oct_com,pos_bit_com=Signal_Position_inFrame(ComBitPosition,ComBitSize)
                    if(ComIPduCounterStartPosition!=pos_bit_com):
                        ComIPduCounterStartPositiontst=False
                else:
                    ComIPduCountertst=None  
                    ComIPduCounterSizetst=None
                    ComIPduCounterStartPositiontst=None

                frame_mode_trans=selected_frame["Mode_Transmission_T"].values[0]
                if(ComIPduDirection=="SEND" and(frame_mode_trans=="Periodique" or frame_mode_trans=="Mixte" ) ):
                    if(ComTxModeMode!="PERIODIC"):
                        ComTxModeModetst=False
                    if(ComTxModeTimePeriod!=selected_frame["Offset"].values[0]):
                        ComTxModeTimePeriodtst=None#change this to False when sheet updates
                else:
                    ComTxModeModetst=None
                    ComTxModeTimePeriodtst=None
                result_data = {
                    'Signal Name': [frame_name],
                    'Passed?':[" " if ComIPduDirectiontst==False or ComIPduSignalProcessingtst==False or ComIPduTypetst==False or ComPduIdReftst==False or signalstst==False or ComIPduCallouttst==False or ComIPduCountertst==False or ComIPduCounterSizetst==False or ComIPduCounterStartPositiontst==False or ComTxModeModetst==False or ComTxModeTimePeriodtst==False else "X" ],
                    'ComIPduDirection':["Error(ComIPduDirection is neither of value 'SEND' or 'RECEIVE')" if ComIPduDirectiontst==False else ComIPduDirection],
                    'ComIPduSignalProcessing':["Error(ComIPduSignalProcessing is not of value 'DEFERRED')" if ComIPduSignalProcessingtst==False else ComIPduSignalProcessing],
                    'ComIPduType':["Error(ComIPduType is not of value 'NORMAL')" if ComIPduTypetst==False else ComIPduType],
                    'ComPduIdRef':["Error(Frame Name is not in ComPduIdRef )" if ComPduIdReftst==False else ComPduIdRef],
                    'ComIPduSignalRef':["Error(ComIPduSignalRef Signals are conform to the ComSignal signals)" if signalstst==False else signals],
                    'ComIPduCallout':["---" if ComIPduCallouttst==None else "Error(ComIPduCallout is not of the correct value)" if ComIPduCallouttst==False else ComIPduCallout],
                    'ComIPduCounter':["---" if ComIPduCountertst==None else "Error(ComIPduCountertst is not of the correct value)" if ComIPduCountertst==False else ComIPduCounter],
                    #adddddddd
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
root.title("Frame Info Verification in ComIPdu")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=browse_excel)
excel_file_button.grid(row=0, column=2, padx=5, pady=5)

xdm_file_label = tk.Label(frame, text="Select Com File:")
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
