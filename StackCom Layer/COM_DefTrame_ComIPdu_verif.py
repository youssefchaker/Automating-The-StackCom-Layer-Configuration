import statfuncs
from statfuncs import *

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
        ComIPduCallout = elements[0].xpath("d:var[@name='ComIPduCallout']/@value", namespaces=namespace)
        ComIPduCallout = list(filter(None, ComIPduCallout))
        if(ComIPduCallout):
            ComIPduCallout=ComIPduCallout[0]
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

def get_CptSignal_Position_Size(signal_name,xdm_file):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()

    root = etree.fromstring(xdm_content)
    elements = root.xpath(".//d:lst[@name='ComSignal']/d:ctr[@name=$name]", namespaces=namespace, name=signal_name)
    if elements:
        ComBitPosition=int(elements[0].xpath("d:var[@name='ComBitPosition']/@value", namespaces=namespace)[0])
        ComBitSize = int(elements[0].xpath("d:var[@name='ComBitSize']/@value", namespaces=namespace)[0])
    else:
        ComBitPosition=ComBitSize=None

    return ComBitPosition,ComBitSize


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
                'Frame Name': [frame_name],
                'Passed?':["Frame Not Found in COM"],
                'ComIPduDirection':' ',
                'ComIPduSignalProcessing':' ',
                'ComIPduType':' ',
                'ComPduIdRef':' ',
                'ComIPduSignalRef':' ',
                'ComIPduCallout':' ',
                'ComIPduCounter':' ',
                'ComIPduCounterSize':' ',
                'MessagerieIPduCounterSize':' ',
                'ComIPduCounterSize/MessagerieIPduCounterSize':' ',
                'ComIPduCounterStartPosition':' ',
                'MessagerieIPduCounterStartPosition':' ',
                'ComIPduCounterStartPosition/MessagerieIPduCounterStartPosition':' ',
                'ComTxModeMode':' ',
                'ComTxModeTimePeriod':' '
                }
            write_to_Excel(result_data,file_path,sheet_name)
            return False
        else:
            frames_data = cleanExcelFrameData(excel_file_path)
            signals_data = getFullSignalData(excel_file_path)
            selected_frame = frames_data[frames_data['Radical'] == frame_name]
            if selected_frame.empty:
                result_data = {
                    'Frame Name': [frame_name],
                    'Passed?':["Frame Not Found in Messagerie"],
                    'ComIPduDirection':' ',
                    'ComIPduSignalProcessing':' ',
                    'ComIPduType':' ',
                    'ComPduIdRef':' ',
                    'ComIPduSignalRef':' ',
                    'ComIPduCallout':' ',
                    'ComIPduCounter':' ',
                    'ComIPduCounterSize':' ',
                    'MessagerieIPduCounterSize':' ',
                    'ComIPduCounterSize/MessagerieIPduCounterSize':' ',
                    'ComIPduCounterStartPosition':' ',
                    'MessagerieIPduCounterStartPosition':' ',
                    'ComIPduCounterStartPosition/MessagerieIPduCounterStartPosition':' ',
                    'ComTxModeMode':' ',
                    'ComTxModeTimePeriod':' '
                }
                write_to_Excel(result_data,file_path,sheet_name)
            else:
                ComIPduDirectiontst=ComIPduSignalProcessingtst=ComIPduTypetst=ComPduIdReftst=signalstst=ComIPduCallouttst=ComIPduCountertst=ComIPduCounterSizetst=ComIPduCounterStartPositiontst=ComTxModeModetst=ComTxModeTimePeriodtst=True
                selected_signals = signals_data[signals_data['Radical_T'] == frame_name]
                if(ComIPduDirection=="RECEIVE" or ComIPduDirection=="SEND"):
                    pass
                else:
                    ComIPduDirectiontst=False
                    
                if(ComIPduSignalProcessing!="DEFERRED"):
                    ComIPduSignalProcessingtst=False
                
                if(ComIPduType!="NORMAL"):
                    ComIPduTypetst=False
                
                if(frame_name not in ComPduIdRef):
                    ComPduIdReftst=False
                
                if(ComSignal_signals!=signals):
                    signalstst=False

                chk_exist=cpt_exist=False
                signal_cpt=None
                for sig in selected_signals.iterrows():
                    if sig[1][13]=="CHK":
                        chk_exist=True
                    if sig[1][13]=="CPT":
                        cpt_exist=True
                        signal_cpt=sig

                frame_id=selected_frame["Identifiant_T"].values[0]
                Tx_test=selected_frame["UCE Emetteur"].str.endswith("E_VCU").any()

                if(chk_exist):
                    if(not ComIPduCallout):
                        ComIPduCallout="Null"
                        ComIPduCallouttst=False
                    elif(Tx_test and ComIPduCallout=="ISCAN_EveTxF"+ str(frame_id)+"_Callout"):
                        pass
                    elif(not Tx_test and ComIPduCallout=="ISCAN_EveRxF"+ str(frame_id)+"_Callout"):
                        pass
                    else:
                        ComIPduCallouttst=False
                else:
                    ComIPduCallouttst=None

                if(cpt_exist):
                    if(not ComIPduCounter):
                        ComIPduCounter="Null"
                        ComIPduCountertst=False
                    elif(Tx_test and ComIPduCounter=="ISCAN_EveTxF"+ str(frame_id)+"_Callout"):
                        pass
                    elif(not Tx_test and ComIPduCounter=="ISCAN_EveRxF"+ str(frame_id)+"_Callout"):
                        pass
                    else:
                        ComIPduCountertst=False
                    if(not ComIPduCounterSize):
                        ComIPduCounterSize="Null"
                        ComIPduCounterSizetst=False
                    elif(int(ComIPduCounterSize)!=int(signal_cpt[1][3])):
                        ComIPduCounterSizetst=False
                    pos_bit_excel=signal_cpt[1][3]
                    ComBitPosition,ComBitSize=get_CptSignal_Position_Size(signal_cpt[1][5]+"_"+signal_cpt[1][1],xdm_file_path)
                    pos_oct_com,pos_bit_com=Signal_Position_inFrame(ComBitPosition,ComBitSize)
                    if(pos_bit_excel!=pos_bit_com):
                        ComIPduCounterStartPositiontst=False
                else:
                    ComIPduCountertst=None  
                    ComIPduCounterSizetst=None
                    ComIPduCounterStartPositiontst=None

                frame_mode_trans=selected_frame["Mode_Transmission_T"].values[0]

                if(Tx_test and(frame_mode_trans=="Periodique" or frame_mode_trans=="Mixte" or frame_mode_trans=="Périodique" ) ):
                    if(not ComTxModeMode ):
                        ComTxModeMode="Null"
                        ComTxModeModetst=False
                    elif(ComTxModeMode!="PERIODIC"):
                        ComTxModeModetst=False
                    if(not ComTxModeTimePeriod):
                        ComTxModeTimePeriod="Null"
                    #if(ComTxModeTimePeriod!=selected_frame["Offset"].values[0]): #uncomment this 
                    ComTxModeTimePeriodtst=None#change this to False
                else:
                    ComTxModeModetst=None
                    ComTxModeTimePeriodtst=None

                result_data = {
                    'Frame Name': [frame_name],
                    'Passed?':["NOK" if ComIPduDirectiontst==False or ComIPduSignalProcessingtst==False or ComIPduTypetst==False or ComPduIdReftst==False or signalstst==False or ComIPduCallouttst==False or ComIPduCountertst==False or ComIPduCounterSizetst==False or ComIPduCounterStartPositiontst==False or ComTxModeModetst==False or ComTxModeTimePeriodtst==False else "OK" ],
                    'ComIPduDirection':["Error(ComIPduDirection is neither of value 'SEND' or 'RECEIVE')" if ComIPduDirectiontst==False else ComIPduDirection],
                    'ComIPduSignalProcessing':["Error(ComIPduSignalProcessing is not of value 'DEFERRED')" if ComIPduSignalProcessingtst==False else ComIPduSignalProcessing],
                    'ComIPduType':["Error(ComIPduType is not of value 'NORMAL')" if ComIPduTypetst==False else ComIPduType],
                    'ComPduIdRef':["Error(Frame Name is not in ComPduIdRef )" if ComPduIdReftst==False else ComPduIdRef],
                    'ComIPduSignalRef':["Error(ComIPduSignalRef Signals are not conform to the ComSignal signals)" if signalstst==False else signals],
                    'ComIPduCallout':["---" if ComIPduCallouttst==None else "Error(ComIPduCallout is not of the correct value)" if ComIPduCallouttst==False else ComIPduCallout],
                    'ComIPduCounter':["---" if ComIPduCountertst==None else "Error(ComIPduCounter is not of the correct value)" if ComIPduCountertst==False else ComIPduCounter],
                    'ComIPduCounterSize':["---" if ComIPduCounterSizetst==None else ComIPduCounterSize],
                    'MessagerieIPduCounterSize':["---" if ComIPduCounterSizetst==None else signal_cpt[1][3]],
                    'ComIPduCounterSize/MessagerieIPduCounterSize':["---" if ComIPduCounterSizetst==None else "Error(CPT Signal Size Mismatch)" if ComIPduCounterSizetst==False else "None"],
                    'ComIPduCounterStartPosition':["---" if ComIPduCounterStartPositiontst==None else pos_bit_com],
                    'MessagerieIPduCounterStartPosition':["---" if ComIPduCounterStartPositiontst==None else pos_bit_excel],
                    'ComIPduCounterStartPosition/MessagerieIPduCounterStartPosition':["---" if ComIPduCounterStartPositiontst==None else "Error(CPT Signal Position in Frame Mismatch)" if ComIPduCounterStartPositiontst==False else "None"],
                    'ComTxModeMode':["---" if ComTxModeModetst==None else "Error(ComTxModeMode is not of the value 'PERIODIC')" if ComTxModeModetst==False else ComTxModeMode],
                    'ComTxModeTimePeriod':["---" if ComTxModeTimePeriodtst==None else "Error(ComTxModeTimePeriod is not of the correct value)" if ComTxModeTimePeriodtst==False else ComTxModeTimePeriod]#change this
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
root.title("Com Frame Info Verification in ComIPdu")

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

xdm_file_label = tk.Label(frame, text="Select Com File:")
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
