import statfuncs
from statfuncs import *

sheet_name="COM_DefSignal_ComSignal_verif"

# Function to extract necessary attributes for the target signal from the .xdm file
def extract_ComValues(xdm_file, signal_name):
    with open(xdm_file, 'r') as file:
        xdm_content = file.read()

    root = etree.fromstring(xdm_content)
    elements = root.xpath(".//d:lst[@name='ComSignal']/d:ctr[@name=$name]", namespaces=namespace, name=signal_name)
    if elements:
        ComBitPosition=int(elements[0].xpath("d:var[@name='ComBitPosition']/@value", namespaces=namespace)[0])
        ComBitSize = int(elements[0].xpath("d:var[@name='ComBitSize']/@value", namespaces=namespace)[0])
        ComSignalEndianness = elements[0].xpath("string(d:var[@name='ComSignalEndianness']/@value)", namespaces=namespace)
        ComSignalInitValue = elements[0].xpath("string(d:var[@name='ComSignalInitValue']/@value)", namespaces=namespace)
        ComSignalType = elements[0].xpath("d:var[@name='ComSignalType']/@value", namespaces=namespace)[0]
        ComTransferProperty = elements[0].xpath("d:var[@name='ComTransferProperty']/@value", namespaces=namespace)
        ComTransferProperty = list(filter(None, ComTransferProperty))
        if(ComTransferProperty):
            ComTransferProperty=ComTransferProperty[0]
        
        ComNotification=elements[0].xpath("d:var[@name='ComNotification']/@value", namespaces=namespace)
        ComNotification = list(filter(None, ComNotification))
        if(ComNotification):
            ComNotification=ComNotification[0]
        ComTimeoutNotification=elements[0].xpath("d:var[@name='ComTimeoutNotification']/@value", namespaces=namespace)
        ComTimeoutNotification = list(filter(None, ComTimeoutNotification))
        if(ComTimeoutNotification):
            ComTimeoutNotification=ComTimeoutNotification[0]
        ComTimeout=elements[0].xpath("d:var[@name='ComTimeout']/@value", namespaces=namespace)
        ComTimeout = list(filter(None, ComTimeout))
        if(ComTimeout):
            ComTimeout=float(ComTimeout[0])
    else:
        ComBitPosition= ComBitSize= ComSignalEndianness= ComSignalInitValue= ComSignalType= ComTransferProperty= ComNotification= ComTimeoutNotification=ComTimeout= None

    return ComBitPosition, ComBitSize, ComSignalEndianness, ComSignalInitValue, ComSignalType, ComTransferProperty, ComNotification, ComTimeoutNotification, ComTimeout


def verify_signal(excel_file_path,xdm_file_path, signal_name):
    try:
        ComBitPosition, ComBitSize, ComSignalEndianness, ComSignalInitValue, ComSignalType, ComTransferProperty, ComNotification, ComTimeoutNotification, ComTimeout= extract_ComValues(xdm_file_path, signal_name)
        if ComBitPosition==None and ComBitSize==None and ComSignalEndianness==None and ComSignalInitValue==None and ComSignalType==None and ComTransferProperty==None and ComNotification==None and ComTimeoutNotification==None and ComTimeout==None:
            result_data = {
                'Signal Name': [signal_name],
                'Passed?':["Signal Not Found in COM"],
                'Frame Name': ' ',
                'Frame Type': ' ',
                'ComBitPosition': ' ',
                'MessagerieBitPosition': ' ',
                'ComBytePosition': ' ',
                'MessagerieBytePosition': ' ',
                'ComBitPosition/ComBytePosition Errors': ' ',
                'ComBitSize': ' ',
                'MessagerieBitSize': ' ',
                'ComBitSize/MessagerieBitSize Errors': ' ',
                'ComSignalEndianness': ' ',
                'ComSignalInitValue': ' ',
                'MessagerieSignalInitValue': ' ',
                'ComSignalInitValue/MessagerieSignalInitValue Errors': ' ',
                'ComSignalType': ' ',
                'MessagerieSignalType': ' ',
                'ComSignalType/MessagerieSignalType': ' ',
                'ComTransferProperty': ' ',
                'MessagerieTransferProperty': ' ',
                'ComTransferProperty/MessagerieTransferProperty': ' ',
                'ComNotification': ' ',
                'ComTimeoutNotification': ' ',
                'ComTimeout': ' ',
                'MessagerieTimeout': ' ',
                'ComTimeout/MessagerieTimeout': ' '
            }
            write_to_Excel(result_data,file_path,sheet_name)
            return False
        else:
            signals_data = cleanExcelSignalData(excel_file_path)
            frames_data = getFullFrameData(excel_file_path)
            selected_signal = signals_data[signals_data['Mnemonique_S']+"_"+signals_data['Radical_T'] == signal_name]
            if selected_signal.empty:
                result_data = {
                    'Signal Name': [signal_name],
                    'Passed?':["Signal Not Found in Messagerie"],
                    'Frame Name': ' ',
                    'Frame Type': ' ',
                    'ComBitPosition': ' ',
                    'MessagerieBitPosition': ' ',
                    'ComBytePosition': ' ',
                    'MessagerieBytePosition': ' ',
                    'ComBitPosition/ComBytePosition Errors': ' ',
                    'ComBitSize': ' ',
                    'MessagerieBitSize': ' ',
                    'ComBitSize/MessagerieBitSize Errors': ' ',
                    'ComSignalEndianness': ' ',
                    'ComSignalInitValue': ' ',
                    'MessagerieSignalInitValue': ' ',
                    'ComSignalInitValue/MessagerieSignalInitValue Errors': ' ',
                    'ComSignalType': ' ',
                    'MessagerieSignalType': ' ',
                    'ComSignalType/MessagerieSignalType': ' ',
                    'ComTransferProperty': ' ',
                    'MessagerieTransferProperty': ' ',
                    'ComTransferProperty/MessagerieTransferProperty': ' ',
                    'ComNotification': ' ',
                    'ComTimeoutNotification': ' ',
                    'ComTimeout': ' ',
                    'MessagerieTimeout': ' ',
                    'ComTimeout/MessagerieTimeout': ' '
                }
                write_to_Excel(result_data,file_path,sheet_name)
            else:
                selected_frame = frames_data[frames_data['Radical'] == selected_signal['Radical_T'].values[0]]
                identifiant_t_hex = selected_frame["Identifiant_T"].values[0]
                frame_id=int(identifiant_t_hex,16)
                ComBitPositiontst= ComBitSizetst= ComSignalEndiannesstst= ComSignalInitValuetst= ComSignalTypetst= ComTransferPropertytst= ComNotificationtst= ComTimeoutNotificationtst=ComTimeouttst= True
                #excel signal info
                pos_oct_excel=selected_signal["Position_octet_S"].values[0]
                pos_bit_excel=selected_signal["Position_bit_S"].values[0]
                val_min_excel=selected_signal["Valeur_Min_S"].values[0]
                val_max_excel=selected_signal["Valeur_Max_S"].values[0]
                taille_excel=selected_signal["Taille_Max_S"].values[0]
                resolution_excel=selected_signal["Resolution_S"].values[0]
                offset_excel=selected_signal["Offset_S"].values[0]

                    
                pos_oct_com,pos_bit_com=Signal_Position_inFrame(ComBitPosition,ComBitSize)
                signal_type_excel=Signal_Type(taille_excel,val_min_excel,val_max_excel,resolution_excel,offset_excel)

                if(pos_oct_com!=pos_oct_excel or pos_bit_com!=pos_bit_excel):
                    ComBitPositiontst=False
                
                if(taille_excel!=ComBitSize):
                    ComBitSizetst=False
                
                if(ComSignalEndianness!="BIG_ENDIAN"):
                    ComSignalEndiannesstst=False

                signal_Rx_test=not (selected_signal["Emetteur"].str.endswith("E_VCU").any())
                signal_init_value_excel=hex(0)
                if signal_Rx_test:
                    signal_init_value_excel=selected_signal["CONS_INIT"].values[0]
                else:
                    signal_init_value_excel=selected_signal["PROD_INIT"].values[0]

                if(signal_init_value_excel=="Non applicable" or signal_init_value_excel==" " or signal_init_value_excel=='Null'):  
                        signal_init_value_excel="Null"
                        ComSignalInitValuetst=False
                elif(signal_init_value_excel==0 or signal_init_value_excel=="0"):
                    signal_init_value_excel=hex(0)
                elif(int(signal_init_value_excel,16)!=int(ComSignalInitValue)):
                        ComSignalInitValuetst=False
                else:
                    signal_init_value_excel=int(signal_init_value_excel,16)

                
                signal_modetrans=selected_frame["Mode_Transmission_T"].values[0]
                if(not ComTransferProperty):
                    ComTransferProperty="Null"
                    ComTransferPropertytst=False
                elif(signal_modetrans=="Periodique" or signal_modetrans=="Périodique"  and ComTransferProperty=="PENDING"):
                    pass
                elif (signal_modetrans=="Evenmentielle" and ComTransferProperty=="TRIGGERED"):
                    pass
                elif (signal_modetrans=="Mixte" and ComTransferProperty=="TRIGGERED_ON_CHANGE"):
                    pass
                else:
                    ComTransferPropertytst=False
                
                if(not ComNotification):
                    ComNotification="Null"
                    ComNotificationtst=False
                elif(signal_Rx_test and ComNotification!="FHCAN_EveRxF"+str(frame_id)+"_AckClbk"):
                    ComNotificationtst=False

                if(not ComTimeoutNotification):
                    ComTimeoutNotification="Null"  
                    ComNotificationtst=False
                elif(signal_Rx_test and ComTimeoutNotification!="FHCAN_EveRxF"+str(frame_id)+"_TOutClbk"):
                    ComTimeoutNotificationtst=False

                signal_period_excel=selected_frame["Periode_T"].values[0]
                if(not ComTimeout or signal_period_excel=="Non applicable"):
                    ComTimeout="Null"
                    ComTimeouttst=False
                elif(signal_modetrans=="Periodique" or signal_modetrans=="Périodique" and signal_Rx_test):
                    if(signal_period_excel==10 and ComTimeout>=3*signal_period_excel):
                        pass
                    if(signal_period_excel>=20 and signal_period_excel<=30 and ComTimeout>=2*signal_period_excel):
                        pass
                    if(signal_period_excel>=40 and signal_period_excel<=90 and ComTimeout>=signal_period_excel+10):
                        pass
                    if(signal_period_excel>=100 and ComTimeout>=signal_period_excel+(signal_period_excel*10/100)):
                        pass
                    else:
                        ComTimeouttst=False
                result_data = {
                    'Signal Name': [signal_name],
                    'Passed?':[" " if ComBitPositiontst == False or ComBitSizetst == False or ComSignalEndiannesstst == False or ComSignalInitValuetst == False or ComSignalTypetst == False or ComTransferPropertytst == False or ComNotificationtst == False or ComTimeoutNotificationtst == False or ComTimeouttst == False else "X"],
                    'Frame Name':[selected_frame["Radical"].values[0]],
                    'Frame Type': ["Tx" if selected_signal["Emetteur"].str.endswith("E_VCU").any() else "Rx"],
                    'ComBitPosition':[pos_bit_com],
                    'MessagerieBitPosition':[pos_bit_excel],
                    'ComBytePosition':[pos_oct_com],
                    'MessagerieBytePosition':[pos_oct_excel],
                    'ComBitPosition/ComBytePosition Errors':["Error(Bit/Byte Mismatch)" if ComBitPositiontst==False else "None"],
                    'ComBitSize':[ComBitSize],
                    'MessagerieBitSize':[taille_excel],
                    'ComBitSize/MessagerieBitSize Errors':["Error(Signal Size Mismatch)" if ComBitSizetst==False else 'None'],
                    'ComSignalEndianness':["Error(ComSignalEndianness is not of the value 'BIG_ENDIAN')" if ComSignalEndiannesstst==False else ComSignalEndianness],
                    'ComSignalInitValue':[ComSignalInitValue],
                    'MessagerieSignalInitValue':[signal_init_value_excel],
                    'ComSignalInitValue/MessagerieSignalInitValue Errors':["Error(Signal Init Value Mismatch)" if ComSignalInitValuetst==False else "None"],
                    'ComSignalType':[ComSignalType],
                    'MessagerieSignalType':[signal_type_excel],
                    'ComSignalType/MessagerieSignalType':["Error(Signal Type Mismatch)" if ComSignalTypetst==False else "None"],
                    'ComTransferProperty':[ComTransferProperty],
                    'MessagerieTransferProperty':[signal_modetrans],
                    'ComTransferProperty/MessagerieTransferProperty':["Error(Transfer Property Mismatch)" if ComTransferPropertytst==False else "None"],
                    'ComNotification':["---" if signal_Rx_test==False else "Error(ComNotification is not in the correct form)" if signal_Rx_test and ComNotificationtst==False else ComNotification],
                    'ComTimeoutNotification':["---" if signal_Rx_test==False else "Error(ComTimeoutNotification is not in the correct form)" if signal_Rx_test and ComTimeoutNotificationtst==False else ComTimeoutNotification],
                    'ComTimeout':["---" if signal_Rx_test==False else ComTimeout],
                    'MessagerieTimeout':["---" if signal_Rx_test==False else signal_period_excel],
                    'ComTimeout/MessagerieTimeout':["---" if signal_Rx_test==False else "Error(ComTimeout/MessagerieTimeout values are not Correct)" if signal_Rx_test and ComTimeouttst==False else"None"],
                }
                write_to_Excel(result_data,file_path,sheet_name)

    except Exception as e:
                print(f"Error occurred : {e}")
                return False   

#execute functionality on button click
def verify_button_click():
    excel_file_path = excel_file_entry.get()
    xdm_file_path = xdm_file_entry.get()
    signal_name = signal_entry.get()

    verify_signal(excel_file_path, xdm_file_path, signal_name)
    completion_label.config(text="Output Created", fg="green")

# tkinter Interface
root = tk.Tk()
root.title("Com Signal Info Verification in ComSignal")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

excel_file_label = tk.Label(frame, text="Select Excel File:")
excel_file_label.grid(row=0, column=0, padx=5, pady=5)

excel_file_entry = tk.Entry(frame)
excel_file_entry.grid(row=0, column=1, padx=5, pady=5)

signal_label = tk.Label(frame, text="Enter Signal Name:")
signal_label.grid(row=2, column=0, padx=5, pady=5)

signal_entry = ttk.Combobox(frame)
signal_entry.grid(row=2, column=1, padx=5, pady=5)

excel_file_button = tk.Button(frame, text="Browse", command=lambda:browse_excel_signals(excel_file_entry,signal_entry))
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
