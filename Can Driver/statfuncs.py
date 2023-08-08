import pandas as pd
import os
from openpyxl import load_workbook
from lxml import etree
import tkinter as tk
from tkinter import filedialog

#the excel output file file path
file_path = os.path.join(os.getcwd(), 'Output.xlsx')

#the namespace used for the XDM documents
namespace = {'d': 'http://www.tresos.de/_projects/DataModel2/06/data.xsd','a':'http://www.tresos.de/_projects/DataModel2/08/attribute.xsd'}

# Define expected headers for cleaning the Excel data
expected_headers = {'FRAMES': ['Checked','Radical', 'Activation trame', 'Protocole_M', 'Identifiant_T', 'Taille_Max_T', 'Lmin_T', 'Mode_Transmission_T', 'Nature_Evenement_FR_T', 'Nature_Evenement_GB_T', 'Periode_T', 'UCE Emetteur', 'AEE10r3 Reseau_T']}

# Function to clean the Excel data and keep only the necessary columns
def cleanExcelData(excel_file):
    df = pd.read_excel(excel_file, sheet_name='FRAMES', header=0)
    headers = [col for col in df.columns if col in expected_headers['FRAMES']]
    return df[headers]

#function responsible for writiting the output to the excel file
def write_to_Excel(result_data, file_path,sheet_name):
    df = pd.DataFrame(result_data)

    if not os.path.exists(file_path):
        # Create the Excel file with the specified columns
        df.to_excel(file_path, sheet_name=sheet_name, index=False, header=True)
    else:
        # Load the existing workbook
        book = load_workbook(file_path)
        writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        writer.book = book

        if sheet_name in pd.ExcelFile(file_path).sheet_names:
            # Check if the 'Passed?' column already exists in the sheet
            sheet = book[sheet_name]
            # Append the data to the existing sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=writer.sheets[sheet_name].max_row)

        else:
            # Create a new sheet if it doesn't exist
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

        writer.save()

# function responsable for emptying the designated excel sheet
def clear_excel(sheet_name):
    if os.path.exists(file_path):
        book = load_workbook(file_path)
        if sheet_name in book.sheetnames:
            sheet = book[sheet_name]
            sheet.delete_rows(2, sheet.max_row - 1) 
        book.save(file_path)
    

#Ordering by index table CanIF
def ordered_by_id_CanIf(xdm_file,order_var,parent):
    try:
        with open(xdm_file, 'r') as file:
            xdm_content = file.read()
        root = etree.fromstring(xdm_content)
        elements = root.xpath(f".//d:lst[@name='{parent}']/d:ctr", namespaces=namespace)
        frames_data = [(ctr.attrib['name'], ctr.xpath(f"string(d:var[@name='{order_var}']/@value)", namespaces=namespace)) for ctr in elements]
        frames_data = [(name, Id) for name, Id in frames_data if Id.strip()]
        first_Id = int(frames_data[0][1])
        if first_Id != 0:
            return "The first frame's ("+order_var+") should be (0)', but found ("+str(first_Id)+")."

        Ids = [int(Id) for _, Id in frames_data]
        if len(Ids) != len(set(Ids)):
            duplicates = [frame_name for frame_name, Id in frames_data if Ids.count(int(Id)) > 1]
            errorstring=""
            for frame_name in duplicates:
                errorstring=errorstring+" "+"The frame ("+frame_name+") has a duplicate ("+order_var+").\n"
            return errorstring

        Last_Id = int(frames_data[-1][1])
        total_frames = len(frames_data)
        if Last_Id != total_frames - 1:
            return "The last frame's ("+order_var+") should be ("+str(total_frames-1)+"), but found ("+str(Last_Id)+")."

        if any(int(frames_data[i - 1][1]) > int(frames_data[i][1]) for i in range(1, len(frames_data))):
            frame_name = frames_data[next(i for i in range(1, len(frames_data)) if int(frames_data[i - 1][1]) > int(frames_data[i][1]))][0]
            return "The frame ("+frame_name+") has a jump in ("+order_var+")."

        return True

    except Exception as e:
        print(f"Error: {e}")
        return False

#Ordering by index table PDUR
def ordered_by_id_PDUR(xdm_file,nodes):
    try:
        with open(xdm_file, 'r') as file:
            xdm_content = file.read()

        root = etree.fromstring(xdm_content)
        if(nodes[2]=='PduRDestPdu'):
            ctr_elements = root.xpath(f".//d:ctr[@name='{nodes[0]}']/d:lst[@name='{nodes[1]}']/d:ctr/d:lst[@name='{nodes[2]}']/d:ctr", namespaces=namespace)
            frames_data = [(ctr.attrib['name'], ctr.xpath(f"string(.//d:var[@name='{nodes[3]}']/@value)", namespaces=namespace)) for ctr in ctr_elements]
        else:
            elements = root.xpath(f".//d:ctr[@name='{nodes[0]}']/d:lst[@name='{nodes[1]}']/d:ctr", namespaces=namespace)
            frames_data = [(ctr.attrib['name'], ctr.xpath(f"string(.//d:ctr[@name='{nodes[2]}']/d:var[@name='{nodes[3]}']/@value)", namespaces=namespace)) for ctr in elements]

        frames_data = [(name, Id) for name, Id in frames_data if Id.strip()]
        first_Id = int(frames_data[0][1])
        if first_Id != 0:
            return "The first frame's ("+nodes[3]+") should be (0), but found ("+str(first_Id)+")."

        Ids = [int(Id) for _, Id in frames_data]
        if len(Ids) != len(set(Ids)):
            duplicates = [frame_name for frame_name, Id in frames_data if Ids.count(int(Id)) > 1]
            errorstring=""
            for frame_name in duplicates:
                errorstring=errorstring+"The frame ("+frame_name+") has a duplicate ("+nodes[3]+").\n"
            return errorstring

        Last_Id = int(frames_data[-1][1])
        total_frames = len(frames_data)
        if Last_Id != total_frames - 1:
            return "The last frame's ("+nodes[3]+") should be ("+str(total_frames-1)+"), but found ("+str(Last_Id)+")."

        if any(int(frames_data[i - 1][1]) > int(frames_data[i][1]) for i in range(1, len(frames_data))):
            frame_name = frames_data[next(i for i in range(1, len(frames_data)) if int(frames_data[i - 1][1]) > int(frames_data[i][1]))][0]
            return "The frame ("+frame_name+") has a jump in ("+nodes[3]+")."

        return True

    except Exception as e:
        print(f"Error: {e}")
        return False