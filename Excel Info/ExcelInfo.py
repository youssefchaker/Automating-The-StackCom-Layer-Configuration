import pandas as pd

expected_headers = {
    'FRAMES': ['Radical', 'Activation trame', 'Protocole_M', 'Identifiant_T', 'Taille_Max_T', 'Lmin_T', 'Mode_Transmission_T', 'Nature_Evenement_FR_T', 'Nature_Evenement_GB_T', 'Periode_T', 'UCE Emetteur', 'AEE10r3 Reseau_T'],
    'SIGNALS': ['Radical_T', 'Position_octet_S', 'Position_bit_S', 'Taille_Max_S', 'Mnemonique_S', 'Nom_FR_S', 'Nom_GB_S', 'Type_S', 'Nom_FR_V', 'Nom_GB_V', 'Unite_FR_U', 'Unite_GB_U', 'Valeur_Min_S', 'Valeur_Max_S', 'Resolution_S', 'Offset_S', 'Valeur_Invalide_S', 'Valeur_Indisponible_S', 'Valeur_Interdite_S', 'Emetteur', 'Récepteur', 'Gateway']
}

def cleanExcelData(excel_file):
    frames_data = {}
    signals_data = {}
    
    for sheet_name, expected_header_names in expected_headers.items():
        # Read the sheet into a DataFrame, specifying the header row
        df = pd.read_excel(excel_file, sheet_name, header=0)

        # Filter columns based on expected headers
        headers = [col for col in df.columns if col in expected_header_names]

        # Store the data for each column
        data = df[headers].values.tolist()

        # Store the data in variables
        if sheet_name == 'FRAMES':
            frames_data = {header: col_data for header, col_data in zip(headers, zip(*data))}
        elif sheet_name == 'SIGNALS':
            signals_data = {header: col_data for header, col_data in zip(headers, zip(*data))}
    
    return frames_data, signals_data

# Read the Excel file
excel_file = r'C:\Users\Youssefch\Desktop\studies\capgemini\py\Excel Info\Frames&Signals.xlsx'

# Clean the data and retrieve frames and signals data
frames_data, signals_data = cleanExcelData(excel_file)

# Print frames data
print("Frames Data:")
for header, data in frames_data.items():
    print(f"{header}: {data}\n")

print("///////////////////////////////////////////////////\n")

# Print signals data
print("Signals Data:")
for header, data in signals_data.items():
    print(f"{header}: {data}\n")
