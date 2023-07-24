#!/usr/bin/env python3
import os
import shutil
import datetime
import sys
import re
#### [ WARNING ] #####################################################
# Tool version to be update for each modification
ToolVersion = "01.00"
######################################################################

datestring = str(datetime.datetime.now())

LogFile = []  # Save Log file
LogFile.append("-----------------------------------------------------------\n")
LogFile.append("Generated on " + datestring + "\n")
LogFile.append("-----------------------------------------------------------\n")
LogFile.append("Update ComHandleId configuration for CAN signals (based on Tresos configuration)\n")
LogFile.append("ComHandleIdConfig version: " + ToolVersion + "\n\n")

file, dir = None, None
if len(sys.argv) > 1:
    file = sys.argv[1]
if len(sys.argv) > 2:
    dir = sys.argv[2]

if file is None:
    sys.exit("\n  --> ERROR: Com.xdm file argument is mandatory to continue\n")

if dir is None:
    # sys.exit("Need directory path\n")
    print("\n  --> WARNING: Use active directory path by default\n")
    dir = os.getcwd()

LogFile.append("#### [ Open XDM file ] ###############################################\n")
file = os.path.join(dir, file)

LogTool = os.path.join(dir, "ComHandleIdConfig.log")
if os.path.exists(LogTool):
    os.remove(LogTool)

def mainMode1():
    # BEGIN process
    try:
        with open(file, 'r') as in_file:
            content = in_file.readlines()
        os.chmod(file, 0o664)
    except FileNotFoundError:
        sys.exit("\n  --> ERROR: Could not open file for reading " + file + "\n")

    LogFile.append("  -> done\n")
    OutFile = []  # Save modified xdm file
    ComHandleId = 0
    for line in content:
        if "ComHandleId" in line:
            ComHandleId += 1
            line = line.replace(re.search(r'\d+', line).group(), str(ComHandleId))
        OutFile.append(line.strip())

    # END process
    LogFile.append("#### [ Update XDM file ] #############################################\n")

    # os.rename(file, file + ".old")
    shutil.move(file, file + ".old")
    try:
        with open(file, 'w') as out_file:
            out_file.write("\n".join(OutFile))
    except IOError:
        sys.exit("ERROR: Could not open file for writing " + file + "\n")
    LogFile.append("  -> done\n")
    LogFile.append("#### [ End process ] #################################################\n")

    # Write Log file
    try:
        with open(LogTool, 'w') as log_file:
            log_file.write("\n".join(LogFile))
    except IOError:
        sys.exit("ERROR: Can't create file for writing ComHandleIdConfig.log\n")

def mainMode2():
    # BEGIN process
    try:
        with open(file, 'r') as in_file:
            content = in_file.readlines()
        os.chmod(file, 0o664)
    except FileNotFoundError:
        sys.exit("\n  --> ERROR: Could not open file for reading " + file + "\n")

    LogFile.append("  -> done\n")
    OutFile = []  # Save modified xdm file
    OutComHandleId = []  # Save ComHandleId in array
    ID = 0
    previousID = 0
    warningNb = 0
    errorNb = 0

    LogFile.append("#### [ Checking XDM file ] ###########################################\n")
    LogFile.append("#### [ ID consistency ] ##############################################\n")
    for line in content:
        if "ComHandleId" in line:
            ID = int(re.search(r'\d+', line).group())
            OutComHandleId.append(ID)
            if ((ID != 0 and previousID >= ID) or (ID - previousID) > 1):
                warningNb += 1
                LogFile.append("  --> WARNING: ID consistency problem identified in line " + str(len(OutFile) + 1) + "\n")
            previousID = ID
        OutFile.append(line.strip())

    LogFile.append("  -> done\n")
    LogFile.append("#### [ Multiple definition ] #########################################\n")
    if warningNb != 0:
        count = {}
        for element in OutComHandleId:
            count[element] = count.get(element, 0) + 1
        for element, count in count.items():
            if count > 1:
                errorNb += 1
                LogFile.append("  --> ERROR: Multiple definition of ID [" + str(element) + "] : " + str(count) + "\n")
    LogFile.append("  -> done\n")
    # END process

    LogFile.append("#### [ General report ] ##############################################\n")
    if warningNb == 0:
        LogFile.append("  Problems detected :\n")
        LogFile.append("     No error detected\n")
        LogFile.append("     No warning detected\n")
    else:
        LogFile.append("  Problems detected :\n")
        if errorNb == 0:
            LogFile.append("     No error detected\n")
        else:
            if errorNb == 1:
                LogFile.append("     1 error detected\n")
            else:
                LogFile.append("     " + str(errorNb) + " errors detected\n")
        if warningNb == 1:
            LogFile.append("     1 warning detected\n")
        else:
            LogFile.append("     " + str(warningNb) + " warnings detected\n")
    LogFile.append("#### [ End process ] #################################################\n")

    # Write Log file
    try:
        with open(LogTool, 'w') as log_file:
            log_file.write("\n".join(LogFile))
    except IOError:
        sys.exit("ERROR: Can't create file for writing ComHandleIdConfig.log\n")


print("#### [ Start process ] #################################################\n")
print("  Help: ARG1 - Mandatory : Com.xdm file\n")
print("        ARG2 - Optional  : Process directory path\n\n")
print("#### [ Update/Check ComHandleId for CAN signals ] ######################\n")
print("  --> Please choose an option : \n")
print("        [1] Update and Checking ComHandleId of CAN signals\n")
print("        [2] Update ComHandleId of CAN signals\n")
print("        [3] Checking an existing Tresos configuration\n")
print("        [4] Exit\n\n")
entreeUser = input("  --> Choice : ").strip()

if entreeUser == "1":
    mainMode1()
    mainMode2()
elif entreeUser == "2":
    mainMode1()
elif entreeUser == "3":
    mainMode2()
elif entreeUser == "4":
    pass
else:
    print("  --> WARNING: Please enter a valid option\n")

print("#### [ End process ] ###################################################\n")
