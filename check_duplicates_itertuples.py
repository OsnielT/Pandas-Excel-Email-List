###############################################################################################
#
#  This is a script to compare optout emails from KEAP to DDI emails from Customer level
#   [*] The script will compare emails, if is a match then it will keep the Customer 
#       numbe, Name, and Email
#   [*] When checking finished, script will output results to excel named Results.xlsx
#
#   CMD CALL:   python check_duplicates_itertuples.py
#
###############################################################################################
import pandas as pd
from openpyxl import Workbook
import os, sys
from progress.bar import ChargingBar

def checkIfExist(Name, OPTOUT_EMAILS_FILE):
    
    # reads optout email list for loop to compare with DDI emails

    try:
        with open(OPTOUT_EMAILS_FILE + Name, "r+"):
            return Name
    except IOError:
        print("")
        print("*** Error: File is open or does not exists! ***\n")
        print("File you spcified: \n"+ Name +"\n")
        print("Folder path: \n" + OPTOUT_EMAILS_FILE+"\n")
        print("!    Make sure file exists in Excel_A folder    !")
        print("!        Make sure file name is correct         !")
        sys.exit()

    return fileName


os.system('cls')

print("Enter file name from \"Excel_A\" folder of optout emails: ")

# File path to email lists
OPTOUT_EMAILS_FILE  = "D:\\Projects\\Python-Excel-ReadWrite\\Excel_A\\" 
fileName = checkIfExist(str(input()), OPTOUT_EMAILS_FILE)

OPTOUT_EMAILS_FILE  = OPTOUT_EMAILS_FILE + fileName

# File path of email to compare too
TARGET_PATH_FILE    = "Excel_A/tests_list.csv"

# Where generated result weill be saved
RESULT_PATH_FILE    = "Results/Results.xlsx"

optOutEmail     = pd.read_csv(OPTOUT_EMAILS_FILE)
optOutEmailHead = optOutEmail.head()

# reads email list from DDI for compare with optout emails
targetFile      = pd.read_csv(TARGET_PATH_FILE)
targetFileHead  = targetFile.head()

# array for results to hold matching emails
rows={
    'Customer Number':[], 
    'Company':[],
    'Primary Email':[]
}

# number of rows in optout excel
optoutLeng = int(optOutEmail.shape[0])

os.system('cls')    #clear console

print("\nChecking Emails...\n") 

# CHECK IF EMAIL EXISTS
with ChargingBar('Processing', max=optoutLeng) as bar:
    for row_a in optOutEmail.itertuples():
        for row_b in targetFile.itertuples():
            if (row_a[3] == row_b[3]) and (row_a[2] == row_b[2]):
                # row_a[2] #optout: companmy name
                # row_a[3] #optout: email
                rows['Customer Number'].append(row_b[1]) #target: customer number
                rows['Company'].append(row_b[2])         #target: company name
                rows['Primary Email'].append(row_a[3])   #target: email
        bar.next() #increments loading bar

# generates excel sheet from Results array
result  = pd.ExcelFile(RESULT_PATH_FILE)
df1     = pd.DataFrame(data=rows)
df1.to_excel(result, sheet_name= 'results', header = True, index=False )

print("\nDone.\n")


