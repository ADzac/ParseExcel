import xlrd #pip install xlrd
import csv
import glob


def parse(filename): #open and get the infos
    loc = (filename)
    cpt=0
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    list_email=[]
    list_phone=[]

    for i in range(sheet.nrows):
        if(sheet.cell_value(i, 0)!='Email' and sheet.cell_value(i, 0)!='Email Data' and sheet.cell_value(i, 0)!='' and cpt==0):
            if(sheet.cell_value(i, 0)=='Phone Data'):
                cpt=1
            else:
                list_email.append(sheet.cell_value(i, 0))
            
        if(sheet.cell_value(i, 0)!='Phone Data' and sheet.cell_value(i, 0)!='Phone' and cpt==1 and sheet.cell_value(i, 0)!='Twitter Data' and sheet.cell_value(i, 0)!='Twitter Handle'):
            list_phone.append(sheet.cell_value(i, 0))

    email=tuple(list_email) #put in a seperate tuple for email and phone
    phone=tuple(list_phone)
    
    return (email,phone)

def append_csv(csv_file, list): #use all the info taken from all xls file and
    with open(csv_file, 'r') as read_file:
        reader = csv.DictReader(read_file)
        with open (csv_file, 'a', newline='') as append_file:
            writer = csv.DictWriter(append_file, reader.fieldnames)
            for email in list[0]:
                writer.writerow({'Email' : email})
            for phone in list[1]:
                writer.writerow({'Phone' : phone})

import win32com.client as win32 #pip install pywin32 (note that this works only in win11 for some reason)
def saveExcel(): #save the file so it is readable and the most important part
    # getting excel files from selected Directory 
    path = "C:\\Users\\A N I Q\\AppliRT\\Stage\\"

    # read all the files with extension .xls i.e. excel 
    filenames = glob.glob(path + "\*.xls")
    
    for fname in filenames:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(fname)
        wb.SaveAs(fname, FileFormat = 56)  # 56 for .xls file and 51 for .xlsx file
        wb.Close()                               
        excel.Application.Quit()
        y=parse(fname)
        append_csv('C:\\Users\\A N I Q\\AppliRT\\Stage\\google-ads_bdd.csv',y)
