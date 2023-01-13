from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# CREATE A DICTIONARY WITH KEY AS CUSTOMER ID, AND VALUE AS LIST WITH CUSTOMER 
# INFORMATION.  FIRST ITEM HAS KEY "ID" AND VALUE AS LIST WITH ALL COLUMN IDs


datadump1_organized = {}
datadump2_organized = {}


wb = load_workbook("DATADUMP_01.xlsx")
ws = wb.active

# DESCRIPTIONS OF CLIENT IDs HAVE TO MATCH
ws["A2"].value = "Franchise ID"



# GET ALL CLIENTS IN THESE DICTS

for row in range(2, 79):
    for column in range(1, 36):
        char = get_column_letter(column)
        
        if char == 'A':
            datadump1_organized[ws[char + str(row)].value] = []
            
        else:
            datadump1_organized[ws['A' + str(row)].value] += [ws[char + str(row)].value]       



wb = load_workbook("DATADUMP_02.xlsx")
ws = wb.active


for row in range(3, 79):
    for column in range(1, 8):
        char = get_column_letter(column)
        
        if char == 'A':
            datadump2_organized[ws[char + str(row)].value] = []
            
        else:
            datadump2_organized[ws['A' + str(row)].value] += [ws[char + str(row)].value]
    

    
# SORT THE DICTIONARIES:
    # IDEAL: KEEP THE ORDER OF THE SHEETS.
    # FOR EACH ITEM OF DIC1, LOOP THROUGH ALL OF DIC2. 
    # IF YOU FIND THE MATCH, MERGE IT
    # ELSE, ADD IT ELSEWHERE SO WE KNOW

joint_data = {}
datadump1_leftovers = datadump1_organized.copy()
datadump2_leftovers = datadump2_organized.copy()

for client1 in datadump1_organized:
    for client2 in datadump2_organized:
        if client1 in client2:
            
            client = client1
            
            joint_data[client] = datadump1_organized[client1]
            joint_data[client] += datadump2_organized[client2]
            
            del datadump1_leftovers[client1]
            del datadump2_leftovers[client2]


    
# Produce the JOINT EXCEL SHEET


wb = Workbook()
ws = wb.active
ws.title = "Merged Data"


   

keys = list(joint_data)


for ID, row in enumerate(joint_data):
    counter = 1
    
    # before it loops through the list, add the key to the first column.
    ws["A"+str(ID+1)].value = keys[ID]
    
    for column in joint_data[row]:
        counter += 1
        char = get_column_letter(counter)
        ws[char + str(ID+1)].value = column
    
    


wb.save('TEST.xlsx')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    