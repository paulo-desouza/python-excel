from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import datetime 

def days_to_months(dt_days):
    
    if dt_days == str(dt_days):
        return dt_days
    
    
    days = dt_days.days
    months = 0
    while days > 30:
        months += 1
        days -= 30
    return str(months)



def time_check(td, time=75):
    """
    Parameters
    ----------
    td : deltatime
        Deltatime value.
    time : Integer
        Day limit that determines if we are on or off track. The default is 75.

    Returns
    -------
    """
    on_track = True
    int_days = 0
    int_months = 0
    
    if td == str(td):
        return td
    else:
        
        int_days = int(td.days)
        if int_days > time:
            on_track = False
        
        
        while int_days > 30:
            int_months += 1
            int_days -= 30
        
        if time == 0:
            if int_months == 0:
                return str(int_days) + " day(s)"
            else:
                return str(int_months) + " month(s) and " + str(int_days) + " day(s)"
        if on_track:
            if int_months == 0:
                return "On Track (" + str(int_days) + " day(s))" 
            else:
                return "On Track (" + str(int_months) + " month(s) and " + str(int_days) + " day(s))" 
        else:
            if int_months == 0:
                return "Off Track (" + str(int_days) + " day(s))" 
            else:
                return "Off Track (" + str(int_months) + " month(s) and " + str(int_days) + " day(s))" 
            
        

def date_subtraction(start, finish, format="days"):
    """
    Parameters
    ----------
    start : MM/DD/YYYY
    finish : MM/DD/YYYY
    
    Desc: Converts both dates into DateTime Objects and subtracts them from one another.
    
    Returns 
    -------
    Time period in between the 2 dates.

    """
    
    if "/" in start and "/" in finish:
    
        strip1 = start.split("/")
        strip2 = finish.split("/")
        
    else:
        
        return "Missing Information"
    
    
    datetime1 = datetime.datetime(int(strip1[2]), int(strip1[0]), int(strip1[1]))  # y, m, d
    datetime2 = datetime.datetime(int(strip2[2]), int(strip2[0]), int(strip2[1]))
    
    return datetime2 - datetime1
    



def time_from_present(start, present=datetime.datetime.now(), format="days"):
    """
    Parameters
    ----------
    start : MM/DD/YYYY
    
    Desc: Converts the date into DateTime Object and subtracts it from the present.
    
    Returns 
    -------
    Time period in between the date and the present.

    """
    
    if "/" in start:
    
        strip1 = start.split("/")

        
    else:
        
        return "Missing Information"
    
    
    datetime1 = datetime.datetime(int(strip1[2]), int(strip1[0]), int(strip1[1]))  # y, m, d
    
    return datetime1 - present


# CREATE A DICTIONARY WITH KEY AS CUSTOMER ID, AND VALUE AS LIST WITH CUSTOMER 
# INFORMATION.  FIRST ITEM HAS KEY "ID" AND VALUE AS LIST WITH ALL COLUMN IDs




datadump1_organized = {}
datadump2_organized = {}


wb = load_workbook("DATADUMP_01.xlsx")
ws = wb.active

# DESCRIPTIONS OF CLIENT IDs HAVE TO MATCH
ws["A2"].value = "Franchise ID"


# COUNT THE ROWS

count = 0
for row in ws:
    if not all([cell.value is None for cell in row]):
        count += 1


# GET ALL CLIENTS IN THESE DICTS

for row in range(2, count + 1):
    for column in range(1, 36):
        char = get_column_letter(column)
        
        if char == 'A':
            datadump1_organized[ws[char + str(row)].value] = []
            
        else:
            datadump1_organized[ws['A' + str(row)].value] += [ws[char + str(row)].value]       



wb = load_workbook("DATADUMP_02.xlsx")
ws = wb.active

count = 0
for row in ws:
    if not all([cell.value is None for cell in row]):
        count += 1


for row in range(3, count + 1):
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
    
    
filler_list = ["--", "--", "--", "--", "--", "--"]
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
        
        
for client in datadump1_leftovers:
    
    joint_data[client] = datadump1_leftovers[client]
    joint_data[client] += filler_list
    
            

keys = list(joint_data)    

# FROM JOINT DATA,MAKE ONE SEPARATE LIST WITH THE DATA NEEDED FOR  
# EACH OF THE TABLES.



contruction_pipeline = []

# loop through the dictionary, and during each iteration, pull and append the data you need to the list.

for c, client in enumerate(joint_data):
    if c == 0:
        pass
    
    else:
        
        if joint_data[client][13].strip() == "--" :
            finance = "In Process / Self Funded"
        else:
            finance = "Funded"
            
        contruction_pipeline.append(
            {   "Franchise ID" : client,
                "Project Status" : joint_data[client][35],
                "City, State" : joint_data[client][37] + ", " + joint_data[client][38],
                "Architecturals" : time_check(date_subtraction(joint_data[client][9], joint_data[client][10])),
                "Permitting" : time_check(date_subtraction(joint_data[client][10], joint_data[client][11])),
                "Active Construction" : time_check(date_subtraction(joint_data[client][15], joint_data[client][25])),
                "Final Fitout" : time_check(date_subtraction(joint_data[client][25], joint_data[client][31])),
                "Financing Completed" : finance,
                "Project Start Date" : joint_data[client][32],
                "Projected Opening" : joint_data[client][39],
                "Total Months in Process" : days_to_months(time_from_present(joint_data[client][39])),
                "Projected Total Project" : "???",
                "Notes" : ""
                
             })
        

realestate_pipeline = []

for c, client in enumerate(joint_data):
    if c == 0:
        pass
    
    else:
        realestate_pipeline.append({  
                "Franchise ID" : client,
                "Project Status" : joint_data[client][35],
                "City" : joint_data[client][37],
                "State/Province" : joint_data[client][37],
                "Site Selection" : time_check(date_subtraction(joint_data[client][2], joint_data[client][3]), time = 45),
                "LOI Negotiation" : time_check(date_subtraction(joint_data[client][4], joint_data[client][6]), time = 150),
                "Lease Negotiation" : time_check(date_subtraction(joint_data[client][6],joint_data[client][9]), time = 90),
                "Expected Opening Date" : joint_data[client][39],
                "Months in Process" : days_to_months(time_from_present(joint_data[client][32])),
                "Projected Total Months" : "???"
             })


open_schools = []

for c, client in enumerate(joint_data):
    if c == 0:
        pass
    
    else:
        if joint_data[client][28] != "--":
            open_schools.append({  
                     "Franchise ID" : client,
                     "City, State" : joint_data[client][37] + ", " + joint_data[client][38],
                     "Open Date" : joint_data[client][28],
                     "Months to open" : time_check(date_subtraction(joint_data[client][1],joint_data[client][28]), time = 0),
                     "Opening Order" : "index of list sorted by the open date"
                 })


# REAL ESTATE PIPELINE TABLE ITERATION


wb = Workbook()
ws = wb.active
ws.title = "Real Estate Pipeline"

titles = list(realestate_pipeline[0])
values = []

ws.append(titles)
row = 0
for client in realestate_pipeline:
    column = 0
    row += 1
    for key in client:
        column += 1
        char = get_column_letter(column)
        ws[char + str(row)].value = client[key]
        

# CONTRUCTION PIPELINE TABLE ITERATION

wb.create_sheet("Construction Pipeline")
ws = wb["Construction Pipeline"]

titles = list(contruction_pipeline[0])
values = []

ws.append(titles)
row = 0
for client in contruction_pipeline:
    column = 0
    row += 1
    for key in client:
        column += 1
        char = get_column_letter(column)
        ws[char + str(row)].value = client[key]
        
    
wb.create_sheet("Open Schools")
ws = wb["Open Schools"]

titles = list(open_schools[0])
values = []

ws.append(titles)
row = 0
for client in open_schools:
    column = 0
    row += 1
    for key in client:
        column += 1
        char = get_column_letter(column)
        ws[char + str(row)].value = client[key]

    
wb.save('TEST11.xlsx') 

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    