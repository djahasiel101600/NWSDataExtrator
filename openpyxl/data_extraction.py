import sqlite3
from openpyxl import load_workbook
import os
from datetime import datetime
from glob import glob


# Register adapter
def adapt_datetime(dt):
    return dt.isoformat()

# Register converter
def convert_datetime(s):
    return datetime.fromisoformat(s.decode('utf-8'))

sqlite3.register_adapter(datetime, adapt_datetime)
sqlite3.register_converter("DATETIME", convert_datetime)

# Sqlite Codes
db = sqlite3.connect("db.sqlite3")
cur = db.cursor()

try:
    with open("code.sql", "r") as sqlscipt:
        # cur.executescript(sqlscipt.read())
        pass
except sqlite3.OperationalError:
    print("Schema is already created")
    input("Press Enter to Continue")




file_paths = glob(".\\excel_dir\\*.xlsx")


def clean_whitespace(char):
    char = char.strip()
    char = " ".join(char.split())
    return char

def target_finder(field):
    target_col = 0
    target_row = 1
    for cell in ws.iter_rows():
        for c in cell:
            
            try:
                match_field = clean_whitespace(c.value)
                if match_field == field:
                    target_col = c.column
                    target_row = c.row
                else:
                    # print("Not Found")
                    pass
                    
            except AttributeError:
                pass
                
           
    n = 1
    if target_col == 6:
        return ws.cell(row=target_row, column=target_col + 4).value
    else:
        return ws.cell(row=target_row, column=target_col + 1).value
    
    # while True:
    #     if not ws.cell(row=target_row, column=target_col + n).value:
    #         n += 1
    #     else:
    #         return ws.cell(row=target_row, column=target_col + n).value

def target_finder_row_col(field):
    for cells in ws.iter_rows():
        for cell in cells:
            if cell.value == field:
                target_row = cell.row
                target_col = cell.column

                result = {"col":target_col, "row":target_row}
                
                return result    

data = []

n = 1
skip = False
for file_path in file_paths:
    wb = load_workbook(file_path,data_only=True)
    print(file_path)
    user = input("Proceed or Skip? y/n: ")
    while True:
        if user.lower() == "y":
            skip = True
            break
        elif user.lower() == "n":
            skip = False
            break
        else:
            print("Invalid Choice")
    if skip:
        skip = False
        continue
    else:
        pass
     
    location = input("location: ").lower()
    prefix = input("PREFIX: ")
    
    SQL_SCRIPT = \
f"""--DELETE FROM {location}_requests_{location}requestitems where description is null and quantity is null and  unit is null and unit_cost is null and amount is null;
UPDATE {location}_requests_{location}requestitems SET unit_cost=0 WHERE unit_cost is null;
UPDATE {location}_requests_{location}requestitems SET amount=0 WHERE amount is null;
UPDATE {location}_requests_{location}requestitems SET quantity=0 WHERE quantity is null;
update {location}_requests_monitoring set header_id=id where header_id="";"""

    for ws in list(reversed(wb.worksheets)):
        print("*" * 60)
        PAYEE = target_finder("Payee:")
        PARTICULARS = target_finder("Particulars:")
        PROJECT = target_finder("Project:")
        RS_NO = str(target_finder("R.S. #"))
        RS_NO = f"{prefix}-" + RS_NO.zfill(3)
        DATE_REQUESTED = target_finder("Date Requested:")
        NOTE = target_finder("NOTE:")

        print("Payee:",PAYEE)
        print("Particulars:",PARTICULARS)
        print("Project:",PROJECT)
        print("RS No.:", RS_NO)
        print("Date Requested:", DATE_REQUESTED.date())

        HEADERS = ["QTY", "UNIT", "DESCRIPTION", "UNIT COST", "AMOUNT"]

        start_location = target_finder_row_col("QTY")
        v_end_location = target_finder_row_col("NOTE:")
        h_end_location = target_finder_row_col("AMOUNT")
        
        if not target_finder_row_col("QTY"):
            input("[QTY] Column not Found, press Enter to Continue...")
            with open("log.txt", "a") as fw:
                fw.write(f"{RS_NO}\n")
        else:
            start_location = target_finder_row_col("QTY")
            
        if not target_finder_row_col("NOTE:"):
            input("[NOTE:] Column not Found, press Enter to Continue...")
            with open("log.txt", "a") as fw:
                fw.write(f"{RS_NO}\n")
        else:
            v_end_location = target_finder_row_col("NOTE:")
        
        if not target_finder_row_col("AMOUNT"):
            input(" Column not Found, press Enter to Continue...")
            with open("log.txt", "a") as fw:
                fw.write(f"{RS_NO}\n")
        else:
            h_start_location = target_finder_row_col("AMOUNT")
        
        def col_items(start_location, v_end_location, h_end_location):        
            start_location = target_finder_row_col("QTY")
            v_end_location = target_finder_row_col("NOTE:")
            h_end_location = target_finder_row_col("AMOUNT")
            
            start_row = start_location['row']
            start_col = start_location['col']

            V_end_row = v_end_location['row']
            V_end_col = v_end_location['col']

            H_end_row = h_end_location['row']
            H_end_col = h_end_location['col']
                
                    
            min_col = start_col
            min_row = start_row + 1
            max_col = H_end_col
            max_row = V_end_row - 1

            ITEMS = []
            for cells in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                
                array = [cell.value for cell in cells if cell.column not in [7,8,9] + list(range(11,99))]
                array.pop(3)
                nonetype = 0
                for i in array:
                    if i == None:
                        nonetype += 1
                    else:
                        pass
                if len(array) == nonetype:
                    print("NoneType Array: ", array)
                else:
                    ITEMS.append(array)
            
            return ITEMS
        
        def sanitizer(var):
            if not var:
                return " "
            else:
                return var
            
        rs_number = sanitizer(RS_NO)
        particulars = sanitizer(PARTICULARS)
        payee = PAYEE
        project = sanitizer(PROJECT)
        urgent = ""
        note = NOTE
        date_requested = sanitizer(DATE_REQUESTED.date())
        date_needed = ""
        user_id = 1
        last_modified = datetime.today().date()
        created = datetime.today().date()
        
        try:
            cur.execute(f'''INSERT INTO {location}_requests_{location}requestheader (rs_number, particulars, payee, project, urgent, note, date_requested, date_needed, last_modified, created, user_id) VALUES (?,?,?,?,?,?,?,?,?,?,?)''', (rs_number, particulars, payee, project, urgent, note, date_requested, date_needed, last_modified, created, user_id))
            print(f"Running: {n}")
            n += 1
            
            ITEM_ARRAY = col_items(start_location, v_end_location, h_end_location)

            for i in ITEM_ARRAY:
                print(i)
                cur.execute(f'''INSERT INTO {location}_requests_{location}requestitems (quantity, unit, description, unit_cost, amount, header_id, served, ignore, item_id) VALUES (?,?,?,?,?,?,?,?,?)''', tuple(i) + (rs_number, False, False,"-"))
                cur.execute(f'''INSERT INTO {location}_requests_monitoring (PO_no, PO_date, delivery_date, receiving_report, DR_no, SI_no, OR_no, CR_no, withdrawal_no, item_date, header_id) VALUES (?,?,?,?,?,?,?,?,?,?,?)''', ("","","","","","","","","","","",))
                os.system('cls')
                print(col_items(start_location, v_end_location, h_end_location).index(i))
            
        except sqlite3.IntegrityError:
            with open("log.txt", "a") as fw:
                content = "************************\n"
                content += rs_number + "\n"
                fw.write(content)
        
        
        
        cur.executescript(SQL_SCRIPT)
        
        
        
    while True:
        user = input("Do you want to save changes to database? y/n: ")
        if user == "y":
            db.commit()
            break
        elif user == "n":
            break
        else:
            print("Invalid Choice")

print("Done")
    
    
