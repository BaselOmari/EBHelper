import openpyxl

wb = openpyxl.load_workbook('Publications Assortment (Our Project) Finished.xlsx')

dupesSheet = wb['Dupes Removed']

def remove_row_dupes_removed(row_number):
    global dupesSheet

    cellInt = row_number
    while dupesSheet[f"B{cellInt}"].value != None:
        col = 'B'
        while col != 'E':
            dupesSheet[f"{col}{cellInt}"].value = dupesSheet[f"{col}{cellInt+1}"].value
            col = chr(ord(col) + 1)
        
        cellInt += 1
    print("################# REMOVED", row_number, "####################")


def duplicate_finder():
    global dupesSheet
    cellInt = 3

    while dupesSheet[f"B{cellInt}"].value != None:
        original_value = dupesSheet[f"B{cellInt}"].value

        travInt = 3
        while dupesSheet[f"B{travInt}"].value != None:
            if travInt is cellInt:
                travInt += 1
                continue
            
            if dupesSheet[f"B{travInt}"].value+"." == original_value:
                remove_row_dupes_removed(travInt)
                if (travInt < cellInt):
                    cellInt = travInt
                
                cellInt -= 1
                break

            travInt += 1

        print(cellInt)
        cellInt += 1
    

duplicate_finder()
wb.save(filename = 'Publications Assortment (Our Project) Finished.xlsx')