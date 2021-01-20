import openpyxl

wb = openpyxl.load_workbook('Publications Assortment (Our Project) Finished.xlsx')

ebscoSheet = wb.get_sheet_by_name('MedLine')
pubmedSheet = wb.get_sheet_by_name('PubMed')
combinedSheet = wb.get_sheet_by_name('End Combined')

ebsco_titles = []
pubmed_titles = []
combined_titles = []

def filling_titles():
    global ebsco_titles,pubmed_titles, combined_titles

    cell_int = 2
    while True:
        done = True

        if ebscoSheet[f"B{cell_int}"].value != None:
            done = False
            ebsco_titles.append(ebscoSheet[f"B{cell_int}"].value)
        
        if pubmedSheet[f"B{cell_int}"].value != None:
            done = False
            pubmed_titles.append(pubmedSheet[f"B{cell_int}"].value)
        
        if combinedSheet[f"B{cell_int+1}"].value != None:
            done = False
            combined_titles.append(combinedSheet[f"B{cell_int+1}"].value)
        
        if done:
            break
        
        cell_int += 1


def reset_row(row):
    col = 'C'
    while col != 'P':
        combinedSheet[f"{col}{row}"].value = None
        col = chr(ord(col) + 1)

        

def fillEBSCO(eb_cellInt, comb_cellInt):
    reset_row(comb_cellInt)
    
    # Author
    combinedSheet[f"C{comb_cellInt}"].value = ebscoSheet[f"C{eb_cellInt}"].value

    # Journal Title
    combinedSheet[f"D{comb_cellInt}"].value = ebscoSheet[f"D{eb_cellInt}"].value

    # Publication Date
    combinedSheet[f"E{comb_cellInt}"].value = ebscoSheet[f"F{eb_cellInt}"].value

    # Publisher
    combinedSheet[f"F{comb_cellInt}"].value = ebscoSheet[f"G{eb_cellInt}"].value

    # Publication Type
    combinedSheet[f"G{comb_cellInt}"].value = ebscoSheet[f"H{eb_cellInt}"].value

    # Subjects
    combinedSheet[f"H{comb_cellInt}"].value = ebscoSheet[f"I{eb_cellInt}"].value

    # Keywords
    combinedSheet[f"I{comb_cellInt}"].value = ebscoSheet[f"J{eb_cellInt}"].value

    # Abstract
    combinedSheet[f"J{comb_cellInt}"].value = ebscoSheet[f"K{eb_cellInt}"].value

    # P-Link
    combinedSheet[f"K{comb_cellInt}"].value = ebscoSheet[f"L{eb_cellInt}"].value
    

def fillPub(pub_cellInt, comb_cellInt):
    reset_row(comb_cellInt)

    # Author
    combinedSheet[f"C{comb_cellInt}"].value = pubmedSheet[f"C{pub_cellInt}"].value

    # Journal Title
    combinedSheet[f"D{comb_cellInt}"].value = pubmedSheet[f"D{pub_cellInt}"].value

    # Publication Date
    combinedSheet[f"E{comb_cellInt}"].value = pubmedSheet[f"E{pub_cellInt}"].value

    # Publisher
    combinedSheet[f"F{comb_cellInt}"].value = None

    # Publication Type
    combinedSheet[f"G{comb_cellInt}"].value = pubmedSheet[f"F{pub_cellInt}"].value

    # Subjects
    combinedSheet[f"H{comb_cellInt}"].value = None

    # Keywords
    combinedSheet[f"I{comb_cellInt}"].value = None

    # Abstract
    combinedSheet[f"J{comb_cellInt}"].value = pubmedSheet[f"G{pub_cellInt}"].value

    # P-Link
    combinedSheet[f"K{comb_cellInt}"].value = None


filling_titles()

for index, title in enumerate(combined_titles):
    combined_cell = index + 3
    try:
        eb_cell = ebsco_titles.index(title) + 2
        fillEBSCO(eb_cell,combined_cell)

    except ValueError:
        try:
            pub_cell = pubmed_titles.index(title) + 2
            fillPub(pub_cell,combined_cell)

        except ValueError:
            print("ERROR IN", index)
    


    print(index, "DONE")


wb.save(filename='Publications Assortment (Our Project) Finished.xlsx')