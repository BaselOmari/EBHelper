import pandas as pd
import openpyxl

wb = openpyxl.load_workbook('EBSCOPT2.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

data = pd.read_csv("Downloads//deliveryMerged.csv")


data = data[['Article Title','Author','Journal Title','ISSN','Publication Date','Publisher','Doctype','Subjects','Keywords','Abstract','PLink']]


for i in range(len(data.keys())):
    char_val = chr(i + 66)
    sheet[f'{char_val}1'] = data.keys()[i]

for i in range(len(data)):
    articleNumber = i + 1
    cellInt = i + 2

    sheet[f'A{cellInt}'] = articleNumber

    title = data['Article Title'][i]
    try:
        if title[0] == '[':
            title = title[1:-2]
    except:
        print("Something Wrong Happened")
    titleCell = 'B' + str(cellInt)
    sheet[titleCell] = title

    author = data['Author'][i]
    authorCell = 'C' + str(cellInt)
    sheet[authorCell] = author

    journalTitle = data['Journal Title'][i]
    jtCell = 'D' + str(cellInt)
    sheet[jtCell] = journalTitle

    ISSN = data['ISSN'][i]
    ISSNCell = 'E' + str(cellInt)
    sheet[ISSNCell] = ISSN

    pubDate = data['Publication Date'][i]
    pubDateCell = 'F' + str(cellInt)
    sheet[pubDateCell] = pubDate

    publisher = data['Publisher'][i]
    pubCell = 'G' + str(cellInt)
    sheet[pubCell] = publisher

    doctype = data['Doctype'][i]
    doctypeCell = 'H' + str(cellInt)
    sheet[doctypeCell] = doctype

    subjects = data['Subjects'][i]
    subjectsCell = 'I' + str(cellInt)
    sheet[subjectsCell] = subjects

    keywords = data['Keywords'][i]
    keywordsCell = 'J' + str(cellInt)
    sheet[keywordsCell] = keywords

    abstract = data['Abstract'][i]
    abstractCell = 'K' + str(cellInt)
    sheet[abstractCell] = abstract

    pLink = data['PLink'][i]
    pLinkCell = 'L' + str(cellInt)
    sheet[pLinkCell] = pLink

    wb.save(filename='EBSCOPT2.xlsx')
    print(f'Article {articleNumber} Done')
