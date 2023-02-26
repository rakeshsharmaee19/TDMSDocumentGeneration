import os

from .styleTable import *


### max 34 cell for .5 cm or total length is 17cm
def fillOverAllTest(table, data):
    table.allow_autofit = False
    count = 0
    for i in range(len(table.rows)):
        if i == 0:
            rowHeight(table.rows[i], 0.5)
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11, text=data[0][count])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            count += 1
        elif i == 1:
            rowHeight(table.rows[i], 0.5 * data[2])
            for cell in range(len(table.rows[i].cells) - 1):
                if cell == 0:
                    table.rows[i].cells[cell].width = Cm(1.)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                elif cell == 1:
                    table.rows[i].cells[cell].width = Inches(1.8)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                elif cell == 2 or cell == 3:
                    table.rows[i].cells[cell].width = Inches(1.1)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                else:
                    table.rows[i].cells[cell].width = Inches(1.1)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
        elif i == 2:
            rowHeight(table.rows[i], 0.5 * data[2])
            table.rows[i].cells[cell].width = Inches(1.1)
            makeBoldItalic(table, rowNumber=i, cellNumber=4, bold=True, italic=False, fontSize=10, text=data[0][count])
            table.rows[i].cells[4].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            count += 1
            table.rows[0].cells[cell].width = Inches(1.1)
            makeBoldItalic(table, rowNumber=i, cellNumber=5, bold=True, italic=False, fontSize=10, text=data[0][count])
            table.rows[i].cells[5].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            count = 0
        else:
            rowHeight(table.rows[i], 0.5 * data[2])
            for cell in table.rows[i].cells:
                if data[1][count] == None:
                    cell.text = ""
                    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                else:
                    cell.text = str(data[1][count])
                    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillBushingTest(table, data, merge_list, row=0.5):
    count = 0
    mergeCount = 0
    for i in range(len(table.rows)):
        if i == 0:
            rowHeight(table.rows[i], 0.5)
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11, text=data[0][count])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            count += 1
        elif i == 1:
            rowHeight(table.rows[i], row)
            cell = 0
            while cell < len(table.rows[i].cells):
                if cell < 4:
                    table.rows[i].cells[cell].width = Cm(1.5)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    if cell == 2:
                        table.rows[i].cells[cell].width = Cm(2.3)
                    cell += 1
                else:
                    table.rows[i].cells[cell].width = Inches(1.1)
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell += merge_list[mergeCount]
                    mergeCount += 1
                count += 1
        elif i == 2:
            rowHeight(table.rows[i], row)
            # table.rows[i].cells[cell].width = Inches(1.1)
            for cel in range(4, len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cel, bold=True, italic=False, fontSize=10,
                               text=data[0][count])
                table.rows[i].cells[cel].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            rowHeight(table.rows[i], row)
            for cell in table.rows[i].cells:
                if data[1][count] == None:
                    cell.text = ""
                    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                else:
                    cell.text = str(data[1][count])
                    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillExcitationTest(table, data, row=0.5):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillExcitationData(table, data, mergeList):
    count = 0
    mergeCount = 0
    for i in range(len(table.rows)):
        rowHeight(table.rows[i], 0.5)
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][count])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            cell = 0
            while cell < len(table.rows[i].cells):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell += mergeList[mergeCount]
                mergeCount += 1
                count += 1
        elif i == 2:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillRatioTest(table, data, rowValue, rowH=0.5, key=True):
    count = 0
    if key == True:
        for i in range(len(table.rows)):
            rowHeight(table.rows[i], rowH)
            if i == 0 or i == 2:
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            elif i == 1:
                rowHeight(table.rows[i], 1)
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=False, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.LEFT
                count += 1
            elif i == 3:
                cell = 0
                while cell < len(table.rows[i].cells):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell += 3
                    count += 1
            elif i == 4:
                cell = 0
                while cell < len(table.rows[i].cells):
                    if cell in (0, 2, 3, 6):
                        makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                       text=data[0][count])
                        table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                        count += 1
                    cell += 1
            elif i == 5:
                cell = 0
                while cell < len(table.rows[i].cells):
                    if cell != 2:
                        makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                       text=data[0][count])
                        table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                        if cell == 2:
                            table.cell(i, 2).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
                        count += 1
                    cell += 1
                count = rowValue * 9
            else:
                for cell in table.rows[i].cells:
                    if data[1][count] == None:
                        cell.text = ""
                        cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    else:
                        cell.text = str(data[1][count])
                        cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1

    else:
        count = 2
        for i in range(len(table.rows)):
            if i == 0:
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            elif i == 1:
                cell = 0
                while cell < len(table.rows[i].cells):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    if cell == 2:
                        table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
                    cell += 3
                    count += 1
            elif i == 2:
                cell = 0
                while cell < len(table.rows[i].cells):
                    if cell in (0, 2, 3, 6):
                        makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                       text=data[0][count])
                        table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                        count += 1
                    cell += 1
            elif i == 3:
                cell = 0
                while cell < len(table.rows[i].cells):
                    if cell != 2:
                        makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                       text=data[0][count])
                        table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                        count += 1
                    cell += 1
                count = rowValue * 9
            else:
                for cell in table.rows[i].cells:
                    if data[1][count] == None:
                        cell.text = ""
                        cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    else:
                        cell.text = str(data[1][count])
                        cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1


def fillDgaTable(table, data, dataFillCount=0):
    count = 0
    for i in range(len(table.rows)):
        if i == 0:
            for cell in range(len(table.rows[i].cells)):
                set_vertical_cell_direction(cell=table.rows[i].cells[cell], direction="btLr")
                textT = " " + data[0][count]
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=10,
                               text=textT)
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.LEFT
                table.cell(i, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
            # count = 0
        else:
            for cell in table.rows[i].cells:
                if data[1][dataFillCount] == None:
                    cell.text = ""
                    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                else:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    makeBoldPara(cell.paragraphs[0], bold=False, italic=False, fontSize=10,
                                 text=str(data[1][dataFillCount]),
                                 alignment=1)
                dataFillCount += 1

    return dataFillCount


def check_filepath(data):
    foldername = ''
    for i in data["DATA"]:
        foldername = data["DATA"][i]['transformerId']
        break
    # path = data['FileDescription']['filepath']
    path = "C://Users//Rakesh//Desktop"
    path1 = path + "//" + foldername + "//"
    isExist = os.path.exists(path1)
    if not isExist:
        os.makedirs(path1 + "//")
    return path1


def innerTableDataFill(table, data, dataCondition=False):
    table.style = 'TableGrid'
    table.autofit = False
    table.allow_autofit = False
    count = 0
    numberOfColumns = len(table.rows[0].cells)
    table.cell(0, 1).merge(table.cell(0, numberOfColumns - 1))
    if dataCondition:
        for row in range(len(table.rows)):
            rowHeight(table.rows[row], 0.5)
            for cell in range(len(table.rows[row].cells)):
                table.rows[row].cells[cell].width = Cm(2.3)
                if row == 0:
                    if data[0][count] == None:
                        table.rows[row].cells[cell].text = ""
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    else:
                        table.rows[row].cells[cell].text = str(data[0][count])
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
                elif row == 1:
                    if data[0][count] == None:
                        table.rows[row].cells[cell].text = ""
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    else:
                        table.rows[row].cells[cell].text = str(data[0][count])
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
                    if cell == (len(table.rows[row].cells) - 1):
                        count = 0
                else:
                    if data[1][count] == None:
                        table.rows[row].cells[cell].text = ""
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    else:
                        table.rows[row].cells[cell].text = str(data[1][count])
                        table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
    else:
        for row in range(len(table.rows)):
            rowHeight(table.rows[row], 0.5)
            for cell in range(len(table.rows[row].cells)):
                table.rows[row].cells[cell].width = Cm(2.3)
                if data[count] == None:
                    table.rows[row].cells[cell].text = ""
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                else:
                    table.rows[row].cells[cell].text = str(data[count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillTableTopData(table, data):
    count = 0
    for row in table.rows:
        rowHeight(row, 0.5)
        for cell in row.cells:
            if data[count] == None:
                cell.text = ""
                cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            else:
                cell.text = str(data[count])
                cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            count += 1


def addMiddleTable(document, row=2):
    midlleTable = document.add_table(rows=row, cols=1)
    midlleTable.Border = None
    midlleTable.autofit = False
    styleBelowTable(midlleTable)


def fillFirstTestData(table, data):
    count = 0
    for row in range(len(table.rows)):
        rowHeight(table.rows[row], 0.5)
        if row == 0:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=10,
                               text=data[0][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=10,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1


def filltextTable(table, data):
    count = 0
    for row in range(len(table.rows)):
        if row == 0:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
        else:
            for cell in range(len(table.rows[row].cells)):
                textValue = '      '+data[count]
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=textValue)
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.LEFT
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1

def fillEfficienceTestTable(table, data, mergerValue=[2,2]):
    count = 0
    mergeCount = 0
    for row in range(len(table.rows)):
        if row == 0:
            cell = 0
            while cell < len(table.rows[row].cells):
                if cell>1:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell+=mergerValue[mergeCount]
                    mergeCount+=1
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell+=1
                count += 1
        elif row == 1:
            for cell in range(2, len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
            count=0
        elif row == len(table.rows)-1:
            for cell in range(1, len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1

def fillControlTestTable(table, data, mergeValue=[]):
    count = 0
    mergeCount =mergeValue
    for row in range(len(table.rows)):
        if row == 0:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
        elif row == 1:
            for cell in range(len(table.rows[row].cells)):
                if cell ==2:
                    pass
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
        elif row ==2:
            for cell in range(1,len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1

def fillEfficenceTestTable(table, data):
    count = 0
    for row in range(len(table.rows)):
        if row == 0:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
        elif row ==1:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count].upper())
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1
            count=0
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1

def fillPdHvParallelTestTable(table, data,mergeRow = []):
    count = 0
    mergeCount = 0
    for row in range(len(table.rows)):
        rowHeight(table.rows[row], 0.5)
        if row == 0:
            rowHeight(table.rows[row], 1)
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=False, italic=False, fontSize=11,
                           text=data[0][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
        elif row == 1:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
        elif row == 2:
            cell= 0
            while cell < len(table.rows[row].cells):
                if cell<4:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell+=1
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell+=mergeRow[mergeCount]
                    mergeCount+=1
                count+=1
        elif row == 3:
            for cell in range(4,len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count+=1
            count = 0
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count+=1

def fillImpulseTestTable(table, data):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1

def fillPotencialTestTable(table, data):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1

def fillParallelInducedVoltageTestData(table, data):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillSeriesInducedVoltageTestData(table, data):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1

def fillinsulationResistanceFillData(table, data, numOfHeaders=2):
    if numOfHeaders == 2:
        count = 1
        for i in range(len(table.rows)):
            if i == 0:
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                               text=data[0][0])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            elif i == 1:
                for cell in range(len(table.rows[i].cells)):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
                count = 0
            else:
                for cell in range(len(table.rows[i].cells)):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data[1][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
    else:
        count = 2
        for i in range(len(table.rows)):
            if i == 0:
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                               text=data[0][0])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            elif i == 1:
                makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                               text=data[0][1])
                table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            elif i == 2:
                for cell in range(len(table.rows[i].cells)):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1
                count = 0
            else:
                for cell in range(len(table.rows[i].cells)):
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data[1][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    count += 1

def fillThermalTestData(table, data, numberOfHeaderRows = 2):
    count= 0
    mergeCount = 0
    row = 0
    rowMergeCount = 0
    while row<len(table.rows):
        if row <numberOfHeaderRows-1:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data['tableHeader'][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
            row+=1
        elif row == numberOfHeaderRows-1:
            cell =0
            while cell<len(table.rows[row].cells):
                if cell in data['mergeHeaderCell']:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell+=1
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell+=data['mergeHeader'][mergeCount]
                    mergeCount+=1
            cell = 0
            while cell<len(table.rows[row].cells):
                if cell not in data['mergeHeaderCell']:
                    makeBoldItalic(table, rowNumber=row+1, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row+1].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row+1, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                cell+=1
            row += 2
            count = 0
        elif row > numberOfHeaderRows and row < len(table.rows) - 3:
            cell = 0
            while cell < len(table.rows[row].cells):
                if cell in data['tableDataMerge'][rowMergeCount][str(rowMergeCount + 1)]:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableData'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableData'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    makeBoldItalic(table, rowNumber=row+1, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableData'][count])
                    table.rows[row+1].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row+1, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1

                cell += 1
            rowMergeCount += 1
            row+=2
        else:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=False, italic=False, fontSize=11,
                          text=data['tableData'][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
            row+=1

def fillDewPointTestTable(table, data):
    count = 0
    for row in range(len(table.rows)):
        if row == 0:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        else:
            for cell in range(len(table.rows[row].cells)):
                makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                count += 1


def fillCoreExcitationStateTestData(table, data, data_merge, cellToInsertCoreExcitationHeader):
    count = 1

    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

        elif i == 1:
            for cell in range(len(table.rows[i].cells)):

                if cell in data_merge:
                    makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                                   text=data[0][count])
                    table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

                    count += 1

        elif i == 2:
            for cell in cellToInsertCoreExcitationHeader:
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])

                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

                count += 1
            count = 0

        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillLeadResistanceTestTableData(table, data):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif i == 1:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1
            count = 0
        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillZeroSequence(table, data, cell_data_row2, cell_data_row3):
    count = 1
    for i in range(len(table.rows)):
        if i == 0:
            makeBoldItalic(table, rowNumber=i, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data[0][0])
            table.rows[i].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

        elif i == 1:
            for cell in cell_data_row2:
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

                count += 1

        elif i == 2:
            for cell in cell_data_row3:
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=True, italic=False, fontSize=11,
                               text=data[0][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

                count += 1
            count = 0

        else:
            for cell in range(len(table.rows[i].cells)):
                makeBoldItalic(table, rowNumber=i, cellNumber=cell, bold=False, italic=False, fontSize=11,
                               text=data[1][count])
                table.rows[i].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                count += 1


def fillTable(table, data, numberOfHeaderRows = 6, numberOfCells = 13):
    count = 0
    mergeCount = 0
    row = 0
    rowMergeCount = 1
    fieldCount = 0
    while row < len(table.rows):
        if row==0:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text="SOUND TEST")
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row+=1
        elif row < numberOfHeaderRows - 1:
            makeBoldItalic(table, rowNumber=row, cellNumber=0, bold=True, italic=False, fontSize=11,
                           text=data['tableHeader'][count])
            table.rows[row].cells[0].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.cell(row, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            count += 1
            row += 1
        elif row == numberOfHeaderRows - 1:
            cell = 0
            while cell < len(table.rows[row].cells):
                if cell in data['mergeHeaderCell']:
                    # print(cell, 'first')
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell += 1
                else:
                    makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell += data['mergeHeader'][mergeCount]
                    mergeCount += 1
            cell = 0
            while cell < len(table.rows[row].cells):
                if cell not in data['mergeHeaderCell']:
                    makeBoldItalic(table, rowNumber=row + 1, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                   text=data['tableHeader'][count])
                    table.rows[row + 1].cells[cell].paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row + 1, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                cell += 1
            row += 2
            count = 0
        else:
            cell = 0
            # while cell<len(table.rows[row].cells):
            if str(rowMergeCount) in data['tableDataMergeCell'][fieldCount]:
                # if cell in data['tableDataMergeCell'][fieldCount][str(rowMergeCount)]:
                while cell < len(table.rows[row].cells):
                    if cell == 1:
                        makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=9,
                                   text=data['tableData'][count])
                    else:
                        makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                        text = data['tableData'][count])
                    table.rows[row].cells[cell].paragraphs[
                        0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell+=1
                fieldCount+=1
            else:
                cell =1
                while cell < len(table.rows[row].cells):
                    if cell == 1:
                        makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=9,
                                       text=data['tableData'][count])
                    else:
                        makeBoldItalic(table, rowNumber=row, cellNumber=cell, bold=False, italic=False, fontSize=11,
                                       text=data['tableData'][count])
                    table.rows[row].cells[cell].paragraphs[
                        0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    count += 1
                    cell+=1
            rowMergeCount+=1
            row+=1