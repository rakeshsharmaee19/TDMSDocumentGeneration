from docx.shared import Cm
from documentCode import fillTable, dataFromJSON, styleTable


def thermalTest(document, data={}, rowNumber=0):
    print('Thermal Test')
    numberOfRowsLeft = rowNumber
    countTotalTest = 0
    if numberOfRowsLeft != 35:
        fillTable.addMiddleTable(document, row=numberOfRowsLeft - 2)
        numberOfRowsLeft = 35
    thermalTestData = dataFromJSON.thermalTestData(data[0])
    textTableRows = 1 + len(thermalTestData['textHeight'])
    # tableRows = len(data[0]['Table']['headers']) + (len(thermalTestData['tableDataMerge']) * 2) - 2
    # tableColumns = len(data[0]['Table']['data'][0])
    tableRows = len(data[0]['Table']['headers'])+(len(thermalTestData['tableDataMerge'])*2)-2
    tableColumns = len(data[0]['Table']['data'][0])
    textTable = document.add_table(rows=textTableRows, cols=1)
    styleTable.styleTextTable(document, textTable, data=thermalTestData['textHeight'])
    fillTable.filltextTable(textTable, data=thermalTestData['textHeader'] + thermalTestData['textData'])
    fillTable.addMiddleTable(document, row=1)
    thermalTable = document.add_table(rows=tableRows, cols=tableColumns, style='TableGrid')
    styleTable.styleThermalTestTable(thermalTable, data=thermalTestData, numberOfColumns=tableColumns,
                                     numberOfHeaderRow=len(data[0]['Table']['headers']))
    fillTable.fillThermalTestData(thermalTable, thermalTestData, len(data[0]['Table']['headers']))
    rowValue = 2 +sum(thermalTestData['textHeight']) + 6+(len(thermalTestData['tableDataMerge'])*2)
    numberOfRowsLeft-=rowValue
    return numberOfRowsLeft
