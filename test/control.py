from documentCode import fillTable, dataFromJSON, styleTable

def controlTest(document, data, rowNumber=0):
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['control']):
        numberOfRows = 3 + len(data['DATA']['control'][countTotalTest]['data'])
        numberOfColumns = len(data['DATA']['control'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            controlTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            controlTestData = dataFromJSON.controlTestData(data['DATA']['control'][countTotalTest])
            styleTable.styleControlTable(controlTestTable)
            fillTable.fillControlTestTable(controlTestTable, controlTestData)
            document.add_table(rows=1, cols=1)  # will remove after adding more test

            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)

                controlTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                controlTestData = dataFromJSON.controlTestData(data['DATA']['control'][countTotalTest])
                styleTable.styleControlTable(controlTestTable)
                fillTable.fillControlTestTable(controlTestTable, controlTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                controlTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                controlTestData = dataFromJSON.controlTestData(data['DATA']['control'][countTotalTest])
                styleTable.styleControlTable(controlTestTable)
                fillTable.fillControlTestTable(controlTestTable, controlTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test

                numberOfRowsLeft -= numberOfRows


        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
