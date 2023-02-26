from documentCode import fillTable, dataFromJSON, styleTable
def impulseTest(document, data, rowNumber= 0):

    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['impulseTests']):
        numberOfRows = len(data['DATA']['impulseTests'][countTotalTest]['data']) + 2
        numberOfColumns = len(data['DATA']['impulseTests'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            impulseTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            impulseTestData = dataFromJSON.impulseTestData(data['DATA']["impulseTests"][countTotalTest])
            styleTable.styleImpulseTestTable(impulseTestTable,numberOfColumns)
            fillTable.fillImpulseTestTable(impulseTestTable, impulseTestData)
            # document.add_table(rows=1, cols=1) # will remove after adding more test

            numberOfRowsLeft -= numberOfRows

        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                impulseTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                impulseTestData = dataFromJSON.impulseTestData(data['DATA']["impulseTests"][countTotalTest])
                styleTable.styleImpulseTestTable(impulseTestTable,numberOfColumns)
                fillTable.fillImpulseTestTable(impulseTestTable, impulseTestData)

                # document.add_table(rows=1, cols=1)   #will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                impulseTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                impulseTestData = dataFromJSON.impulseTestData(data['DATA']["impulseTests"][countTotalTest])
                styleTable.styleImpulseTestTable(impulseTestTable,numberOfColumns)
                fillTable.fillImpulseTestTable(impulseTestTable, impulseTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
    if numberOfRowsLeft == 0:
        return 35
    else:
        return numberOfRowsLeft

