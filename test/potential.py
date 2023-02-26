from documentCode import fillTable, dataFromJSON, styleTable

def potencialTest(document, data, rowNumber=0):
    print("potencialTest")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['potencialTest'][0]):
        numberOfRows = len(data['DATA']["potencialTest"][countTotalTest]['data']) + 2
        numberOfColumns = len(data['DATA']["potencialTest"][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            potencialTestData = dataFromJSON.potencialTestData(data['DATA']["potencialTest"][countTotalTest])
            if len(potencialTestData[0][0]) > 82:
                numberOfRowsLeft = numberOfRowsLeft - len(potencialTestData[0][0])//82
            potencialTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            styleTable.stylePotencialTestTable(potencialTestTable, numberOfColumns)
            fillTable.fillPotencialTestTable(potencialTestTable, potencialTestData)
            # document.add_table(rows=1, cols=1) # will remove after adding more test
            numberOfRowsLeft -= numberOfRows

        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                potencialTestData = dataFromJSON.potencialTestData(data['DATA']["potencialTest"][countTotalTest])
                if len(potencialTestData[0][0]) > 82:
                    numberOfRowsLeft = numberOfRowsLeft - - len(potencialTestData[0][0])//82
                potencialTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                styleTable.stylePotencialTestTable(potencialTestTable, numberOfColumns)
                fillTable.fillPotencialTestTable(potencialTestTable, potencialTestData)
                # document.add_table(rows=1, cols=1)   #will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                potencialTestData = dataFromJSON.potencialTestData(data['DATA']["potencialTest"][countTotalTest])
                if len(potencialTestData[0][0]) > 82:
                    numberOfRowsLeft = numberOfRowsLeft - len(potencialTestData[0][0])//82
                potencialTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                styleTable.stylePotencialTestTable(potencialTestTable, numberOfColumns)
                fillTable.fillPotencialTestTable(potencialTestTable, potencialTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        fillTable.addMiddleTable(document)
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
