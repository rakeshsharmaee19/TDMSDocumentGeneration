from documentCode import fillTable, dataFromJSON, styleTable

def dewPointTest(document, data, rowNumber=0):
    print('Dew Poni Test')
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data):
        numberOfRows = len(data[countTotalTest]['headers']) + len(data[countTotalTest]['data'])
        numberOfColumns = len(data[countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            dewpointTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            dewpointTestData = dataFromJSON.dewpointTestData(data[countTotalTest])
            styleTable.styleDewPointTestTable(dewpointTestTable)
            print(dewpointTestData)
            fillTable.fillDewPointTestTable(dewpointTestTable, dewpointTestData)
            document.add_table(rows=1, cols=1)  # will remove after adding more test
            # document.add_table(rows=1, cols=1) # will remove after adding more test

            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)

                dewpointTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                dewpointTestData = dataFromJSON.dewpointTestData(data[countTotalTest])
                styleTable.styleDewPointTestTable(dewpointTestTable)
                print(dewpointTestData)
                fillTable.fillDewPointTestTable(dewpointTestTable, dewpointTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                # document.add_table(rows=1, cols=1)   #will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                dewpointTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                dewpointTestData = dataFromJSON.dewpointTestData(data[countTotalTest])

                styleTable.styleDewPointTestTable(dewpointTestTable)
                print(dewpointTestData)
                fillTable.fillDewPointTestTable(dewpointTestTable, dewpointTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1


        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
    dewPointTestData = dataFromJSON.dewpointTestData(data[0])
    print(dewPointTestData)