from documentCode import fillTable, dataFromJSON, styleTable
def insulationResistanceTest(document, data, rowNumber=0):
    print("INSULATION RESISTANCE")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['insulatinResistance']):
        numberOfRows = len(data['DATA']['insulatinResistance'][countTotalTest]['data']) + len(data['DATA']['insulatinResistance'][countTotalTest]['headers'])
        numberOfColumns = len(data['DATA']['insulatinResistance'][countTotalTest]['data'][0])
        numOfHeaders = len(data['DATA']['insulatinResistance'][0]['headers'])
        if numberOfRowsLeft == 35:
            InsulationResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                               style='TableGrid')
            InsulationResistenceTestData = dataFromJSON.insulationResistanceTableTestData(data['DATA']
                                                                                          ['insulatinResistance']
                                                                                          [countTotalTest])
            styleTable.styleInsulationResistanceTable(InsulationResistenceTestTable, numberOfColumns, numOfHeaders)
            fillTable.fillinsulationResistanceFillData(InsulationResistenceTestTable,
                                                       InsulationResistenceTestData)
            document.add_table(rows=1, cols=1) # will remove after adding more test

            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                InsulationResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                   style='TableGrid')
                InsulationResistenceTestData = dataFromJSON.insulationResistanceTableTestData(data['DATA']
                                                                                              ['insulatinResistance']
                                                                                              [countTotalTest])
                styleTable.styleInsulationResistanceTable(InsulationResistenceTestTable, numberOfColumns, numOfHeaders)
                fillTable.fillinsulationResistanceFillData(InsulationResistenceTestTable,
                                                           InsulationResistenceTestData)
                numberOfRowsLeft-= numberOfRows
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                InsulationResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                   style='TableGrid')
                InsulationResistenceTestData = dataFromJSON.insulationResistanceTableTestData(data['DATA']
                                                                                              ['insulatinResistance']
                                                                                              [countTotalTest])
                styleTable.styleInsulationResistanceTable(InsulationResistenceTestTable, numberOfColumns, numOfHeaders)
                fillTable.fillinsulationResistanceFillData(InsulationResistenceTestTable,
                                                           InsulationResistenceTestData)
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
