from documentCode import fillTable, dataFromJSON, styleTable
def leadResistanceTest(document, data, rowNumber=0):
    print('leadResistanceTest')
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['leadResistance']):
        numberOfRows = len(data['DATA']['leadResistance'][countTotalTest]['data']) + 2
        numberOfColumns = len(data['DATA']['leadResistance'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            leadResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            leadResistenceTestData = dataFromJSON.leadResistanceTableTestData(data['DATA']['leadResistance'][countTotalTest])
            styleTable.styleLeadResistance(leadResistenceTestTable, numberOfColumns)
            fillTable.fillinsulationResistanceFillData(leadResistenceTestTable, leadResistenceTestData)
            document.add_table(rows=1, cols=1)  # will remove after adding more test
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                leadResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                leadResistenceTestData = dataFromJSON.leadResistanceTableTestData(data['DATA']['leadResistance'][countTotalTest])
                styleTable.styleLeadResistance(leadResistenceTestTable, numberOfColumns)
                fillTable.fillinsulationResistanceFillData(leadResistenceTestTable, leadResistenceTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                leadResistenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                leadResistenceTestData = dataFromJSON.leadResistanceTableTestData(data['DATA']['leadResistance'][countTotalTest])
                styleTable.styleLeadResistance(leadResistenceTestTable, numberOfColumns)
                fillTable.fillinsulationResistanceFillData(leadResistenceTestTable, leadResistenceTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
