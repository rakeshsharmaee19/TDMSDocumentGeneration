from documentCode import fillTable, dataFromJSON, styleTable
def voltageParallelTest(document, data, rowNumber=0):
    print("VoltageParallelTest")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['voltageParallel']):
        numberOfRows = 2 + len(data['DATA']['voltageParallel'][countTotalTest]['data'])
        numberOfColumns = len(data['DATA']['voltageParallel'][countTotalTest]['data'][countTotalTest])
        if numberOfRowsLeft == 35:
            ParallelInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                 style='TableGrid')
            parallelInducedVoltageTestData = dataFromJSON.parallelInducedVoltageTestData(data['DATA']['voltageParallel']
                                                                                         [countTotalTest])
            styleTable.styleParallelInducedVoltageTestTable(ParallelInducedVoltageTestTable, numberOfColumns)
            fillTable.fillParallelInducedVoltageTestData(ParallelInducedVoltageTestTable,
                                                         parallelInducedVoltageTestData)
            # document.add_table(rows=1, cols=1) # will remove after adding more test
            #
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                ParallelInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                     style='TableGrid')
                parallelInducedVoltageTestData = dataFromJSON.parallelInducedVoltageTestData(
                    data['DATA']['voltageParallel']
                    [countTotalTest])
                styleTable.styleParallelInducedVoltageTestTable(ParallelInducedVoltageTestTable, numberOfColumns)
                fillTable.fillParallelInducedVoltageTestData(ParallelInducedVoltageTestTable,
                                                             parallelInducedVoltageTestData)
                # document.add_table(rows=1, cols=1)   #will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                ParallelInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                     style='TableGrid')
                parallelInducedVoltageTestData = dataFromJSON.parallelInducedVoltageTestData(
                    data['DATA']['voltageParallel']
                    [countTotalTest])
                styleTable.styleParallelInducedVoltageTestTable(ParallelInducedVoltageTestTable, numberOfColumns)
                fillTable.fillParallelInducedVoltageTestData(ParallelInducedVoltageTestTable,
                                                             parallelInducedVoltageTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft


def voltageSeriesTest(document, data, rowNumber=0):
    print("VoltageSeriesTest")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['voltageSeries']):
        numberOfRows = 2 + len(data['DATA']['voltageSeries'][countTotalTest]['data'])
        numberOfColumns = len(data['DATA']['voltageSeries'][countTotalTest]['data'][countTotalTest])
        if numberOfRowsLeft == 35:
            SeriesInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                               style='TableGrid')
            SeriesInducedVoltageTestData = dataFromJSON.seriesInducedVoltageTestData(data['DATA']['voltageSeries']
                                                                                     [countTotalTest])
            styleTable.styleSeriesInducedVoltageTestTable(SeriesInducedVoltageTestTable, numberOfColumns)
            fillTable.fillParallelInducedVoltageTestData(SeriesInducedVoltageTestTable,
                                                         SeriesInducedVoltageTestData)
            # document.add_table(rows=1, cols=1) # will remove after adding more test
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                SeriesInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                   style='TableGrid')
                SeriesInducedVoltageTestData = dataFromJSON.seriesInducedVoltageTestData(data['DATA']['voltageSeries']
                                                                                         [countTotalTest])
                styleTable.styleSeriesInducedVoltageTestTable(SeriesInducedVoltageTestTable, numberOfColumns)
                fillTable.fillParallelInducedVoltageTestData(SeriesInducedVoltageTestTable,
                                                             SeriesInducedVoltageTestData)
    #             # document.add_table(rows=1, cols=1)   #will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                SeriesInducedVoltageTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                                   style='TableGrid')
                SeriesInducedVoltageTestData = dataFromJSON.seriesInducedVoltageTestData(data['DATA']['voltageSeries']
                                                                                         [countTotalTest])
                styleTable.styleSeriesInducedVoltageTestTable(SeriesInducedVoltageTestTable, numberOfColumns)
                fillTable.fillParallelInducedVoltageTestData(SeriesInducedVoltageTestTable,
                                                             SeriesInducedVoltageTestData)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
