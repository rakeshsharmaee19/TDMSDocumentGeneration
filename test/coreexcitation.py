from documentCode import fillTable, dataFromJSON, styleTable


def coreExcitation1(document, data, rowNumber=0):
    print("core excitation table 1")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['coreExcitationTable1']):
        numberOfRows = len(data['DATA']['coreExcitationTable1'][countTotalTest]['data']) + 3
        numberOfColumns = len(data['DATA']['coreExcitationTable1'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            coreExcitationTestTable1 = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                          style='TableGrid')
            coreExcitationTable1TestData = dataFromJSON.coreExcitationStateTestData(
                data['DATA']['coreExcitationTable1'])

            mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                data['DATA']['coreExcitationTable1'])
            styleTable.styleCoreExcitationStateTestTable(coreExcitationTestTable1, numberOfColumns,
                                                         mergeListHorizontalHeader, mergeListVerticalHeader)
            data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
            cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                       numberOfColumns)
            fillTable.fillCoreExcitationStateTestData(coreExcitationTestTable1,
                                                      coreExcitationTable1TestData, data_merge,
                                                      cellToInsertCoreExcitationHeader)
            document.add_table(rows=1, cols=1)  # will remove after adding more test
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                coreExcitationTestTable1 = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                              style='TableGrid')
                coreExcitationTable1TestData = dataFromJSON.coreExcitationStateTestData(
                    data['DATA']['coreExcitationTable1'])

                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['coreExcitationTable1'])
                styleTable.styleCoreExcitationStateTestTable(coreExcitationTestTable1, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
                cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                           numberOfColumns)
                fillTable.fillCoreExcitationStateTestData(coreExcitationTestTable1,
                                                          coreExcitationTable1TestData, data_merge,
                                                          cellToInsertCoreExcitationHeader)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                coreExcitationTestTable1 = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                              style='TableGrid')
                coreExcitationTable1TestData = dataFromJSON.coreExcitationStateTestData(
                    data['DATA']['coreExcitationTable1'])

                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['coreExcitationTable1'])
                styleTable.styleCoreExcitationStateTestTable(coreExcitationTestTable1, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
                cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                           numberOfColumns)
                fillTable.fillCoreExcitationStateTestData(coreExcitationTestTable1,
                                                          coreExcitationTable1TestData, data_merge,
                                                          cellToInsertCoreExcitationHeader)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft

def coreExcitation2(document, data, rowNumber=0):
    print("core excitation table 2")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['coreExcitationTable2']):
        numberOfRows = len(data['DATA']['coreExcitationTable2'][countTotalTest]['data']) + 3
        numberOfColumns = len(data['DATA']['coreExcitationTable2'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            CoreExcitationTestTable2 = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            CoreExcitationTable2TestData = dataFromJSON.coreExcitationStateTestData(
                data['DATA']['coreExcitationTable2'])
            mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                data['DATA']['coreExcitationTable2'])
            styleTable.styleCoreExcitationStateTestTable(CoreExcitationTestTable2, numberOfColumns,
                                                         mergeListHorizontalHeader, mergeListVerticalHeader)
            data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
            cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                       numberOfColumns)
            fillTable.fillCoreExcitationStateTestData(CoreExcitationTestTable2, CoreExcitationTable2TestData,
                                                      data_merge, cellToInsertCoreExcitationHeader)
            document.add_table(rows=1, cols=1)  # will remove after adding more test
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                CoreExcitationTestTable2 = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                              style='TableGrid')
                CoreExcitationTable2TestData = dataFromJSON.coreExcitationStateTestData(
                    data['DATA']['coreExcitationTable2'])
                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['coreExcitationTable2'])
                styleTable.styleCoreExcitationStateTestTable(CoreExcitationTestTable2, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
                cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                           numberOfColumns)
                fillTable.fillCoreExcitationStateTestData(CoreExcitationTestTable2, CoreExcitationTable2TestData,
                                                          data_merge, cellToInsertCoreExcitationHeader)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                CoreExcitationTestTable2 = document.add_table(rows=numberOfRows, cols=numberOfColumns,
                                                              style='TableGrid')
                CoreExcitationTable2TestData = dataFromJSON.coreExcitationStateTestData(
                    data['DATA']['coreExcitationTable2'])
                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['coreExcitationTable2'])
                styleTable.styleCoreExcitationStateTestTable(CoreExcitationTestTable2, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                data_merge = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
                cellToInsertCoreExcitationHeader = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader,
                                                                                           numberOfColumns)
                fillTable.fillCoreExcitationStateTestData(CoreExcitationTestTable2, CoreExcitationTable2TestData,
                                                          data_merge, cellToInsertCoreExcitationHeader)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
