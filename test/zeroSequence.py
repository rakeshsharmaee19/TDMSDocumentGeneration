from documentCode import fillTable, dataFromJSON, styleTable


def zeroSequenceTest(document, data, rowNumber=0):
    print("Zero Sequence Test")
    countTotalTest = 0
    numberOfRowsLeft = rowNumber
    while countTotalTest < len(data['DATA']['zeroSequence']):
        numberOfRows = len(data['DATA']['zeroSequence'][countTotalTest]['data']) + 3
        numberOfColumns = len(data['DATA']['zeroSequence'][countTotalTest]['data'][0])
        if numberOfRowsLeft == 35:
            zeroSequenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
            zeroSequenceTestData = dataFromJSON.zeroSequenceTableTestData(data['DATA']['zeroSequence'])
            mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(data['DATA']['zeroSequence'])
            styleTable.styleCoreExcitationStateTestTable(zeroSequenceTestTable, numberOfColumns,mergeListHorizontalHeader,mergeListVerticalHeader)
            cell_data_row2 = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numberOfColumns)
            cell_data_row3 = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader, numberOfColumns)
            fillTable.fillZeroSequence(zeroSequenceTestTable, zeroSequenceTestData, cell_data_row2, cell_data_row3)
            document.add_table(rows=1, cols=1)  # will remove after adding more test
            numberOfRowsLeft -= numberOfRows
        else:
            if numberOfRows + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document)
                zeroSequenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                zeroSequenceTestData = dataFromJSON.zeroSequenceTableTestData(data['DATA']['zeroSequence'])
                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['zeroSequence'])
                styleTable.styleCoreExcitationStateTestTable(zeroSequenceTestTable, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                cell_data_row2 = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader,
                                                                                numberOfColumns)
                cell_data_row3 = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader, numberOfColumns)
                fillTable.fillZeroSequence(zeroSequenceTestTable, zeroSequenceTestData, cell_data_row2, cell_data_row3)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
            else:
                fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                numberOfRowsLeft = 35
                zeroSequenceTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                zeroSequenceTestData = dataFromJSON.zeroSequenceTableTestData(data['DATA']['zeroSequence'])
                mergeListHorizontalHeader, mergeListVerticalHeader = dataFromJSON.mergeListCoreExcitation(
                    data['DATA']['zeroSequence'])
                styleTable.styleCoreExcitationStateTestTable(zeroSequenceTestTable, numberOfColumns,
                                                             mergeListHorizontalHeader, mergeListVerticalHeader)
                cell_data_row2 = dataFromJSON.coreExcitationDataInsertMergeList(mergeListHorizontalHeader,
                                                                                numberOfColumns)
                cell_data_row3 = dataFromJSON.cellToInsertCoreExcitation(mergeListVerticalHeader, numberOfColumns)
                fillTable.fillZeroSequence(zeroSequenceTestTable, zeroSequenceTestData, cell_data_row2, cell_data_row3)
                document.add_table(rows=1, cols=1)  # will remove after adding more test
                numberOfRowsLeft -= numberOfRows
        countTotalTest += 1
        if numberOfRowsLeft == 0:
            return 35
        else:
            return numberOfRowsLeft
