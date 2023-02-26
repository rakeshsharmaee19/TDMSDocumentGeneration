from documentCode import dataFromJSON, styleTable, fillTable
def overallTest(document, data, i ):
    countTotalTest = 0
    numberOfRowsLeft = 35
    rowValue = 0
    while countTotalTest < len(data['DATA'][i]):
        if countTotalTest == len(data['DATA'][i]) - 1 and countTotalTest == 0:
            overAllTestData = dataFromJSON.overAllTestData(data['DATA'][i][countTotalTest])
            rowValue = rowValue + 3 + len(data['DATA'][i][countTotalTest]['OverallData']) * overAllTestData[
                2]
            numberOfRows = 3 + len(data['DATA'][i][countTotalTest]['OverallData'])

            table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
            styleTable.styeleOverAllTable(table)
            fillTable.fillOverAllTest(table, overAllTestData)

            fillTable.addMiddleTable(document, row=(35 - rowValue))

        elif countTotalTest == len(data['DATA'][i]) - 1:
            overAllTestData = dataFromJSON.overAllTestData(data['DATA'][i][countTotalTest])
            heightValue = 3 + len(data['DATA'][i][countTotalTest]['OverallData']) * overAllTestData[
                2]
            numberOfRows = 3 + len(data['DATA'][i][countTotalTest]['OverallData'])
            if heightValue + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document, row=2)

                rowValue += 2 + heightValue
                table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
                styleTable.styeleOverAllTable(table)
                fillTable.fillOverAllTest(table, overAllTestData)

                fillTable.addMiddleTable(document, row=(35 - rowValue))

            else:
                fillTable.addMiddleTable(document, row=(35 - rowValue))

                rowValue = 0
                numberOfRowsLeft = 35
                table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
                styleTable.styeleOverAllTable(table)
                fillTable.fillOverAllTest(table, overAllTestData)
                rowValue += heightValue

                fillTable.addMiddleTable(document, row=(35 - rowValue))

        elif countTotalTest == 0:
            overAllTestData = dataFromJSON.overAllTestData(data['DATA'][i][countTotalTest])
            heightValue = 3 + len(data['DATA'][i][countTotalTest]['OverallData']) * overAllTestData[
                2]
            numberOfRows = 3 + len(data['DATA'][i][countTotalTest]['OverallData'])
            table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
            styleTable.styeleOverAllTable(table)
            fillTable.fillOverAllTest(table, overAllTestData)
            rowValue = rowValue + heightValue
            if rowValue == numberOfRowsLeft:
                rowValue = 0
                numberOfRowsLeft = 35
            else:
                numberOfRowsLeft = numberOfRowsLeft - rowValue
        else:
            overAllTestData = dataFromJSON.overAllTestData(data['DATA'][i][countTotalTest])
            heightValue = 3 + len(data['DATA'][i][countTotalTest]['OverallData']) * overAllTestData[
                2]
            numberOfRows = 3 + len(data['DATA'][i][countTotalTest]['OverallData'])
            if heightValue + 2 <= numberOfRowsLeft:
                fillTable.addMiddleTable(document, row=2)

                rowValue += 2 + heightValue
                table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
                styleTable.styeleOverAllTable(table)
                fillTable.fillOverAllTest(table, overAllTestData)
                if rowValue == 35:
                    rowValue = 0
                    numberOfRowsLeft = 35
                else:
                    numberOfRowsLeft = 35 - rowValue
            else:
                fillTable.addMiddleTable(document, row=numberOfRows)

                rowValue = 0
                numberOfRowsLeft = 35
                heightValue = + 3 + len(data['DATA'][i][countTotalTest]['OverallData']) * overAllTestData[2]
                table = document.add_table(rows=numberOfRows, cols=6, style='TableGrid')
                styleTable.styeleOverAllTable(table)
                fillTable.fillOverAllTest(table, overAllTestData)
                rowValue += heightValue
                if rowValue == numberOfRowsLeft:
                    rowValue = 0
                    numberOfRowsLeft = 35
                else:
                    numberOfRowsLeft = numberOfRowsLeft - rowValue
        countTotalTest += 1
