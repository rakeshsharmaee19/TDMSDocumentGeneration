from documentCode import fillTable, styleTable, dataFromJSON

def excitationTest(document, data, i):
    print('excitation')
    countTotalTest = 0
    numberOfRowsLeft = 35
    if len(data['DATA'][i]) > 0:
        numberOfRows = 2 + len(data['DATA'][i]['basicTempData']['basicsTemplateDataList'])
        numberOfColumns = len(data['DATA'][i]['basicTempData']['basicsTemplateDataList'][0])
        excitationTestData = dataFromJSON.excitationTestData(data['DATA'][i])
        excitationTestTable = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
        styleTable.styleExcitationTestTable(excitationTestTable, numberOfColumns)
        fillTable.fillExcitationTest(excitationTestTable, excitationTestData, 0.5)
        rowValue = numberOfRows
        numberOfRowsLeft = numberOfRowsLeft - rowValue
        while countTotalTest < len(data['DATA'][i]['listOfExciData']):
            if countTotalTest == len(data['DATA'][i]['listOfExciData']) - 1 and countTotalTest == 0:
                numberOfRows = 3 + len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'])
                numberOfColumns = len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'][0])
                excitationData = dataFromJSON.excitationData(
                    data['DATA'][i]['listOfExciData'][countTotalTest])
                if numberOfRows + 2 <= numberOfRowsLeft:
                    fillTable.addMiddleTable(document, row=2)
                    rowValue += 2 + numberOfRows

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)

                    fillTable.addMiddleTable(document, row=(35 - rowValue))
                else:
                    fillTable.addMiddleTable(document, row=(35 - rowValue))
                    rowValue = 0
                    numberOfRowsLeft = 35

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)
                    rowValue += numberOfRows

                    fillTable.addMiddleTable(document, row=(35 - rowValue))

            elif countTotalTest == len(data['DATA'][i]) - 1:
                excitationData = dataFromJSON.excitationData(
                    data['DATA'][i]['listOfExciData'][countTotalTest])
                numberOfRows = 3 + len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'])
                numberOfColumns = len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'][0])
                if numberOfRows + 2 <= numberOfRowsLeft:
                    fillTable.addMiddleTable(document)
                    rowValue += 2 + numberOfRows

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)

                    fillTable.addMiddleTable(document, row=(35 - rowValue))

                else:
                    fillTable.addMiddleTable(document, row=(35 - rowValue))
                    rowValue = 0
                    numberOfRowsLeft = 35

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)
                    rowValue += numberOfRows

                    fillTable.addMiddleTable(document, row=(35 - rowValue))

            elif countTotalTest == 0:
                excitationData = dataFromJSON.excitationData(
                    data['DATA'][i]['listOfExciData'][countTotalTest])
                numberOfRows = 3 + len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'])
                numberOfColumns = len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'][0])
                if numberOfRows + 2 <= numberOfRowsLeft:
                    fillTable.addMiddleTable(document)
                    rowValue += 2 + numberOfRows

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)
                    if rowValue == numberOfRowsLeft:
                        rowValue = 0
                        numberOfRowsLeft = 35
                    else:
                        numberOfRowsLeft = numberOfRowsLeft - numberOfRows - 2
                else:
                    fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                    rowValue = 0
                    numberOfRowsLeft = 35

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)
                    rowValue += numberOfRows

                    if rowValue == numberOfRowsLeft:
                        rowValue = 0
                        numberOfRowsLeft = 35
                    else:
                        numberOfRowsLeft = numberOfRowsLeft - numberOfRows - 2
            else:
                numberOfRows = 3 + len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'])
                numberOfColumns = len(
                    data['DATA'][i]['listOfExciData'][countTotalTest]['excitingDataList'][0])
                excitationData = dataFromJSON.excitationData(
                    data['DATA'][i]['listOfExciData'][countTotalTest])
                if numberOfRows + 2 <= numberOfRowsLeft:
                    fillTable.addMiddleTable(document)

                    rowValue += 2 + numberOfRows

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)

                    if rowValue == numberOfRowsLeft:
                        rowValue = 0
                        numberOfRowsLeft = 35
                    else:
                        numberOfRowsLeft = numberOfRowsLeft - numberOfRows - 2
                else:
                    fillTable.addMiddleTable(document, row=numberOfRowsLeft)
                    rowValue = 0
                    numberOfRowsLeft = 35

                    merge_list = []
                    for j in range(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                        'excitingHeadersDataList'][1]))):
                        merge_list.append(len((data['DATA'][i]['listOfExciData'][countTotalTest][
                            'excitingHeadersDataList'][1][str(j + 1)]['2'])))
                    table = document.add_table(rows=numberOfRows, cols=numberOfColumns, style='TableGrid')
                    styleTable.styleExcitationTable(table, numberOfColumns, merge_list)
                    fillTable.fillExcitationData(table, excitationData, merge_list)
                    rowValue += numberOfRows

                    if rowValue == numberOfRowsLeft:
                        rowValue = 0
                        numberOfRowsLeft = 35
                    else:
                        numberOfRowsLeft = numberOfRowsLeft - numberOfRows - 2
            countTotalTest += 1
    else:
        pass