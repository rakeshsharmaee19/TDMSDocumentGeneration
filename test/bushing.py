from documentCode import styleTable, dataFromJSON, fillTable

def bushingTest(document, data, i):
    countTotalTest = 0
    while countTotalTest < len(data['DATA'][i]):
        numberOfRows = 3 + len(data['DATA'][i][countTotalTest]['bushingData'])
        numOfColumns = len(data['DATA']['bushing'][countTotalTest]['bushingData'][0])
        rowHeight = 0.5
        if numOfColumns > 10:
            rowHeight = 1
        merge_list = []
        for k in range(5, len(data['DATA']['bushing'][0]["bushingHeaders"][1]) + 1):
            merge_list.append(len(data['DATA']['bushing'][0]["bushingHeaders"][1][str(k)]['2']))

        table = document.add_table(rows=numberOfRows, cols=numOfColumns, style='TableGrid')
        styleTable.styleBushingTable(table, merge_list=merge_list, num_of_columns=numOfColumns)
        bushingTestData = dataFromJSON.bushingTestData(data['DATA'][i][countTotalTest])
        fillTable.fillBushingTest(table, bushingTestData, merge_list, row=rowHeight)

        if rowHeight == 1:
            fillTable.addMiddleTable(document, row=(35 - (numberOfRows * 2) + 1))
        else:
            fillTable.addMiddleTable(document, row=(35 - (numberOfRows)))
        countTotalTest += 1