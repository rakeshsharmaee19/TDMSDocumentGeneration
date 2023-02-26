from docx.shared import Cm
from documentCode import fillTable, dataFromJSON, styleTable
def pdHvSeriesTest(document, data={}, rowNumber = 0):
    print('PD HV Series Test')
    numberOfRowsLeft = rowNumber
    countTotalTest = 0

    while countTotalTest<len(data):
        numberOfRows = 4 + len(data[countTotalTest]['data'])
        numberOfColuns = len(data[countTotalTest]['data'][0])
        mergeCount = []
        for i in range(5, len(data[countTotalTest]['headers'][2])+1):
            if len(data[countTotalTest]['headers'][2][str(i)])>=2:
                mergeCount.append(len(data[countTotalTest]['headers'][2][str(i)]))
        if numberOfRowsLeft>25 and numberOfRows+2<numberOfRowsLeft:
            fillTable.addMiddleTable(document)
            pdHvSeriesTable = document.add_table(rows=numberOfRows, cols=numberOfColuns, style='TableGrid')
            styleTable.stylePDTest(pdHvSeriesTable, mergeRow=[])
            pdTestData = dataFromJSON.pdDvParallelTestData(data[countTotalTest])
            fillTable.fillPdHvParallelTestTable(pdHvSeriesTable, pdTestData, mergeRow=mergeCount)
            numberOfRowsLeft -= (numberOfRows + 1+2)
        else:
            fillTable.addMiddleTable(document, row=numberOfRowsLeft)
            numberOfRowsLeft=35
            pdHvSeriesTable = document.add_table(rows=numberOfRows, cols=numberOfColuns, style="TableGrid")
            styleTable.stylePDTest(pdHvSeriesTable, mergeRow=mergeCount)
            pdTestData  = dataFromJSON.pdDvParallelTestData(data[countTotalTest])
            fillTable.fillPdHvParallelTestTable(pdHvSeriesTable, pdTestData, mergeRow = mergeCount)
            numberOfRowsLeft-=(numberOfRows+1)
        countTotalTest+=1
    return numberOfRowsLeft
