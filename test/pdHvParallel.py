from docx.shared import Cm
from documentCode import fillTable, dataFromJSON, styleTable
def pdHvParallelTest(document, data={}, rowNumber = 0):
    print('PD HV Parallel Test')
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
            pdHvParallelTable = document.add_table(rows=numberOfRows, cols=numberOfColuns, style='TableGrid')
            styleTable.stylePDTest(pdHvParallelTable, mergeRow=[])
            pdTestData = dataFromJSON.pdDvParallelTestData(data[countTotalTest])
            fillTable.fillPdHvParallelTestTable(pdHvParallelTable, pdTestData, mergeRow=mergeCount)
            numberOfRowsLeft-=((numberOfRows+1)+3)
        else:
            fillTable.addMiddleTable(document, row=numberOfRowsLeft)
            numberOfRowsLeft=35
            pdHvParallelTable = document.add_table(rows=numberOfRows, cols=numberOfColuns, style="TableGrid")
            styleTable.stylePDTest(pdHvParallelTable, mergeRow=mergeCount)
            pdTestData  = dataFromJSON.pdDvParallelTestData(data[countTotalTest])
            fillTable.fillPdHvParallelTestTable(pdHvParallelTable, pdTestData, mergeRow = mergeCount)
            numberOfRowsLeft-=(numberOfRows+1)
        countTotalTest+=1

    return numberOfRowsLeft
