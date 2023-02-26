from math import ceil
from .styleTable import *


### <<<< ------------------------------ Generate Graph Data ---------------------------------- >>>>
def graphData(data, windingKey):
    datawinding = {}
    for i in windingKey:
        datawinding[i] = []
    dataKeys = datawinding.keys()
    for i in data['DATA']:
        # countTrue = True
        dataBlue = []
        dataGreen = []
        for j in data['DATA'][i]["calculatedValues"]:
            # if j['vindicationFactor'] == False and countTrue:
            if (j['extraPolatedResistance']) != None:
                dataBlue.append(j['extraPolatedResistance'])
            else:
                dataBlue.append(0)
            if j['vindicationFactor'] == True:
                if j['hotRExtraPolated'] != None:
                    dataGreen.append(j['hotRExtraPolated'])
                else:
                    dataGreen.append(0)
            else:
                dataGreen.append(None)
        if i in dataKeys:
            datawinding[i].append(dataBlue)
            datawinding[i].append(dataGreen)
        else:
            pass
    return datawinding


def overAllTestData(data):
    returnData = []
    fieldData = []
    headerData = []
    rowMerge = 1
    for i in range(len(data['OverallHeaders'])):
        if i == 0:
            headerData.append(data['OverallHeaders'][i]['1'])
        else:
            headerRowKey = [k for k in data['OverallHeaders'][i].keys()]
            for j in headerRowKey:
                if j == '5':
                    headerData.append(data['OverallHeaders'][i][j]['1'])
                    headerData.append(data['OverallHeaders'][i][j]['2'][0])
                    headerData.append(data['OverallHeaders'][i][j]['2'][1])
                else:
                    headerData.append(data['OverallHeaders'][i][j])
    for i in data['OverallData']:
        a = ceil(len(i['2']) / 19)
        if a > rowMerge:
            rowMerge = a
        fieldData.extend(i.values())
    returnData.append(headerData)
    returnData.append(fieldData)
    returnData.append(rowMerge)
    return returnData


def bushingTestData(data):
    returnData = []
    fieldData = []
    headerData = []
    for i in range(len(data['bushingHeaders'])):
        if i == 0:
            headerData.append(data['bushingHeaders'][i]['1'])
        else:
            for j in range(1, 5):
                headerData.append(data["bushingHeaders"][i][str(j)])

            for k in range(5, len(data["bushingHeaders"][i]) + 1):
                headerData.append(data["bushingHeaders"][i][str(k)]['1'])

            for k in range(5, len(data["bushingHeaders"][i]) + 1):
                headerData.extend(data["bushingHeaders"][i][str(k)]['2'])

    for i in data['bushingData']:
        fieldData.extend(i.values())
    returnData.append(headerData)
    returnData.append(fieldData)
    return returnData


def excitationTestData(data):
    returnData = []
    fieldData = []
    headerData = []
    rowMerge = 1
    for header in data['basicTempData']['basicTemplateHeadersList']:
        headerData.extend(header.values())
    for data in data['basicTempData']['basicsTemplateDataList']:
        a = ceil(len(data['2']) / 17)
        b = ceil(len(data['1']) / 17)
        if a > rowMerge or b > rowMerge:
            if a > b:
                rowMerge = a
            else:
                rowMerge = b
        fieldData.extend(data.values())
    returnData.append(headerData)
    returnData.append(fieldData)
    returnData.append(rowMerge)
    return returnData


def excitationData(data):
    returnData = []
    fieldData = []
    headerData = []
    for i in range(len(data['excitingHeadersDataList'])):
        if i == 0:
            for k in range(len(data['excitingHeadersDataList'][i])):
                headerData.append(data['excitingHeadersDataList'][i][str(k + 1)])
        else:
            for k in range(len(data['excitingHeadersDataList'][i])):
                headerData.append(data["excitingHeadersDataList"][i][str(k + 1)]['1'])
            for k in range(len(data['excitingHeadersDataList'][i])):
                headerData.extend(data["excitingHeadersDataList"][i][str(k + 1)]['2'])
    for i in data['excitingDataList']:
        fieldData.extend(i.values())
    returnData.append(headerData)
    returnData.append(fieldData)
    return returnData


def ratioTestData(data):
    ratioHeader = []
    ratioData = []
    overallData = []

    for j in range(len(data['RatioHeader'])):
        if j in (0, 1, 2):
            ratioHeader.append(data['RatioHeader'][j]['1'])
        elif j == 3:
            ratioHeader.append(' ')
            ratioHeader.append(data['RatioHeader'][j]['2'])
            ratioHeader.append(data['RatioHeader'][j]['3'])

        elif j == 4:
            ratioHeader.append(data['RatioHeader'][j]['1']['1'])
            ratioHeader.append(data['RatioHeader'][j]['2'])
            ratioHeader.append(data['RatioHeader'][j]['3']['1'].upper())
            ratioHeader.append(data['RatioHeader'][j]['4']['1'].upper())
            ratioHeader.extend(data['RatioHeader'][j]['1']['2'])
            ratioHeader.extend(data['RatioHeader'][j]['3']['2'])
            ratioHeader.extend(data['RatioHeader'][j]['4']['2'])

    for k in range(len(data['RatioData'])):
        ratioData.extend(data['RatioData'][k].values())

    overallData.append(ratioHeader)
    overallData.append(ratioData)
    return overallData


def getTerminals(data):
    list_of_windings = {}
    for i in data:
        if data[i]['winding'] not in list_of_windings.keys():
            list_of_windings[data[i]['winding']] = [i]
        else:
            list_of_windings[data[i]['winding']].append(i)
    return list_of_windings


def dgaData(data):
    dga_headers = []
    dga_data = []
    mergeVaLue = 0

    dict_of_headers = {'sampleId': ['Sample ID', 1.2], 'TestId': ['Test ID', 2], 'Date': ['Date', 2],
                       'SyringeID/BottleId': ['Syringe ID', 2.0], 'Apparatus': ['Apparatus', 1.2],
                       'Hydrogen_H2': ['Hydrogen H2', 1.1], 'Carbon_Monoxide_CO': ['Carbon Monoxide CO', 1.1],
                       'Carbon_Dioxide_CO2': ['Carbon Dioxide CO2', 1.1], 'Methane_CH4': ['Methane CH4', 1.1],
                       'Ethane_C2H6': ['Ethane C2H6', 1.1], 'Ethylene_C2H4': ['Ethylene C2H4', 1.1],
                       'Acetylene_C2H2': ['Acetylene C2H2', 1.1],
                       'Moisture_H2O': ['Moisture H2O', 1.1], 'Temp_C': ['Temp °C', 1.2]}

    # dict_of_headers = {'CycleId': ['DGA #', 1.2], "sampleId": ["sampleId", 1], 'TestId': ['Test ID', 2],
    #                    'Date': ['Date', 2],
    #                    'SyringeID/BottleId': ['Syringe ID', 2.0], 'Apparatus': ['Apparatus', 1.2],
    #                    'Hydrogen_H2': ['Hydrogen H2', 1.1], 'Carbon_Monoxide_CO': ['Carbon Monoxide CO', 1.1],
    #                    'Carbon_Dioxide_CO2': ['Carbon Dioxide CO2', 1.1], 'Methane_CH4': ['Methane CH4', 1.1],
    #                    'Ethane_C2H6': ['Ethane C2H6', 1.1], 'Ethylene_C2H4': ['Ethylene C2H4', 1.1],
    #                    'Acetylene_C2H2': ['Acetylene C2H2', 1.1],
    #                    'Moisture_H2O': ['Moisture H2O', 1.1], 'Temp_C': ['Temp °C', 1.2],
    #                    'Oxygen_O2': ['Oxygen O2', 1.3],
    #                    'Nitrogen_N2': ['Nitrogen N2', 1.3], 'TDG': ['TDG', 1.7], 'TDCG': ['TDCG', 1.8],
    #                    'TDG_%': ['TDG %', 1.6], 'BDV-2mm': ['BDV-2mm', 1.2]}

    for i in data['DATA'][0]:
        dga_headers.append(i)
    #
    dga_headers2 = dga_headers[:]
    #
    for i in range(len(dga_headers)):
        if dga_headers[i] in dict_of_headers.keys():
            dga_headers[i] = dict_of_headers[dga_headers[i]][0]

    # s = 0
    # for i in dict_of_headers.keys():
    #     if i in dga_headers2:
    #         pass
    #     else:
    #         s += dict_of_headers[i][1]

    # val = s / len(dga_headers)
    #
    # for i in dga_headers2:
    #     dict_of_headers[i][1] += val

    widthValue = []
    for i in range(len(dga_headers2)):
        if dga_headers2[i] in dict_of_headers.keys():
            widthValue.append(dict_of_headers[dga_headers2[i]][1])

    for i in data['DATA']:
        dga_data.extend(i.values())
    for i in range(len(data['DATA'])):
        mergeVaL = ceil(len(data['DATA'][i]['TestId']) / 9)
        if mergeVaLue < mergeVaL:
            mergeVaLue = mergeVaL

    return [dga_headers, dga_data, mergeVaLue, widthValue]


def dataTable(data, value, windingType):
    valueTableKey = ["Measured Resistances", "Time", "0:00", "0:15", "0:30",
                     "0:45", "1:00", "1:15", "1:30", "1:45", "2:00", "2:15", "2:30", "2:45", "3:00",
                     "3:15", "3:30", "3:45", "4:00", "4:15", "4:30", "4:45", "5:00", "5:15", "5:30", "5:45", "6:00",
                     "6:15", "6:30", "6:45", "7:00",
                     "7:15", "7:30", "7:45", "8:00", "8:15", "8:30", "8:45", "9:00", "9:15", "9:30", "9:45", "10:00"]
    tableData = []
    tableDataOne = []
    tableDataTwo = []
    k = 0
    for j in range(len(valueTableKey)):
        if j == 0:
            tableDataOne.append("")
            for i in windingType:
                tableDataOne.append(valueTableKey[j])
        elif j == 1:
            tableDataOne.append(valueTableKey[j])
            for i in windingType:
                tableDataOne.append(i)
        else:
            if j < 23:
                tableDataOne.append(valueTableKey[j])
                for i in windingType:
                    try:
                        if data[i]["calculatedValues"][k]["actualResistance"] == 0:
                            tableDataOne.append(' ')
                        else:
                            tableDataOne.append(round(data[i]["calculatedValues"][k]["actualResistance"], 6))
                    except:
                        tableDataOne.append(None)
            else:
                tableDataTwo.append(valueTableKey[j])
                for i in windingType:
                    try:
                        if data[i]["calculatedValues"][k]["actualResistance"] == 0:
                            tableDataOne.append(' ')
                        else:
                            tableDataTwo.append(round(data[i]["calculatedValues"][k]["actualResistance"], 6))
                    except:
                        tableDataTwo.append(None)
            k += 1
    tableData.append(tableDataOne)
    tableData.append(tableDataTwo)
    return tableData


def tableOneList(data, value, windingType):
    tableOneKey = ["Unit", "MVA", "Cooling", "Winding Terminal", "Cold Resistance", "Cold R Temp °C",
                   "Amb. Temp °C @ SD",
                   "Hot Resistance", "AWR @ Shutdown"]
    tableOneSearch = ["transformerId", "hrMVA", "cooling", "Wdg & Terms", "coldResistance", "coldRTempC",
                      "ambTempShutdown", "hotResistance", "awrShutdown"]
    tableOneData = []
    for j in range(len(tableOneKey)):
        if j in range(3):
            tableOneData.append(tableOneKey[j])
            if j == 2:
                for i in windingType:
                    tableOneData.append(data[windingType[0]]["basicInfo"][tableOneSearch[j]])
            else:
                for i in windingType:
                    tableOneData.append(data[windingType[0]][tableOneSearch[j]])
        elif j == 3:
            tableOneData.append(tableOneKey[j])
            for i in windingType:
                tableOneData.append(i)
        else:
            tableOneData.append(tableOneKey[j])
            for i in windingType:
                try:
                    lenstr = len(str(data[windingType[0]]["basicInfo"][tableOneSearch[4]]).split(".")[1])
                    insertData = round(data[i]["basicInfo"][tableOneSearch[j]], lenstr)
                    if len(str(insertData).split(".")[1]) < lenstr:
                        insertData = str(insertData) + '0' * (lenstr - len(str(insertData).split(".")[1]))
                    tableOneData.append(insertData)
                except:
                    tableOneData.append(None)
    return tableOneData


def firstTestData(data):
    return_data = []
    header_list = ['']
    data_list = []
    # Header Data
    for i in data['Headers']:
        header_list.extend(i.values())

    # Data
    for i in data['Data']:
        data_list.extend(i.values())

    columnWidth = firstTestColumnWidth(len(data['Data'][0]))
    return_data.append(header_list)
    return_data.append(data_list)
    return_data.append(columnWidth)
    return return_data


def EfficeinceTestData(data):
    return_data = []
    tableHeader = []
    tableData = []
    textHeader = []
    textData = []
    textHeight = []
    #     Text header
    if data[0]["istextAvailable"] == 'true':
        textHeader.append(data[0]["textHeader"])
        return_data.append(textHeader)

    #     Text Data
    if data[0]["istextAvailable"] == 'true':
        for i in (data[0]["textValue"].split('\r\n')):
            textData.append(i)
        return_data.append(textData)

    #     Header Data
    for i in data[0]['Table']['header'][0].keys():
        if i in ['1', '2']:
            tableHeader.append(data[0]['Table']['header'][0][i])
        else:
            tableHeader.append(data[0]['Table']['header'][0][i]['1'])

    # text Height.
    for i in textData:
        mergeVaL = ceil(len(i) / 112)
        textHeight.append(mergeVaL)
    return_data.append(textHeight)

    for i in data[0]['Table']['header'][0].keys():
        if i not in ['1', '2']:
            tableHeader.extend(data[0]['Table']['header'][0][i]['2'])

    return_data.append(tableHeader)

    # Data
    for i in data[0]['Table']['Data']:
        tableData.extend(i.values())

    return_data.append(tableData)

    return return_data


def merge_list(data):
    merge_list = []
    for i in range(3, len(data['DATA']['efficience'][0]['Table']['header'][0]) + 1):
        merge_list.append(len(data['DATA']['efficience'][0]['Table']['header'][0][str(i)]['2']))

    return merge_list


def controlTestData(data):
    headerData = []
    tableData = []
    returnData = []
    testKey = []
    for i in range(len(data['headers'])):
        if i == 0:
            headerData.append(data['headers'][i]['testHeder'])
        else:
            testKey = [key for key in data['headers'][i].keys()]
            for j in testKey:
                if j != '1':
                    headerData.append(data['headers'][i][j]['1'])
                else:
                    headerData.append(data['headers'][i][j])

    for i in testKey:
        if i != '1':
            headerData.extend(data['headers'][1][i]['2'])

    for i in range(len(data['data'])):
        for j in data['data'][i].keys():
            tableData.append(data['data'][i][j])

    returnData.append(headerData)
    returnData.append(tableData)
    return returnData

def dewpointTestData(data):
    headerData = []
    tableData = []
    returnData = []

    for j in data['headers']:
        headerData.extend(j.values())
    for j in data['data']:
        tableData.extend(j.values())
    returnData.append(headerData)
    returnData.append(tableData)
    return returnData


def efficenceTestData(data):
    headerData = []
    tableData = []
    returnData = []
    cellWidht = [2.5, 2.5, 2.5]
    for i in range(len(data['headers'])):
        headerData.extend(data['headers'][i].values())
    for i in range(len(data['data'])):
        tableData.extend(data['data'][i].values())
    returnData.append(headerData)
    returnData.append(tableData)
    return returnData


def pdDvParallelTestData(data):
    headerData = []
    tableData = []
    returnData = []
    for i in range(len(data['headers'])):
        if i == 2:
            headerKey = [key for key in data['headers'][i].keys()]
            for j in headerKey:
                if j in ('1', '2', '3', '4'):
                    headerData.append(data['headers'][i][j])
                else:
                    headerData.append(data['headers'][i][j]['1'])
            for j in headerKey:
                if j not in ('1', '2', '3', '4'):
                    headerData.extend(data['headers'][i][j]['2'])
            pass
        else:
            headerData.extend(data['headers'][i].values())
    for i in range(len(data['data'])):
        tableData.extend(data['data'][i].values())
    returnData.append(headerData)
    returnData.append(tableData)
    return returnData


def thermalTestData(data):
    returnData = {"textHeader": [], "textData": [], "textHeight": [], "mergeHeader": [], "mergeHeaderCell": [],
                  "tableHeader": [], "tableData": [], "tableDataMerge": []}
    mergeCount = 0
    #     Text header
    if data["istextAvailable"] == 'true':
        returnData["textHeader"].append(data["textHeader"])
        for i in (data["textValue"].split('\r\n')):
            returnData['textData'].append(i)

    #     Header Data
    for i in range(len(data['Table']['headers'])):
        if i == 0:
            returnData['tableHeader'].extend(data['Table']['headers'][i].values())
        else:
            headerKey = [key for key in data['Table']['headers'][i].keys()]
            for j in headerKey:
                if type(data['Table']['headers'][i][j]) is not str:
                    returnData['mergeHeader'].append(len(data['Table']['headers'][i][j]['2']))
                    returnData['tableHeader'].append(data['Table']['headers'][i][j]['1'])
                    mergeCount += len(data['Table']['headers'][i][j]['2'])
                else:
                    returnData['tableHeader'].append(data['Table']['headers'][i][j])
                    returnData['mergeHeaderCell'].append(mergeCount)
                    mergeCount += 1
            for j in headerKey:
                if type(data['Table']['headers'][i][j]) is not str:
                    returnData['tableHeader'].extend(data['Table']['headers'][i][j]['2'])

    # text Height.
    for i in returnData['textData']:
        mergeVaL = ceil(len(i) / 112)
        returnData['textHeight'].append(mergeVaL)

    # Data

    for i in range(len(data['Table']['data'])):
        mergeCount = 0
        headerKey = [key for key in data['Table']['data'][i].keys()]
        returnData['tableDataMerge'].append({str(i + 1): []})
        for j in headerKey:
            if type(data['Table']['data'][i][j]) is not str:
                returnData['tableData'].extend(data['Table']['data'][i][j])
                mergeCount += 1
            else:
                returnData['tableData'].append(data['Table']['data'][i][j])
                returnData['tableDataMerge'][i][str(i + 1)].append(mergeCount)
                mergeCount += 1
    return returnData


def impulseTestData(data):
    impulse_header = []
    impulse_data = []
    return_data = []

    for i in (data['headers']):
        impulse_header.extend(i.values())

    for i in (data['data']):
        impulse_data.extend(i.values())

    return_data.append(impulse_header)
    return_data.append(impulse_data)

    return return_data


def potencialTestData(data):
    potential_header = []
    potential_data = []
    return_data = []

    for i in (data['headers']):
        potential_header.extend(i.values())

    for i in (data['data']):
        potential_data.extend(i.values())

    return_data.append(potential_header)
    return_data.append(potential_data)

    return return_data


def parallelInducedVoltageTestData(data):
    parallel_header = []
    parallel_data = []
    return_data = []
    for i in (data['headers']):
        parallel_header.extend(i.values())
    for i in (data['data']):
        parallel_data.extend(i.values())
    return_data.append(parallel_header)
    return_data.append(parallel_data)
    return return_data


def seriesInducedVoltageTestData(data):
    series_header = []
    series_data = []
    return_data = []
    for i in (data['headers']):
        series_header.extend(i.values())
    for i in (data['data']):
        series_data.extend(i.values())
    return_data.append(series_header)
    return_data.append(series_data)
    return return_data


def insulationResistanceTableTestData(data):
    insulationResistance_header = []
    insulationResistance_data = []
    return_data = []

    for i in (data['headers']):
        insulationResistance_header.extend(i.values())

    for i in (data['data']):
        insulationResistance_data.extend(i.values())

    return_data.append(insulationResistance_header)
    return_data.append(insulationResistance_data)

    return return_data

def soundTestData(data, textRequired = True):
    returnData = { "mergeHeader":[], "mergeHeaderCell":[], "tableHeader":[], "tableData":[], "tableDataMergeCell":[], "dataRowMerge":[],"numberOfRows":0}
    mergeCount = 0
    rowNumber = 1
    headerRowCount = 0
    for i in range(len(data['headers'])):
        if i < len(data['headers'])-1:
            try:
                key = [key for key in data["headers"][i].keys()]
                for j in (data["headers"][i][key[0]].split('\r\n')):
                    returnData['tableHeader'].append(j)
                    headerRowCount+=1
            except:
                returnData['tableHeader'].extend(data['headers'][i].values())
                headerRowCount+=1
        else:
            headerKey = [key for key in data['headers'][i].keys()]
            for j in headerKey:
                if type(data['headers'][i][j]) is not str:
                    returnData['mergeHeader'].append(len(data['headers'][i][j]['2']))
                    returnData['tableHeader'].append(data['headers'][i][j]['1'])
                    mergeCount += len(data['headers'][i][j]['2'])
                else:
                    returnData['tableHeader'].append(data['headers'][i][j])
                    returnData['mergeHeaderCell'].append(mergeCount)
                    mergeCount += 1
            for j in headerKey:
                if type(data['headers'][i][j]) is not str:
                    returnData['tableHeader'].extend(data['headers'][i][j]['2'])
    # print(hea)
    returnData['numberOfHeaderRow'] = headerRowCount + 2
    for i in range(len(data['data'])):
        mergeCount = 0
        headerKey = [key for key in data['data'][i].keys()]

        checkMerge = False
        checkString = True
        for j in headerKey:
            if type(data['data'][i][j]) is str:
                checkMerge = True
                break
        for j in headerKey:
            if type(data['data'][i][j]) is not str:
                checkString = False
                break
        if checkMerge:
            returnData['numberOfColumns'] = len(headerKey)
            if checkString:
                returnData['tableData'].extend(data['data'][i].values())
                returnData['numberOfRows'] = i + 1
            else:
                returnData['tableDataMergeCell'].append({str(i + 1): []})
                returnData['dataRowMerge'].append({str(rowNumber): []})
                for j in headerKey:
                    if type(data['data'][i][j]) is list:
                        returnData['tableData'].extend(data['data'][i][j])
                        returnData['dataRowMerge'][i][str(rowNumber)].append(len(data['data'][i][j]))
                        mergeCount += 1
                    elif type(data['data'][i][j]) is dict:
                        returnData['tableData'].extend(data['data'][i][j].values())
                        returnData['dataRowMerge'][i][str(rowNumber)].append(len(data['data'][i][j]))
                        mergeCount += 1
                    else:
                        returnData['tableData'].append(data['data'][i][j])
                        returnData['tableDataMergeCell'][i][str(i + 1)].append(mergeCount)
                        mergeCount += 1
                if i < len(data['data']) - 3:
                    rowNumber += max(returnData['dataRowMerge'][i][str(rowNumber)])
                else:
                    returnData['dataRowMerge'][i][str(rowNumber)].append(2)
                    rowNumber += 2
                if i == len(data['headers']) - 1:
                    returnData['numberOfRows'] = rowNumber
        else:
            mergeValue = 0
            a = 0
            for j in headerKey:
                returnData['tableDataMergeCell'].append({str(rowNumber): []})
                returnData['dataRowMerge'].append({str(rowNumber): []})
                returnData['tableData'].append(j)
                returnData['tableDataMergeCell'][mergeValue][str(rowNumber)].append(0)
                for k in data['data'][i][j]:
                    returnData['numberOfColumns'] = 1 + len(k)
                    returnData['tableData'].extend(k.values())
                returnData['dataRowMerge'][mergeValue][str(rowNumber)].append(len(data['data'][i][j]))
                rowNumber += len(data['data'][i][j])
                mergeValue += 1
                a += 1
                if a == len(headerKey):
                    returnData['numberOfRows'] = rowNumber - 1
    returnData['cellTextHeight'] = []
    if textRequired:
        rowNumber = 1
        count= 0
        headerKey = [key for key in data['data'][0].keys()]
        for j in headerKey:
            for k in data['data'][0][j]:
                returnData['cellTextHeight'].append({str(rowNumber):[]})
                for m in k.values():
                    # if (m.replace(' ',"i").isalpha()):
                    mergeVaL = ceil(len(m) /12)
                    returnData['cellTextHeight'][count][str(rowNumber)].append(mergeVaL)
                count+=1
                rowNumber+=1
    return returnData


def coreExcitationStateTestData(data):
    coreExcitation_header = []
    coreExcitation_data = []
    return_data = []

    for i in range(len(data[0]['headers'])):
        if i == 0:
            coreExcitation_header.extend(data[0]['headers'][i].values())
        elif i == 1:
            for i, j in data[0]['headers'][i].items():
                if type(j) == dict:
                    coreExcitation_header.append(j['1'])
                else:
                    coreExcitation_header.append(j)

    for i in range(len(data[0]['headers'])):
        if i == 1:
            for i, j in data[0]['headers'][i].items():
                if type(j) == dict:
                    coreExcitation_header.extend(j['2'])

    for i in (data[0]['data']):
        coreExcitation_data.extend(i.values())

    return_data.append(coreExcitation_header)
    return_data.append(coreExcitation_data)

    return return_data


def mergeListCoreExcitation(data):
    mergeCount = 0
    mergeListVerticalHeader = []
    mergeListHorizontalHeader = []

    for i, j in data[0]['headers'][1].items():

        if type(j) == str:

            mergeListVerticalHeader.append(mergeCount)

            mergeCount = mergeCount + 1

            mergeListHorizontalHeader.append(1)
        else:

            mergeCount = mergeCount + len(j['2'])

            mergeListHorizontalHeader.append(len(j['2']))

    return mergeListHorizontalHeader, mergeListVerticalHeader


def coreExcitationDataInsertMergeList(mergeListHorizontalHeader, numOfColumns):
    lis = []
    k = 0
    for j in mergeListHorizontalHeader:
        lis.append(k)
        for i in range(k, numOfColumns - 1, j):
            break
        k = k + j
    return lis


def cellToInsertCoreExcitation(mergeListVerticalHeader, numOfColumns):
    list_to_insert = []
    for i in range(0, numOfColumns):
        if i not in mergeListVerticalHeader:
            list_to_insert.append(i)
    return list_to_insert


def leadResistanceTableTestData(data):
    insulationResistance_header = []
    insulationResistance_data = []
    return_data = []

    for i in (data['headers']):
        insulationResistance_header.extend(i.values())

    for i in (data['data']):
        insulationResistance_data.extend(i.values())

    return_data.append(insulationResistance_header)
    return_data.append(insulationResistance_data)

    return return_data


def zeroSequenceTableTestData(data):
    zeroSequence_header = []
    zeroSequence_data = []
    return_data = []
    for i in range(len(data[0]['headers'])):
        if i == 0:
            zeroSequence_header.extend(data[0]['headers'][i].values())
        elif i == 1:
            for i, j in data[0]['headers'][i].items():
                if type(j) == dict:
                    zeroSequence_header.append(j['1'])
                else:
                    zeroSequence_header.append(j)

    for i in range(len(data[0]['headers'])):
        if i == 1:
            for i, j in data[0]['headers'][i].items():
                if type(j) == dict:
                    zeroSequence_header.extend(j['2'])
    for i in (data[0]['data']):
        zeroSequence_data.extend(i.values())

    return_data.append(zeroSequence_header)
    return_data.append(zeroSequence_data)

    return return_data
