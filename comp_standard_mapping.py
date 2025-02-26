import io
from openpyxl import load_workbook


async def MappingDataPrepear(file):
    try:
        contents = await file.read()
        workbook = load_workbook(io.BytesIO(contents), read_only=True)
        sheet = workbook.active

        mappingList = []
        for i, row in enumerate(sheet.iter_rows()):
            if (i == 0):
                continue
            if (row[0].value == None):
                break
            mappingDict = {}
            mappingDict["SrcStandardCode"] = row[0].value.strip()
            mappingDict["SrcItemCode"] = row[1].value.strip()
            mappingDict["DestStandardCode"] = row[2].value.strip()
            mappingDict["DestItemCode"] = row[3].value.strip()
            mappingDict["BertMeanStr"] = row[4].value
            mappingDict["BertMean"] = int(round(float(row[4].value) * 100, 0))
            mappingList.append(mappingDict)

        return mappingList

    except FileNotFoundError:
        raise Exception(f"Error: {file.filename} file is not found.")
    except Exception as e:
        raise Exception(f"An error occured: {e}")


def CheckValidations(mappingList, sFrmStr, iFrmStr):
    try:
        unvalidData = []
        for mapping in mappingList:
            if mapping["SrcStandardCode"] in iFrmStr and mapping["DestStandardCode"] in sFrmStr:
                unvalidData.append(f"Unvalid mapping. {mapping["SrcStandardCode"]} --> {mapping["DestStandardCode"]}")

        if len(unvalidData) > 0:
            raise Exception(unvalidData)

    except Exception as e:
        raise e


def PrepareSql(mappingList):
    sqlList = []
    for mapping in mappingList:
        sqlTxt = f"INSERT INTO TblCompStandardItemMapping([SrcStandardCode],[SrcItemCode],[DestStandardCode],[DestItemCode],[Relevence],[Confidence],[AreaScore],[ItemScore],[InsertDate],[Status])VALUES('{mapping["SrcStandardCode"]}', '{mapping["SrcItemCode"]}', '{mapping["DestStandardCode"]}', '{mapping["DestItemCode"]}', {mapping["BertMean"]} , {mapping["BertMean"]}, {mapping["BertMean"]}, {mapping["BertMean"]}, GETUTCDATE(), 1);"
        sqlList.append(sqlTxt)
    return sqlList
