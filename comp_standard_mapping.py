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
            mappingDict["SrcStandardCode"] = row[0].value
            mappingDict["SrcItemCode"] = row[1].value
            mappingDict["DestStandardCode"] = row[2].value
            mappingDict["DestItemCode"] = row[3].value
            mappingDict["BertMeanStr"] = row[4].value
            mappingDict["BertMean"] = round(float(row[4].value) * 100, 0) 
            mappingList.append(mappingDict)

        return mappingList

    except FileNotFoundError:
        print(f"Error: {file.filename} file is not found.")
    except Exception as e:
        print(f"An error occured: {e}")

def PrepareSql(mappingList, standards):
    sqlList = []
    for mapping in mappingList:
        sqlTxt = f"INSERT INTO TblCompStandardItemMapping([SrcStandardCode],[SrcItemCode],[DestStandardCode],[DestItemCode],[Relevence],[Confidence],[AreaScore],[ItemScore],[InsertDate],[Status])VALUES('{mapping["SrcStandardCode"]}','{mapping["SrcItemCode"]}','{mapping["DestStandardCode"]}','{mapping["DestItemCode"]}',{mapping["BertMean"]} ,{mapping["BertMean"]},{mapping["BertMean"]},{mapping["BertMean"]},GETUTCDATE(),1);"
        sqlList.append(sqlTxt)
    return sqlList

def CheckValidations(mappingList, standards):
    return (False,"An Error ocurred.")




