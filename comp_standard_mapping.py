import openpyxl


def MappingDataPrepear(file_name):
    try:
        file = openpyxl.load_workbook(file_name)
        sheet = file.active

        mappingList = []
        for i, r in enumerate(sheet.iter_rows()):
            if (i == 0):
                continue
            if (r[0].value == None):
                break
            mappingDict = {}
            mappingDict["SrcStandardCode"] = r[0].value
            mappingDict["SrcItemCode"] = r[1].value
            mappingDict["DestStandardCode"] = r[2].value
            mappingDict["DestItemCode"] = r[3].value
            mappingDict["BertMeanStr"] = r[4].value
            mappingDict["BertMean"] = round(float(r[4].value) * 100, 0) 
            # for c in r:
            #     if (c.value == None):
            #         break
            #     print(f"{c.value};", end="", flush=True)
            # print()
            # sqlTxt = f"INSERT INTO [dbo].[TblCompStandardItems]([CompItemID],[StandardCode],[ItemCode],[AreaName],[CsiosID],[InsertDate],[Description])VALUES({id_start_index}, '{r[0].value}', '{r[2].value}', '{r[1].value}',1, GETUTCDATE(), '{r[3].value}');"
            # sqlList.append(sqlTxt)
            mappingList.append(mappingDict)
        return mappingList

    except FileNotFoundError:
        print(f"Error: {file_name} file is not found.")
    except Exception as e:
        print(f"An error occured: {e}")

def PrepareSql(mappingList, standardList):
    sqlList = []
    for mapping in mappingList:
        sqlTxt = f"INSERT INTO TblCompStandardItemMapping([SrcStandardCode],[SrcItemCode],[DestStandardCode],[DestItemCode],[Relevence],[Confidence],[AreaScore],[ItemScore],[InsertDate],[Status])VALUES('{mapping["SrcStandardCode"]}','{mapping["SrcItemCode"]}','{mapping["DestStandardCode"]}','{mapping["DestItemCode"]}',{mapping["BertMean"]} ,{mapping["BertMean"]},{mapping["BertMean"]},{mapping["BertMean"]},GETUTCDATE(),1);"
        sqlList.append(mapping)



