import openpyxl


def ItemSqlPrepear(file_name, id_start_index):
    try:
        file = openpyxl.load_workbook(file_name)
        sheet = file.active

        sqlList = []
        for i, r in enumerate(sheet.iter_rows()):
            if (i == 0):
                continue
            if (r[0].value == None):
                break
            sqlTxt = ""
            sqlTxt = f"INSERT INTO [dbo].[TblCompStandardItems]([CompItemID],[StandardCode],[ItemCode],[AreaName],[CsiosID],[InsertDate],[Description])VALUES({id_start_index}, '{r[0].value}', '{r[2].value}', '{r[1].value}',1, GETUTCDATE(), '{r[3].value}');"
            sqlList.append(sqlTxt)
            id_start_index += 1
        return sqlList

    except FileNotFoundError:
        print(f"Hata: {file_name} dosyası bulunamadı.")
    except Exception as e:
        print(f"Bir hata oluştu: {e}")
