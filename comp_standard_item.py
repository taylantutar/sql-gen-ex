import io
from openpyxl import load_workbook

async def ItemSqlPrepear(file, id):
    try:
        contents = await file.read()
        workbook = load_workbook(io.BytesIO(contents), read_only=True)
        sheet = workbook.active
        data = []
        for i, row in enumerate(sheet.iter_rows()):
            if (i == 0):
                continue
            if (row[0].value == None):
                break
            # row_data = [cell.value for cell in row]
            data.append(f"INSERT INTO [dbo].[TblCompStandardItems]([CompItemID],[StandardCode],[ItemCode],[AreaName],[CsiosID],[InsertDate],[Description])VALUES({id}, '{row[0].value}', '{row[2].value}', '{row[1].value}',1, GETUTCDATE(), '{row[3].value}');")
        return {"data": data}

    except Exception as e:
        return {"error": str(e)}
