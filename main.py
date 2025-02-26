import uvicorn
from fastapi import FastAPI, File, Form, UploadFile
from openpyxl import load_workbook
import io

import comp_standard_item
import comp_standard_mapping

app = FastAPI()


@app.post("/sayhello/")
async def create_item(
):
    print("Hello")
    return {
        "result": "Hello"
    }


@app.post("/items/")
async def create_item(new_standard_name: str = Form(...), id: int = Form(...), file: UploadFile = File(...)):
    try:
        print(new_standard_name)
        sqlList = await comp_standard_item.ItemSqlPrepear(file, id, new_standard_name)

        return {
            "Count": len(sqlList),
            "Data": sqlList
        }

    except Exception as e:
        return {"Error": str(e)}


@app.post("/mappings/")
async def create_mapping(sFrmStr: str = Form(...), iFrmStr: str = Form(...), file: UploadFile = File(...)):
    try:
        mappingList = await comp_standard_mapping.MappingDataPrepear(file)

        comp_standard_mapping.CheckValidations(mappingList, sFrmStr, iFrmStr)

        # sqlList = comp_standard_mapping.PrepareSql(mappingList)

        # return {
        #     "Count": len(sqlList),
        #     "Data": sqlList
        # }

    except Exception as e:
        return {"Error": str(e)}


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8001)
