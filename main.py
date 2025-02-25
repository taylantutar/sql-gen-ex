
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
async def create_item(
    new_standard_name: str= Form(...),
    id: int= Form(...),
    file: UploadFile = File(...)
):
    try:
        print(new_standard_name)
        sqlList = await comp_standard_item.ItemSqlPrepear(file, id, new_standard_name)

        return {
            "Count": len(sqlList),
            "Data" : sqlList
            }
    
    except Exception as e:
        return {"error": str(e)}
    

@app.post("/mappings/")
async def create_mapping(
                        standards: str= Form(...),
                         file: UploadFile = File(...)
                         ):
    mappingList = await comp_standard_mapping.MappingDataPrepear(file)

    hasError,result = comp_standard_mapping.CheckValidations(mappingList=mappingList, standards = standards)
    if not hasError :
       successResult = comp_standard_mapping.PrepareSql(mappingList,standards)
       return {"result":successResult}
    else:
       return {"result":result}


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8001)
