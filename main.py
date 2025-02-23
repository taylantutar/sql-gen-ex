from fastapi import FastAPI
from pydantic import BaseModel

import comp_standard_item
import comp_standard_mapping

app = FastAPI()

from typing import List

class RequestDataItem(BaseModel):
    standards: List[str]
    id_start: int

@app.post("/CompliancesCreateItems/")
async def create_item(data: RequestDataItem):
    sql = comp_standard_item.ItemSqlPrepear("1-ControlItems.xlsx", data.id_start)
    return {"result": sql}

class RequestDataMapping(BaseModel):
    standards: List[str]

@app.post("/CompliancesCreateMappings/")
async def create_mapping(data: RequestDataMapping):    
    # print(data.standards)
    mappingList = comp_standard_mapping.MappingDataPrepear("2-Mappings.xlsx")

    hasError,result = comp_standard_mapping.CheckValidations(mappingList=mappingList, standardList=data.standards)
    if not hasError :
       successResult = comp_standard_mapping.PrepareSql(mappingList,data.standards)
       return {"result":successResult}
    else:
       return {"result":result}

        