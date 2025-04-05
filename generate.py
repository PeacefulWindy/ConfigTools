import os
import openpyxl
import json
from slpp import slpp as lua
import xml.etree.ElementTree as ET
import time
import shutil

def export_to_lua(sheet_name, data, output_dir):
    output_file = os.path.join(output_dir, f"{sheet_name}.lua")
    with open(output_file, 'w', encoding='utf-8') as f:
        lua_str = lua.encode(data)
        f.write("return "+lua_str)

def export_to_json(sheet_name, data, output_dir):
    output_file = os.path.join(output_dir, f"{sheet_name}.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=None,separators=(",", ":"))

def export_dict_to_xml(parent, data):
    if isinstance(data, dict):
        for key, value in data.items():
            if isinstance(value, dict):
                sub_element = ET.SubElement(parent, str(key))
                export_dict_to_xml(sub_element, value)
            else:
                sub_element = ET.SubElement(parent, str(key))
                sub_element.text = str(value)

def export_to_xml(sheet_name, data, output_dir):
    output_file = os.path.join(output_dir, f"{sheet_name}.xml")
    root = ET.Element("xml")
    for key, value in data.items():
        export_dict_to_xml(root, {key: value})
    tree = ET.ElementTree(root)
    tree.write(output_file, encoding='utf-8')

formatFunc={
    "lua":export_to_lua,
    "json":export_to_json,
    "xml":export_to_xml,
}

def process_excel(filepath,outputPath):
    wb = openpyxl.load_workbook(filepath,data_only=True)

    sheetNames=[]
    for sheet in wb.worksheets:
        if not sheet.title.startswith("#"):
            continue
        
        sheet_name = sheet.title.strip("#")  # 去掉表名中的第一个特殊符号
        print("sheetName:"+sheet_name)

        output_client_base_dir = os.path.join(outputPath, "client")
        output_server_base_dir = os.path.join(outputPath, "server")
        for ext,func in formatFunc.items():
            output_client_dir = os.path.join(output_client_base_dir, ext)
            if not os.path.exists(output_client_dir):
                os.makedirs(output_client_dir)
            
            output_server_dir = os.path.join(output_server_base_dir, ext)
            if not os.path.exists(output_server_dir):
                os.makedirs(output_server_dir)
            
            client_data_dict = {}
            server_data_dict = {}
            
            for row in range(4, sheet.max_row + 1):
                first_col_value = sheet.cell(row=row, column=1).value
                if not first_col_value:
                    continue
                
                client_row_data = {}
                server_row_data = {}
                for col in range(1, sheet.max_column + 1):
                    col_key = sheet.cell(row=1, column=col).value
                    col_type = sheet.cell(row=2,column=col).value
                    col_value = sheet.cell(row=row, column=col).value

                    if col_type == "json":
                        if col_value:
                            col_value=json.loads(col_value)
                        else:
                            col_value={}
                    elif col_type == "int":
                        if not col_value:
                            col_value=0
                    elif col_type == "float":
                        if not col_value:
                            col_value=0.0
                    elif col_type == "string":
                        if not col_value:
                            col_value=""
                    elif col_type == "bool":
                        if not col_value:
                            col_value=False
                    
                    if col_key and not col_key.startswith("#"):
                        if not col_key.startswith("!"):
                            key=col_key
                            if key.startswith("$"):
                                key=col_key[1:]
                            client_row_data[key] = col_value
                        
                        if not col_key.startswith("$"):
                            key=col_key
                            if key.startswith("!"):
                                key=col_key[1:]
                            server_row_data[key] = col_value
                
                client_data_dict[first_col_value] = client_row_data
                server_data_dict[first_col_value] = server_row_data
            
            start_time = time.time()
            func(sheet_name, client_data_dict,output_client_dir)
            func(sheet_name, server_data_dict,output_server_dir)
            end_time = time.time()
            execution_time = end_time - start_time
            print(ext + " generate time:"+str(execution_time))
        sheetNames.append(sheet_name)
    print("")

    return sheetNames

def move(json_data,type):
    moveRoot=json_data["move"]
    if not type in moveRoot:
        return
    
    if "client" in moveRoot:
        outputClientPath=os.path.join(json_data["output"],type)
        for format in moveRoot[type]:
            outputClientFormatPath=os.path.join(outputClientPath,format)
            if not os.path.exists(outputClientFormatPath):
                os.makedirs(outputClientFormatPath)
            
            for it in moveRoot[type][format]:
                if not os.path.exists(it):
                    os.makedirs(it)
                
                shutil.copytree(outputClientFormatPath, it,dirs_exist_ok=True)

def main():
    json_data={}
    with open("../config.json", "r", encoding="utf-8") as file:
        json_data = json.load(file)

    if not os.path.exists(json_data["output"]):
        os.makedirs(json_data["output"])

    sheetNames=[]
    for root, dirs, files in os.walk(json_data["input"]):
        for file in files:
            if not file.startswith("~") and file.endswith(".xlsx"):
                filepath = os.path.join(root, file)
                print(filepath)
                ret=process_excel(filepath,json_data["output"])
                sheetNames.extend(ret)
    
    for ext,func in formatFunc.items():
        output_client_base_dir = os.path.join(json_data["output"], "client",ext)
        output_server_base_dir = os.path.join(json_data["output"], "server",ext)
        data={}
        
        if os.path.exists(output_client_base_dir):
            for it in sheetNames:
                if os.path.exists(os.path.join(output_client_base_dir,it+"."+ext)):
                    data[it]=True
            func("init",data,output_client_base_dir)
            data.clear()
        
        if os.path.exists(output_server_base_dir):
            for it in sheetNames:
                if os.path.exists(os.path.join(output_server_base_dir,it+"."+ext)):
                    data[it]=True
            func("init",data,output_server_base_dir)

    if not "move" in json_data:
        return
    
    print("will move config...")
    print("move client config...")
    move(json_data,"client")
    print("move server config...")
    move(json_data,"server")
    print("move config finish!")
    print()


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    execution_time = end_time - start_time
    print("total time:"+str(execution_time))
    print("Done!")