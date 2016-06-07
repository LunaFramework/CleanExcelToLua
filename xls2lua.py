# -*- coding: utf-8 -*- 
import  xdrlib ,sys
import xlrd
import os
from slpp import slpp as lua
reload(sys)
sys.setdefaultencoding('utf-8') 

replace_sheet_name = "datafield_replace"

def open_excel(filename):
    try:
        data = xlrd.open_workbook(filename)
        return data
    except Exception,e:
        print str(e)

def excel_table_byindex(data,colnameindex=0,by_index=0,rdict=None,sheetname=""):
    table = data.sheets()[by_index]
    nrows = table.nrows
    ncols = table.ncols
    colnames =  table.row_values(colnameindex)
    tablelist =[]
    # foreach row concat in one list
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             # every row's key-value data
             app = {}
             for i in range(len(colnames)):
                 # replace value,so you can use enum in excel
                 key = replace(colnames[i],rdict,sheetname)
                 value = replace(row[i],rdict,sheetname)
                 # cell type == boolean
                 if table.cell_type(rownum, i) == 4:
                     value = value == 1
                 # make same key as one array                 
                 if key in app:
                     listvalue = []
                     if isinstance(app[key], list):
                         listvalue = app[key]
                     else:
                         listvalue.append(app[key])

                     listvalue.append(value)
                     app[key] = listvalue
                 else:
                 # solo value
                     app[key] = value
             tablelist.append(app)
    return tablelist

def excel_table_byname(data,colnameindex=0,by_name=u'Sheet1'):
    try:
        table = data.sheet_by_name(by_name)
        nrows = table.nrows
        colnames =  table.row_values(colnameindex)
        list =[]
        for rownum in range(1,nrows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                    list.append(app)
        return list
    except :
        return None

def convert_table(filename):    
    data = open_excel(filename)
    dataset = {}
    rdict = excel_table_byname(data, 0, replace_sheet_name)

    for i in range(len(data.sheets())):
        td = data.sheets()[i]
        table = excel_table_byindex(data, 0, i, rdict, td.name)
        dataset[td.name] = table

    return dataset

def replace(value, rdict, sheetname):
    if rdict != None:
        for row in range(len(rdict)):
            rows = rdict[row]
#            print(rows["sheet"], sheetname, rows["sheet"] == sheetname)
            if rows["sheet"] == sheetname:
                for key in rows:
                    if rows["name"] == value:
                        return rows["replace"]
    return value

def convert_lua(dataset, path):
    for name in dataset:
        if name != replace_sheet_name:
            create_lua_file(path, name, dataset[name])

def create_lua_file(path, name, table):
    # replace string to number
    fileobj = open(path + name + ".lua", "w")
    buff = lua.encode(table)
    fileobj.write(buff)
    fileobj.close()    

# feature
# + convert excel data to lua on any os
# + intuitive design, data rule define inside, auto config what data look like(bool first,number second,or string)
# + human friendly text replacement, and auto merge same field into array
# + high performance and easy to use
# - TODO:check data safe,(data dependency/data range)

if __name__ == "__main__":
    print("Useage: xls2lua excel_file_name")
    dataset = convert_table(sys.argv[1])
    convert_lua(dataset, "./")
    print("Success")

