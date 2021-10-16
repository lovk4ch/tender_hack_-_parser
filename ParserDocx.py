import os
import glob
import json
from docx import Document
import EnumerablesFiles
from py_linq import Enumerable

names = []
with open('Name.json') as f:
    names = json.load(f)

values = names[1]
items = names[2]
names = names[0]

"""
def findProcentResult(path):
    result = []
    for order in EnumerablesFiles.EnumerableOrder(path):
        print("order: " + order)
        result.append(0)
        for docx in EnumerablesFiles.EnumerableDoc(path + order):
            result[-1]+=1
    print("empity result = " + str(Enumerable(result).where(lambda count: count == 0).count()))
    print("count = " +  str(Enumerable(result).count()))
    return
"""

def Main2(path):
    for order in EnumerablesFiles.EnumerableOrder(path):
        print("order: " + order)
        arr = []
        for docx in EnumerablesFiles.EnumerableDocx(path + order):
            result , data = BuildData(docx)
            if result is True:
                arr.append(data)
                info = []
                arr.append(info)
                for name in data:
                    info.append(isColumnType(name,names))
        SaveJson(path + order + ".json",arr)
    return





def SaveJson(path,data):
    with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    return


def BuildData(path):
    print("Read ->" + path)
    try:
        Document(path)
    except Exception as e:
            print(e)
            return False ,None
    result = []
    for t in Document(path).tables:
        tab = [];
        result.append(tab)
        try:
            sumColum = []
            for i in range(0,t._column_count):
                if t.rows[0] is None:
                    continue;
                if t._cells is None:
                    continue;
                if parsed(t.rows[0].cells[i].text):
                    sumColum.append(i)
            if len(sumColum) > 2:
                for row in t.rows:
                    r = []
                    tab.append(r)
                    for colum in sumColum:
                        ret = ""
                        try:
                            ret = row.cells[colum].text
                        except Exception as e:
                            print(e)
                        r.append(ret);

        except Exception as e:
             print(e)
    return True, result
 

def parsed(str):
    
    info = []
    info.append(isColumnType(str,names))
    info.append(isColumnType(str,values))
    info.append(isColumnType(str,items))
    return Enumerable(info).max() > 0

def writeRow():
    return

def isColumnType(name,indificators):
    count = 0.0
    for indificator in indificators:
        if indificator in name:
            count += 1.0
    return count/len(indificators) #, count/name.split()


Main2("data/");
#findProcentResult("data/");