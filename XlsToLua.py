import xlrd
import sys
import os
import re
import datetime
import time

today = datetime.date.today
now = time.strftime("%H:%M:%S")
dirName = time.strftime("%Y%m%d-%H%M%S", time.localtime())

print('dn: ' + dirName)
os.makedirs('./' + dirName)

fileError = False


dataStartRow = 4
dataNameRow = 1
dataTypeRow = 2

targetSheetName = 'data'
indent = 0

def line(strs):
    return strs + '\n'

writeData = line('---@')

annotaionData = line('')
annotaionData = line('--- 注释')

luaRequirefileName = 'LuaDataRequire.lua.txt'
luaRequireData = line('--- require')

def AddIdx(i):
    global writeData
    ShowIndent()
    writeData += '[{idx}]'.format(idx = i) + ' = '

def AddKey(key):
    global writeData
    writeData += '{key} = '.format(key = key)

def AddValue(value, t = 0):

    global writeData
    if t == 0:
        writeData += '{v},\n'.format(v = value)
    elif t == 2:
        if value == 'TRUE' or 1 or 'true' or 'True':    
            value = 'true'
        elif value == 'FALSE' or 0 or 'false' or 'Flase':
            value = 'false'
        writeData += '{v},\n'.format(v = value)
    else:
        writeData += "'{v}',\n".format(v = value)



def ShowIndent():

    global indent
    global writeData
    for _ in range(indent):
        writeData += '    '

def AddIndent(n = 0):

    global indent
    indent = indent + n

def SheetParseIdx(booksheet, row):

    global writeData
    global indent

    #print('idx parse:{row}'.format(row = row))
    dataType = str(booksheet.cell(dataTypeRow, 0).value)

    ShowIndent()
    if dataType == 'string':
        return True
    elif dataType == 'int':

        writeData += '[{idx}] = '.format(idx = int(booksheet.cell(row, 0).value))
        writeData += '{\n'
        return True
    else:
        print('ERROR: 第一列的类型非法，必须是int或者string类型..\n')
        return False

def SheetParseCol(booksheet, row, col):

    global writeData
    global annotaionData

    try:

        key = str(booksheet.cell(dataNameRow, col).value)

        if '=HL' in key:
            return True

        d = booksheet.cell(row, col).value

        if d == None or d == '':
            print(line(''))
            print('找到一个空的表格！！: {row},{col}'.format(row = row, col = col))
            print(line(''))
            return True

        valueType = str(booksheet.cell(dataTypeRow, col).value)
        ShowIndent()
        AddKey(key)

        addcontent = False
        if row == 4:
            addcontent = True
            
        if valueType == 'int':
            print('找到一个int: {row},{col}'.format(row = row, col = col))
            value = int(booksheet.cell(row, col).value)
            AddValue(value)

            if addcontent:
                annotaionData += line('---@field ' + key + ' number')

        elif valueType == 'string':
            print('找到一个string: {row},{col}'.format(row = row, col = col))
            value = str(booksheet.cell(row, col).value)
            AddValue(value, 1)

            if addcontent:
                annotaionData += line('---@field ' + key + ' string')

        elif valueType == 'float':
            print('找到一个float: {row},{col}'.format(row = row, col = col))
            value = float(booksheet.cell(row, col).value)
            AddValue(value)

            if addcontent:
                annotaionData += line('---@field ' + key + ' number')

        elif valueType == 'bool':
            print('找到一个bool: {row},{col}'.format(row = row, col = col))
            value = booksheet.cell(row, col).value

            AddValue(value, 2)

            if addcontent:
                annotaionData += line('---@field ' + key + ' boolean')
        else:
            m = re.match('array', valueType)
            if m != None:
                print('找到一个Array: {row},{col}..'.format(row = row, col = col))
                obj = re.search(r'\[.+\]', valueType)
                s = obj.group()
                arrayType = s[1:len(s)-1]
                value = booksheet.cell(row, col).value.split(',')
                
                t = 0
                if arrayType == 'string':
                    t = 1
                    if addcontent:                    
                        annotaionData += line('---@field ' + key + ' string[]')
                elif arrayType == 'bool':
                    t = 2
                    if addcontent:
                        annotaionData += line('---@field ' + key + ' boolean[]')
                else:
                    if addcontent:
                        annotaionData += line('---@field ' + key + ' number[]')

                writeData += '{\n'
                AddIndent(1)
                for idx in range(len(value)):
                    AddIdx(idx + 1)
                    AddValue(value[idx], t)
                AddIndent(-1)
                ShowIndent()
                writeData += '},\n'
                
            m = re.match('table', valueType)
            if m != None:
                print('找到一个Table: {row},{col}..'.format(row = row, col = col))
                writeData += '{\n'
                AddIndent(1)
                ShowIndent()
                value = booksheet.cell(row, col).value
                AddValue(value)
                AddIndent(-1)
                ShowIndent()
                writeData += '},\n'
                
                if addcontent:
                    annotaionData += line('---@field ' + key + ' table<string, number>')

    except ValueError:
        print("############# O_o||| ==> 策划兄弟类型出错啦: {row},{col}..".format(row = row, col = col))
        global fileError
        fileError = True
        return False
            

def SheetParse(fileName, booksheet):
    
    global writeData
    global annotaionData
    global fileError
    global luaRequireData

    fileError = False
    annotaionData = ''
    writeData = ''


    
    for col in range(booksheet.ncols):
        writeData += '--- '
        writeData += str(booksheet.cell(0,col).value) + '\t\t'
        writeData += str(booksheet.cell(1,col).value) + '\t\t'
        writeData += str(booksheet.cell(2,col).value) + '\n'
    
    writeData += line('')
    writeData += line(fileName  + "Data  = {}")
    writeData += line(fileName  + "Data.d = {")
    annotaionData += line('---@class ' + fileName  + "Data")
    annotaionData += line('---@field d ' + fileName + "DataIns" )
    annotaionData += line('')
    annotaionData += line('---@class ' + fileName + "DataIns" )

    AddIndent(1)

    for row in range(dataStartRow, booksheet.nrows):
        SheetParseIdx(booksheet, row)
        AddIndent(1)
        for col in range(1, booksheet.ncols):
            SheetParseCol(booksheet, row, col)
        AddIndent(-1)
        ShowIndent()
        writeData += '},\n'
        

    writeData += '}\n'
    writeData += annotaionData

    if fileError == False:    
        fileOutput = open('./' + dirName + '/' + fileName + 'Data.lua.txt', 'w')
        fileOutput.write(writeData)
        luaRequireData += line("require('LuaData." + fileName + "Data')")
    else:
        fileOutput = open('./' + dirName + '/' + fileName + 'Data.出错啦兄弟', 'w')

    return not fileError




for parent,dirnames,filenames in os.walk("."):

    #print('parent = {parent}, dirnames = {dirnames}, filenames = {filenames}'.format(parent = parent, dirnames = dirnames, filenames = filenames))

    for filename in filenames:

        portion = os.path.splitext(filename)
        
        ext = portion[1]

        if ext == '.xlsx':
            
            finddatasheet = False

            print ('找到文件 -> {target} <- ..'.format(target = filename))
            
            workbook = xlrd.open_workbook(filename)

            for booksheet in workbook.sheets():

                if booksheet.name == targetSheetName: 
                    print ('找到 -> {target} <- sheet ..'.format(target = targetSheetName))
                    finddatasheet = True
                    res = SheetParse(portion[0], booksheet)

                    if res == False:
                        print ('当前 xls 生成失败.\n')
                        break

            else:
                if finddatasheet:
                    print ('当前 xls 生成结束.\n')
                    fileOutput = open('./' + dirName + '/' + luaRequirefileName, 'w')
                    fileOutput.write(luaRequireData)
                else:
                    print ('没有找到 --> {target} <- sheet！生成失败..\n'.format(target = targetSheetName))



