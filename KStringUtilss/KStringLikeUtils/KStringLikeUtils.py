# -*- coding: utf-8 -*-
import xlrd as xlrd
import xlwt as xlwt
from fuzzywuzzy import fuzz
from fuzzywuzzy import process



def readFile(fileName):
    data = xlrd.open_workbook('企业级客户信息系统源系统码值与标准码值转换映射20201204.xlsx')
    table = data.sheet_by_index(0)
    tablest = data.sheet_by_index(1)
    for nrow in range(table.nrows):
        if table.cell(nrow,0).value == 'OCCUPATION_CODE':
            srcCode = str(table.cell(nrow,5).value)
            srcName = str(table.cell(nrow,6).value)
            midRation = 0
            resRation = 0
            stCode = ''
            stName = ''
            for stnrow in range(tablest.nrows):
                midCode = str(tablest.cell(stnrow,2).value)
                #midName1 = str(tablest.cell(stnrow,0).value)+str(tablest.cell(stnrow,1).value)+str(tablest.cell(stnrow,3).value)
                #midName = str(tablest.cell(stnrow,3).value)
                midName = str(tablest..
                #midRation = fuzz.partial_ratio(midName,srcName)
                if resRation <= midRation:
                    resRation = midRation
                    stCode = midCode
                    stName = midName
            print(srcCode,srcName,resRation,stCode,stName)

if __name__ == '__main__':
    #resration  = fuzz.ratio("this is a test", "this is a test!")
    #print(resration)
    readFile('')
    #resration = fuzz.token_sort_ratio("防毒器材装配工", "防化\防毒器材制造\装配工人")
    #print(resration)

