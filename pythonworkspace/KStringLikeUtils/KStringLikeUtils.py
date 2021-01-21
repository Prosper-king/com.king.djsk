# -*- coding: utf-8 -*-
import xlrd as xlrd
import xlwt as xlwt
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import  jieba
import pandas


def readFile(fileName):
    data = xlrd.open_workbook('企业级客户信息系统源系统码值与标准码值转换映射20201204.xlsx')
    table = data.sheet_by_index(0)
    tablest = data.sheet_by_index(1)
    for nrow in range(table.nrows):
        if table.cell(nrow,0).value == 'OCCUPATION_CODE' and table.cell(nrow,6).value == '火炮随动系统装试工':
            srcCode = str(table.cell(nrow,5).value)
            srcName = str(table.cell(nrow,6).value)
            midRation = 0
            resRation = 0
            stCode = ''
            stName = ''
            count = 0
            #midName = tablest.col(3)
            for stnrow in range(tablest.nrows):
                midRation = 0
                count = count + 1
                srcName0 = jieba.cut_for_search(srcName)
                srcName1 = jieba.cut_for_search(srcName)
                srcName3 = jieba.cut_for_search(srcName)
                midCode = str(tablest.cell(stnrow,2).value)
                midName = str(tablest.cell(stnrow,3).value)
                midName0 = str(tablest.cell(stnrow,0).value)
                midName1 = str(tablest.cell(stnrow,1).value)
                midName3 = str(tablest.cell(stnrow,3).value)
                midName00 = jieba.cut_for_search(midName0)
                for i in midName00:
                    for j in srcName0:
                        #print("分类1",i,j,fuzz.ratio(i,j))
                        midRation = fuzz.token_sort_ratio(i,j)+midRation
                midName11 = jieba.cut_for_search(midName1)
                for i in midName11:
                    for j in srcName1:
                        #print("分类2",i,j,fuzz.ratio(i,j))
                        midRation = fuzz.token_sort_ratio(i,j)+midRation
                midName33 = jieba.cut_for_search(midName3)
                for i in midName33:
                    for j in srcName3:
                        #print("分类3",i,j,fuzz.ratio(i,j))
                        midRation = fuzz.token_sort_ratio(i,j)+midRation
                #print(midName)
                #print(srcName,midName0,midName1,midName3)
                #midRation0 = fuzz.token_sort_ratio(srcName,midName0)
                #midRation1 = fuzz.token_sort_ratio(srcName,midName1)
                #midRation3 = fuzz.token_sort_ratio(srcName,midName3)
                #print(srcName,midRation0,midRation1,midRation3)
                #midRation = midRation0+midRation1+midRation3
                #print(midRation)
                if resRation <= midRation:
                    resRation = midRation
                    stCode = midCode
                    stName = midName
            print(srcCode,srcName,resRation,stCode,stName,count)
'''
  pandas读写xlsx
'''
def readFileByPandas(fileName):
    table = pandas.read_excel('企业级客户信息系统源系统码值与标准码值转换映射20201227.xlsx',sheet_name=0)
    tablest = pandas.read_excel('企业级客户信息系统源系统码值与标准码值转换映射20201227.xlsx',sheet_name=1)
    c_seq_nolist = table.index
    c_seq_nolist_st = tablest.index
    table['标准码值']=None
    table['标准码中文']=None
    table['概率']=None
    print(c_seq_nolist_st)
    for c_seq_no in c_seq_nolist:
        c_src_nme = str(table.iloc[c_seq_no,6])
        c_src_nme1 = str(table.iloc[c_seq_no,6])
        print(c_src_nme)
        resRation = 0
        st_seq_no = 0
        c_mid_cde = ''
        c_mid_nme = ''
        #c_src_cde =  table[['序号','C_SOURCE_CODE']][table['序号'] == c_seq_no]
        for c_seq_no_st in c_seq_nolist_st:
            midRation = 0
            c_st_nme = str(tablest.iloc[c_seq_no_st,3])
            c_st_cde = str(tablest.iloc[c_seq_no_st,2])
            c_st_class = str(tablest.iloc[c_seq_no_st,0])
            c_st_clazz = str(tablest.iloc[c_seq_no_st,1])
            c_st_nme_cut = jieba.cut_for_search(c_st_nme)
            c_src_nme_cut = jieba.cut_for_search(c_src_nme)
            midRation= fuzz.token_set_ratio(c_src_nme, c_st_nme)
            # for i in c_src_nme_cut:
            #     for j in c_st_nme_cut:
            #         midRation = fuzz.token_sort_ratio(i,j)+midRation
#             c_st_class_cut = jieba.cut_for_search(c_st_class)
#             c_src_nme_cut = jieba.cut_for_search(c_src_nme)
#             for i in c_src_nme_cut:
#                 for j in c_st_class_cut:
#                     midRation = fuzz.token_sort_ratio(i,j)+midRation
            # c_st_clazz_cut = jieba.cut_for_search(c_st_clazz)
            # c_src_nme_cut = jieba.cut_for_search(c_src_nme)
            # for i in c_src_nme_cut:
            #     for j in c_st_clazz_cut:
            #         midRation = fuzz.token_sort_ratio(i,j)+midRation
            if resRation < midRation:
                resRation = midRation
                c_mid_cde = c_st_cde
                c_mid_nme = c_st_nme
        table.iloc[c_seq_no,9]=c_mid_nme
        table.iloc[c_seq_no,10]=c_mid_cde
        table.iloc[c_seq_no,11]=resRation
        print(c_seq_no)
    table.to_excel("企业级客户信息系统源系统码值与标准码值转换映射202012043.xlsx")

'''
  pandas读写xlsx
'''
def readFileByPandasUtils(fileName):
    table = pandas.read_excel('标准码转换.xlsx',sheet_name=0)
    tablest = pandas.read_excel('标准码转换.xlsx',sheet_name=1)
    c_seq_nolist = table.index
    c_seq_nolist_st = tablest.index
    table['标准码值']=None
    table['标准码中文']=None
    table['概率']=None
    print(c_seq_nolist_st)
    for c_seq_no in c_seq_nolist:
        c_src_nme = str(table.iloc[c_seq_no,0])
        c_src_nme1 = str(table.iloc[c_seq_no,0])
        print(c_src_nme)
        resRation = 0
        st_seq_no = 0
        c_mid_cde = ''
        c_mid_nme = ''
        #c_src_cde =  table[['序号','C_SOURCE_CODE']][table['序号'] == c_seq_no]
        for c_seq_no_st in c_seq_nolist_st:
            midRation = 0
            c_st_nme = str(tablest.iloc[c_seq_no_st,2])
            c_st_cde = str(tablest.iloc[c_seq_no_st,1])
            midRation= fuzz.token_set_ratio(c_src_nme, c_st_nme)
            # for i in c_src_nme_cut:
            #     for j in c_st_nme_cut:
            #         midRation = fuzz.token_sort_ratio(i,j)+midRation
#             c_st_class_cut = jieba.cut_for_search(c_st_class)
#             c_src_nme_cut = jieba.cut_for_search(c_src_nme)
#             for i in c_src_nme_cut:
#                 for j in c_st_class_cut:
#                     midRation = fuzz.token_sort_ratio(i,j)+midRation
            # c_st_clazz_cut = jieba.cut_for_search(c_st_clazz)
            # c_src_nme_cut = jieba.cut_for_search(c_src_nme)
            # for i in c_src_nme_cut:
            #     for j in c_st_clazz_cut:
            #         midRation = fuzz.token_sort_ratio(i,j)+midRation
            if resRation < midRation:
                resRation = midRation
                c_mid_cde = c_st_cde
                c_mid_nme = c_st_nme
        table.iloc[c_seq_no,2]=c_mid_nme
        table.iloc[c_seq_no,3]=c_mid_cde
        table.iloc[c_seq_no,4]=resRation
        print(c_seq_no)
    table.to_excel('标准码转换.xlsx')

if __name__ == '__main__':
    #resration  = fuzz.ratio("this is a test", "this is a test!")
    #print(resration)
    #readFile('')
    # readFileByPandas('')
    readFileByPandasUtils('');
    #resration = fuzz.token_sort_ratio("防毒器材装配工", "防化\防毒器材制造\装配工人")
    #print(resration)
    # srcName1 = jieba.cut_for_search("火炮随动系统装试工",HMM)
    # srcName2 = jieba.cut_for_search("枪炮制造人员",HMM=true)
    # srcName21 = jieba.cut_for_search("火工品\爆破器材制造人员",HMM=true)
    #
    # srcName3 = jieba.cut_for_search("火炮随动系统装试工")
    # srcName4 = jieba.cut_for_search("枪炮制造人员")
    # srcName5 = jieba.cut_for_search("火炮随动系统装试工")
    # srcName6 = jieba.cut_for_search("枪炮制造人员")
    # print(", ".join(srcName1))
    # print(", ".join(srcName2))
    # print(", ".join(srcName21))
    # print(", ".join(srcName3))
    # print(", ".join(srcName4))
    # print(", ".join(srcName5))
    # print(", ".join(srcName6))
    #print(fuzz.token_sort_ratio("火炮随动系统装试工","枪炮制造人员"))
    #print(fuzz.token_sort_ratio("火炮随动系统装试工","火工品\爆破器材制造人员"))

