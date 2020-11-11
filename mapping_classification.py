#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
print(sys.prefix)

from lib import ManipulateTable_v1
import time
import openpyxl
import difflib

class Mapping_classification():
    def __init__(self):
        pass

def getHyoushi_index():
    hyoushi_index = ['[1, 0, 0]', '[1, 1, 0]', '[1, 1, 1]', '[1, 1, 2]', '[1, 1, 3]', '[1, 1, 4]', '[1, 1, 5]', '[1, 1, 6]', '[1, 1, 7]', '[1, 1, 8]', '[1, 1, 9]', '[1, 1, 10]', '[1, 1, 11]', '[1, 1, 12]', '[1, 1, 13]', '[1, 2, 0]', '[1, 2, 1]', '[1, 2, 2]', '[1, 2, 3]', '[1, 2, 4]', '[1, 3, 0]', '[1, 3, 1]', '[1, 3, 2]', '[1, 3, 3]', '[1, 3, 4]', '[1, 3, 5]', '[1, 3, 6]', '[1, 3, 7]', '[1, 3, 8]', '[1, 3, 9]', '[1, 3, 10]', '[1, 3, 11]', '[1, 3, 12]', '[1, 3, 13]', '[1, 4, 0]', '[1, 4, 1]', '[1, 4, 2]', '[1, 4, 3]', '[1, 4, 4]', '[1, 4, 5]', '[1, 4, 6]', '[1, 5, 0]', '[1, 5, 1]', '[1, 5, 2]', '[1, 5, 3]', '[1, 5, 4]', '[1, 5, 5]', '[1, 5, 6]', '[1, 5, 7]', '[1, 5, 8]', '[1, 5, 9]', '[1, 6, 0]', '[1, 6, 1]', '[1, 6, 2]', '[1, 7, 0]', '[1, 7, 1]', '[1, 7, 2]', '[1, 7, 3]', '[2, 0, 0]', '[2, 1, 0]', '[2, 1, 1]', '[2, 1, 2]', '[2, 2, 0]', '[2, 2, 1]', '[2, 2, 2]', '[2, 2, 3]', '[2, 2, 4]', '[2, 3, 0]', '[2, 3, 1]', '[2, 3, 2]', '[2, 3, 3]', '[2, 4, 0]', '[2, 4, 1]', '[3, 0, 0]', '[3, 1, 0]', '[3, 1, 1]', '[3, 1, 2]', '[3, 1, 3]', '[3, 2, 0]', '[3, 2, 1]', '[3, 2, 2]', '[3, 2, 3]', '[3, 2, 4]', '[3, 2, 5]', '[3, 3, 0]', '[3, 3, 1]', '[3, 3, 2]', '[3, 3, 3]', '[4, 0, 0]', '[4, 1, 0]', '[4, 1, 1]', '[4, 1, 2]', '[4, 1, 3]', '[4, 2, 0]', '[4, 2, 1]', '[4, 2, 2]', '[4, 2, 3]', '[4, 2, 4]', '[4, 2, 5]', '[4, 3, 0]', '[4, 3, 1]', '[4, 3, 2]', '[4, 3, 3]', '[4, 3, 4]', '[4, 3, 5]', '[4, 3, 6]', '[4, 3, 7]', '[4, 3, 8]', '[4, 3, 9]', '[4, 4, 0]', '[4, 4, 1]', '[4, 4, 2]', '[4, 4, 3]', '[4, 4, 4]', '[4, 4, 5]', '[4, 4, 6]', '[4, 4, 7]', '[4, 5, 0]', '[4, 5, 1]', '[4, 5, 2]', '[4, 5, 3]', '[4, 5, 4]', '[4, 5, 5]', '[4, 5, 6]', '[4, 5, 7]', '[4, 5, 8]', '[4, 6, 0]', '[4, 6, 1]', '[4, 6, 2]', '[4, 6, 3]', '[4, 6, 4]', '[4, 6, 5]', '[4, 6, 6]', '[5, 0, 0]', '[5, 1, 0]', '[5, 1, 1]', '[5, 1, 2]', '[5, 1, 3]', '[5, 2, 0]', '[5, 2, 1]', '[5, 2, 2]', '[5, 2, 3]', '[5, 3, 0]', '[5, 3, 1]', '[5, 3, 2]', '[5, 3, 3]', '[5, 3, 4]', '[5, 3, 5]', '[5, 3, 6]', '[5, 3, 7]', '[5, 4, 0]', '[5, 4, 1]', '[5, 4, 2]', '[5, 4, 3]', '[5, 4, 4]', '[5, 4, 5]', '[5, 4, 6]', '[5, 4, 7]', '[5, 4, 8]', '[5, 4, 9]', '[5, 4, 10]', '[5, 4, 11]', '[5, 5, 0]', '[5, 5, 1]', '[5, 5, 2]', '[5, 5, 3]', '[5, 6, 0]', '[5, 6, 1]', '[5, 6, 2]', '[5, 6, 3]', '[5, 6, 4]', '[5, 6, 5]', '[6, 0, 0]', '[6, 1, 0]', '[6, 1, 1]', '[6, 1, 2]', '[6, 2, 0]', '[6, 2, 1]', '[6, 2, 2]', '[6, 2, 3]', '[6, 2, 4]', '[6, 2, 5]', '[6, 3, 0]', '[6, 3, 1]', '[6, 3, 2]', '[6, 4, 0]', '[6, 4, 1]', '[6, 4, 2]', '[6, 4, 3]', '[6, 4, 4]', '[6, 5, 0]', '[6, 5, 1]', '[6, 5, 2]', '[6, 5, 3]', '[6, 5, 4]', '[6, 5, 5]', '[6, 6, 0]', '[6, 6, 1]', '[6, 6, 2]', '[6, 6, 3]', '[6, 6, 4]', '[6, 6, 5]', '[6, 6, 6]', '[6, 6, 7]', '[6, 7, 0]', '[6, 7, 1]', '[6, 7, 2]', '[6, 7, 3]', '[6, 8, 0]', '[6, 8, 1]', '[6, 8, 2]', '[6, 8, 3]', '[6, 8, 4]', '[6, 8, 5]', '[6, 9, 0]', '[6, 9, 1]', '[6, 9, 2]', '[6, 9, 3]', '[6, 9, 4]', '[6, 9, 5]', '[6, 9, 6]', '[7, 0, 0]', '[7, 1, 0]', '[7, 1, 1]', '[7, 1, 2]', '[7, 1, 3]', '[7, 1, 4]', '[7, 2, 0]', '[7, 2, 1]', '[7, 2, 2]', '[7, 2, 3]', '[7, 2, 4]', '[7, 2, 5]', '[7, 2, 6]', '[7, 2, 7]', '[7, 2, 8]', '[7, 2, 9]', '[7, 2, 10]', '[7, 3, 0]', '[7, 3, 1]', '[7, 3, 2]', '[7, 3, 3]', '[7, 3, 4]', '[7, 3, 5]', '[7, 3, 6]', '[7, 3, 7]', '[7, 3, 8]', '[7, 3, 9]', '[7, 3, 10]', '[7, 3, 11]', '[7, 4, 0]', '[7, 4, 1]', '[7, 4, 2]', '[7, 4, 3]', '[7, 4, 4]', '[7, 4, 5]', '[7, 4, 6]', '[7, 4, 7]', '[7, 4, 8]', '[7, 4, 9]', '[7, 5, 0]', '[7, 5, 1]', '[7, 5, 2]', '[7, 6, 0]', '[7, 6, 1]', '[7, 6, 2]', '[7, 6, 3]', '[7, 6, 4]', '[7, 6, 5]', '[7, 6, 6]', '[7, 6, 7]', '[7, 6, 8]', '[7, 6, 9]', '[7, 6, 10]', '[7, 6, 11]', '[7, 6, 12]', '[7, 6, 13]', '[7, 7, 0]', '[7, 7, 1]', '[7, 7, 2]', '[7, 7, 3]', '[7, 7, 4]', '[7, 7, 5]', '[7, 7, 6]', '[7, 7, 7]', '[7, 7, 8]', '[7, 8, 0]', '[7, 8, 1]', '[7, 8, 2]', '[7, 8, 3]', '[7, 8, 4]', '[7, 9, 0]', '[7, 9, 1]', '[7, 9, 2]', '[7, 9, 3]', '[7, 9, 4]', '[7, 9, 5]', '[7, 9, 6]', '[7, 9, 7]', '[7, 9, 8]', '[7, 9, 9]', '[8, 0, 0]', '[8, 1, 0]', '[8, 1, 1]', '[8, 1, 2]', '[8, 2, 0]', '[8, 2, 1]', '[8, 2, 2]', '[8, 2, 3]', '[8, 2, 4]', '[8, 2, 5]', '[8, 2, 6]', '[8, 2, 7]', '[8, 2, 8]', '[8, 2, 9]', '[8, 2, 10]', '[8, 2, 11]', '[8, 3, 0]', '[8, 3, 1]', '[8, 3, 2]', '[8, 3, 3]', '[8, 3, 4]', '[8, 3, 5]', '[8, 3, 6]', '[8, 3, 7]', '[8, 3, 8]', '[8, 3, 9]', '[8, 3, 10]', '[8, 4, 0]', '[8, 4, 1]', '[8, 4, 2]', '[8, 4, 3]', '[8, 4, 4]', '[8, 4, 5]', '[8, 4, 6]', '[8, 4, 7]', '[8, 5, 0]', '[8, 5, 1]', '[8, 5, 2]', '[8, 5, 3]', '[8, 5, 4]', '[8, 5, 5]', '[9, 0, 0]', '[9, 1, 0]', '[9, 1, 1]', '[9, 1, 2]', '[9, 1, 3]', '[9, 2, 0]', '[9, 2, 1]', '[9, 2, 2]', '[9, 2, 3]', '[9, 2, 4]', '[9, 2, 5]', '[9, 3, 0]', '[9, 3, 1]', '[9, 3, 2]', '[9, 3, 3]', '[9, 3, 4]', '[9, 4, 0]', '[9, 4, 1]', '[9, 4, 2]', '[9, 4, 3]', '[9, 4, 4]', '[9, 5, 0]', '[9, 5, 1]', '[9, 5, 2]', '[9, 5, 3]', '[9, 5, 4]', '[9, 6, 0]', '[9, 6, 1]', '[9, 6, 2]', '[9, 6, 3]', '[9, 6, 4]', '[9, 7, 0]', '[9, 7, 1]', '[9, 7, 2]', '[9, 7, 3]', '[9, 7, 4]', '[9, 7, 5]', '[10, 0, 0]', '[10, 1, 0]', '[10, 2, 0]', '[10, 3, 0]', '[10, 4, 0]', '[10, 5, 0]', '[10, 6, 0]', '[10, 7, 0]', '[11, 0, 0]', '[11, 1, 0]', '[11, 2, 0]', '[11, 3, 0]', '[12, 0, 0]', '[12, 1, 0]', '[12, 2, 0]', '[12, 3, 0]', '[12, 4, 0]', '[12, 5, 0]', '[12, 6, 0]', '[12, 7, 0]', '[13, 0, 0]', '[13, 1, 0]', '[13, 2, 0]', '[13, 3, 0]', '[13, 4, 0]', '[13, 5, 0]', '[14, 0, 0]', '[14, 1, 0]', '[14, 2, 0]', '[14, 3, 0]', '[14, 4, 0]', '[14, 5, 0]', '[14, 6, 0]', '[14, 7, 0]', '[14, 8, 0]', '[15, 0, 0]', '[15, 1, 0]', '[15, 2, 0]', '[15, 3, 0]', '[15, 4, 0]', '[15, 5, 0]', '[15, 6, 0]', '[15, 7, 0]', '[15, 8, 0]', '[15, 9, 0]', '[16, 0, 0]', '[16, 1, 0]', '[16, 2, 0]', '[16, 3, 0]', '[16, 4, 0]', '[16, 5, 0]', '[16, 6, 0]', '[16, 7, 0]', '[16, 8, 0]', '[16, 9, 0]', '[17, 0, 0]', '[17, 1, 0]', '[17, 2, 0]', '[17, 3, 0]', '[18, 0, 0]', '[18, 1, 0]', '[18, 2, 0]', '[18, 3, 0]', '[18, 4, 0]', '[18, 5, 0]', '[18, 6, 0]', '[18, 7, 0]', '[18, 8, 0]', '[18, 9, 0]', '[19, 0, 0]', '[19, 1, 0]', '[19, 2, 0]', '[19, 3, 0]', '[19, 4, 0]', '[19, 5, 0]', '[19, 6, 0]', '[19, 7, 0]', '[19, 8, 0]', '[19, 9, 0]', '[20, 0, 0]', '[20, 1, 0]', '[20, 2, 0]', '[20, 3, 0]', '[20, 4, 0]', '[21, 0, 0]', '[21, 1, 0]', '[21, 2, 0]', '[21, 3, 0]', '[22, 0, 0]', '[22, 1, 0]', '[22, 2, 0]', '[22, 3, 0]', '[22, 4, 0]', '[22, 5, 0]', '[22, 6, 0]', '[22, 7, 0]', '[22, 8, 0]', '[22, 9, 0]', '[23, 0, 0]', '[23, 1, 0]', '[23, 2, 0]', '[23, 3, 0]', '[23, 4, 0]', '[23, 5, 0]', None, None, None, None]
    return hyoushi_index

def compareTwoLists(uniclass_JP_code, hyoushi_JP_code):
    #uniclass_JP_code, original lists that you want to map
    #hyoushi_JP_code, compare lists that you want to map to

    matched_list = []
    matched_ind_list = []
    for i in range(len(uniclass_JP_code)):
    #for i in range(100):
        matched_val = -1
        matched_ind = -1
        for j in range(len(hyoushi_JP_code)):
        #for j in range(100):
            if uniclass_JP_code[i] != None and hyoushi_JP_code[j] != None:
                matched_temp = difflib.SequenceMatcher(None, uniclass_JP_code[i], hyoushi_JP_code[j]).ratio()
                if matched_val < matched_temp:
                    matched_val = matched_temp
                    matched_ind = hyoushi_JP_code[j]
        matched_list.append(matched_val)
        matched_ind_list.append(matched_ind)

    return matched_ind_list, matched_list
def compareTwoListsWithIndex(uniclass_JP_code, hyoushi_JP_code,  hyoushi_JP_index):
    #uniclass_JP_code, original lists that you want to map
    #hyoushi_JP_code, compare lists that you want to map to

    matched_list = []
    matched_ind_list = []
    matched_tit_list = []
    for i in range(len(uniclass_JP_code)):
    #for i in range(100):
        matched_val = -1
        matched_ind = -1
        matched_tit = -1
        for j in range(len(hyoushi_JP_code)):
        #for j in range(100):
            if uniclass_JP_code[i] != None and hyoushi_JP_code[j] != None:
                matched_temp = difflib.SequenceMatcher(None, uniclass_JP_code[i], hyoushi_JP_code[j]).ratio()
                if matched_val < matched_temp and matched_temp >= 0.50:
                    matched_val = matched_temp
                    matched_ind = hyoushi_JP_index[j]
                    matched_tit = hyoushi_JP_code[j]
        matched_list.append(matched_val)
        matched_ind_list.append(matched_ind)
        matched_tit_list.append(matched_tit)
    return matched_ind_list, matched_tit_list, matched_list

def cleanUpTerms(lst):
    for i in range(len(lst)):
        if lst[i] != None:
            if "システム" in lst[i]:
                lst[i] = lst[i].replace("システム", "")
    print(lst)
    return lst

def Main():
    manipTable = ManipulateTable_v1.ManipulateTable()

    #Extract list from `標準仕様書_JP
    hyoushi_wb = openpyxl.load_workbook("Omniclass.xlsx")
    hyoushi_ws = hyoushi_wb["Sheet2"]

    hyoushi_JP_index = manipTable.getColumnValueByName(hyoushi_ws, "Column2")
    hyoushi_JP_code = manipTable.getColumnValueByName(hyoushi_ws, "Column1")

    #Extract list from Uniclass2015_JP
    uniclass_wb = openpyxl.load_workbook("Uniclass_System.xlsx")
    uniclass_ws = uniclass_wb["Sheet1"]

    uniclass_JP_index = manipTable.getColumnValueByName(uniclass_ws, "code")
    uniclass_JP_code = manipTable.getColumnValueByName(uniclass_ws, "title_jp")


    #Extract list sample
    #hyoushi_JP_index = getHyoushi_index()
    #hyoushi_JP_code = getHyoushi_code()
    #uniclass_JP_index = getUniclass_index()
    #uniclass_JP_code = getUniclass_code()

    hyoushi_JP_code = cleanUpTerms(hyoushi_JP_code)
    uniclass_JP_code = cleanUpTerms(uniclass_JP_code)

    #UniToHyou_code, UniToHyou_val = compareTwoLists(uniclass_JP_code, hyoushi_JP_code)
    UniToHyou_ind, UniToHyou_tit, UniToHyou_val = compareTwoListsWithIndex(uniclass_JP_code, hyoushi_JP_code, hyoushi_JP_index)
    manipTable.insertColumnByValueB(uniclass_ws, "D2", UniToHyou_ind)
    manipTable.insertColumnByValueB(uniclass_ws, "E2", UniToHyou_tit)
    manipTable.insertColumnByValueB(uniclass_ws, "F2", UniToHyou_val)
    #uniclass_wb.save('result_uniclass.xlsx')

    #map index values from uniclass to hyoushi
    hyou_SysList = []
    for i in range(len(hyoushi_JP_index)):
        sysList =[]
        for j in range(len(uniclass_JP_code)):
            if UniToHyou_ind[j] != -1:
                if UniToHyou_ind[j] == hyoushi_JP_index[i]:
                    sysList.append(uniclass_JP_code[j])
        hyou_SysList.append(sysList)
    #print(hyou_SysList)

    uniclass_wb.save('result_uniclass.xlsx')

    """
    #map from hyoushi to uniclass
    HyouToUni_ind, HyouToUni_val = compareTwoLists(hyoushi_JP_code, uniclass_JP_code)
    manipTable.insertColumnByValueB(hyoushi_ws, "H2", HyouToUni_ind)
    """
    manipTable.insertColumnByValueB(hyoushi_ws, "M2", hyou_SysList)
    hyoushi_wb.save('result_hyoushi.xlsx')

    #print(UniToHyou_ind)
    #print(UniToHyou_val)


if __name__=="__main__":
    start_time = time.time()
    Main()
    print("--- %s seconds ---" % (time.time() - start_time))
