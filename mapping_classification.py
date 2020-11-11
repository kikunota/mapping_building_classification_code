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
    hyoushi_wb = openpyxl.load_workbook("標準仕様書.xlsx")
    hyoushi_ws = hyoushi_wb["Sheet2"]

    hyoushi_JP_index = manipTable.getColumnValueByName(hyoushi_ws, "Column1")
    hyoushi_JP_code = manipTable.getColumnValueByName(hyoushi_ws, "Column5")

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
    manipTable.insertColumnByValueB(hyoushi_ws, "I2", hyou_SysList)
    hyoushi_wb.save('result_hyoushi.xlsx')

    #print(UniToHyou_ind)
    #print(UniToHyou_val)


if __name__=="__main__":
    start_time = time.time()
    Main()
    print("--- %s seconds ---" % (time.time() - start_time))
