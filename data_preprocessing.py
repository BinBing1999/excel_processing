#Author: Chen Yibin , all rights preserved
#contact:  rasbinbin@sjtu.edu.cn
#This python file is for disposing unnecessary data rows in excel files.
#To configure conditions for deleting, set up "DeleteOrNot" function by yourself.
#2022/11/6

import os
import numpy
import openpyxl
#import pandas
import time
#alt +3 #alt +4#用于注释

def DeleteOrNot(theList):#使用判定条件决定是否删除这个case
    #if '開單' not in theList:#存在性判定
        #return 1
    if '發藥' not in theList:
        return 1
    if theList.count('收費')>1:#数量判定
        return 1
    if theList.count('開方')>1:#数量判定
        return 1
    if theList.count('發藥')>1:#数量判定
        return 1
##    if theList.count('收費')>1:#数量判定
##        return 1
    #if '開方' not in theList:
     #   return 1
    if theList[-1]=='首次接診':#最后结尾
        return 1
    if theList[-1]=='配藥':
        return 1
    if theList[-1]=='檢查':
        return 1
    return 0#没问题的数据case，返回0

path = r"C:\Users\user\Desktop\BPM\hw1_2022data\preprocessed_data"
os.chdir(path)
workbook = openpyxl.load_workbook('ProData2021 (1)_traditionalChiese.xlsx')
#print("Your sheet name is:")
#print (workbook.sheetnames)
sheet = workbook['Sheet 111']#注意自己的sheet名字，我的是Sheet 111
#print(sheet.dimensions)
case_cell=sheet['A']#选择case列
activity_cell=sheet['F']#选择activity列
#第一行是描述行，删除和遍历从第二行开始
#print(len(case_cell))
print("Initializing...")

current_case=0# initialize guahao id to 0 (no conflict)
activity_list=[]
temp_cnt=0

determined_cnt=0
deleted_cnt=0

#最后一次性删除不要的数据，位置先以列表保存
deleting_base=[]
deleting_bias=[]
row_base=0
row_bias=0#base + bias = address

#test_cell=case_cell[10747]
#print(test_cell.value)
print("Now running...")
for ii in range(len(case_cell)):##len(case_cell)---78
    #start from 0 to n-1
    if ii==0:
        continue#跳过第一行描述行
    target_cell=case_cell[ii]
    target_activity=activity_cell[ii]
    #print(target_cell.value)
    
    if current_case!=target_cell.value:
        if current_case==0:
            current_case=target_cell.value
            activity_list.append(target_activity.value)
            row_base=ii+1
            temp_cnt=1
        else:
            if DeleteOrNot(activity_list):
                row_bias=temp_cnt
                deleting_base.append(row_base)
                deleting_bias.append(row_bias)
            row_base=ii+1
            temp_cnt=1
            print("Case "+str(current_case)+" determined.")#print case number
            determined_cnt+=1
            current_case=target_cell.value
            activity_list=[]
            activity_list.append(target_activity.value)
            
    else:#when case number equals
        activity_list.append(target_activity.value)
        temp_cnt+=1

    if ii==len(case_cell)-1:##len(case_cell)----78
        if DeleteOrNot(activity_list):
            row_bias=temp_cnt
            deleting_base.append(row_base)
            deleting_bias.append(row_bias)
        print("Case "+str(current_case)+" determined.")#print case number
        determined_cnt+=1

print("Showing intervals to be deleted:")
print("Deleted base:")
print(deleting_base)
print("Deleted bias:")
print(deleting_bias)
#print(deleting_base[1])

#reversed list,从后往前删，删除一轮后excel表格自动往上移
#len(deleting_base) = len(deleting_bias), no worry
for kk in range(len(deleting_base)):
    row_base=deleting_base[len(deleting_base)-1-kk]
    row_bias=deleting_bias[len(deleting_bias)-1-kk]
    print("Case "+str(case_cell[row_base-1].value)+" deleted.")
    deleted_cnt+=1
    sheet.delete_rows(idx=row_base, amount=row_bias)#deleting


savename='test_'+str(int(time.time()))+'.xlsx'# world time since 1970 in seconds for time mark
workbook.save(savename)#always the newest name
print("Done")
print("Total determined cases: "+str(determined_cnt))
print("Total deleted cases: "+str(deleted_cnt))
print("Cases remained: "+str(determined_cnt-deleted_cnt))
