import pandas as pd
import numpy as np
from pandas.core.indexes.base import Index
import xlwt
import xlrd
import os
import time
from xlutils.copy import copy

#开始
var = 1
while var == 1 : #循环
    print('如遇到问题，请尝试关闭已打开的表格再试！')
    print('')
    print('          请选择功能'               )
    print('----------------------------------')
    print('[1]                       收费模式')
    print('[2]                       库存管理')
    print('[3]                       用户充值')
    print('[4]                      查看用户表')
    print('[5]                      查看库存表')
    print('----------------------------------')
    Input = input()
    if (Input == '1'):
        print('已进入收费模式！')
        time.sleep(1)
        os.system("cls")
        print('---------正在初始化，请稍后---------')
        #读表(pandas)
        dfOld=pd.read_excel('用户表.xls')
        df = dfOld.fillna(value = '') #过滤空格
        AvailableLine = len(df)+1 

        #读表(xlrd)
        data = xlrd.open_workbook('用户表.xls',formatting_info=True)

        #获取表格数目
        sheet1 = data.sheet_by_index(0) 
 
        """ 获取所有或某个sheet对象"""
        # 通过index获取第一个sheet对象
        sheet1_object = data.sheet_by_index(0)
        print('sheet对象：',sheet1_object)

        """ 判断某个sheet是否已导入"""
        # 通过index判断sheet1是否导入
        sheet1_is_load = data.sheet_loaded(sheet_name_or_index=0)
        print('导入sheet布尔值：',sheet1_is_load)

        """ 对sheet对象中的列执行操作："""
        # 获取sheet1中的有效列数
        ncols = sheet1_object.ncols
        print('sheet中有效列数',ncols)             

        nrows = sheet1.nrows  #获取该sheet中的有效行数
        print('sheet中有效行数',nrows)

        # 获取sheet1中第colx+0列的数据（名字）
        FileUserName = sheet1_object.col_values(colx=0)
        print(FileUserName)           

        # 获取sheet1中第colx+1列的数据（电话）
        UsersPhoneNumber = sheet1_object.col_values(colx=1)
        print(UsersPhoneNumber)          

        # 获取sheet1中第colx+2列的数据（次数）
        HuFuTimeAfterHuFuTimes = sheet1_object.col_values(colx=2)
        print(HuFuTimeAfterHuFuTimes)          

        # 获取sheet1中第colx+1列的数据（护肤）
        HuFuTimes = sheet1_object.col_values(colx=3)
        print(HuFuTimes)           

        # 使用xlutils将xlrd读取的对象转为xlwt可操作对象
        workbook = copy(data)# 完成xlrd对象向xlwt对象转换
        writebook = xlwt.Workbook()
        worksheet = workbook.get_sheet(0) # 获得要操作的页
        table = data.sheets()[0]

        print('---------初始化完成，准备进入程序---------')
        time.sleep(0.8)
        os.system("cls")
        print('请输入服务类型（美甲/护肤）')
        UserInput = input()
        if (UserInput == '美甲'):
            print('设定成功！用户类型为“美甲”')
            time.sleep(0.6)
            os.system("cls")
            print('请输入用户名')
            UserName = input()
            if (UserName in FileUserName):
                #搜索用户所在行
                index = df[df["姓名"]== UserName].index.tolist()[0]
                # 获取用户名所在行的所有内容
                all_row_values = sheet1_object.row_values(rowx=index+1)
                print('找到了该用户的历史！')
                print(all_row_values)
                print('确定？（y/n）')
                if (input() == 'y'):
                    # 获取次数
                    HuFuTimeAfterHuFuTime = df.iloc[index, 2]
                    print('请输入本次扣费金额')
                    ShouldMoney = input()
                    if (int(HuFuTimeAfterHuFuTime) < int(ShouldMoney)):
                        print('失败:该用户余额不足!')
                        time.sleep(1)
                        os.system("cls")
                    else:
                        AfterHuFuTime = int(HuFuTimeAfterHuFuTime) - int(ShouldMoney)
                        print('成功!剩余余额:',AfterHuFuTime)
                        # 写入一个值，括号内分别为行数、列数、内容
                        worksheet.write(index+1,2,str(AfterHuFuTime))
                        workbook.save('用户表.xls')
                        print('完成')
                        time.sleep(1)
                        os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
            else:
                print('未找到该用户的历史！')
                print('是否新建该用户信息？（y/n）')
                if (input() == 'y'):
                    print('请输入用户',UserName,'的电话号码')
                    AddUserPhone = input()
                    print('请输入该用户的可用护肤余额')
                    AddHuFuTime = input()
                    print('请输入该用户的可用美甲余额')
                    AddMeiJiaTime = input()
                    print('正在新建用户数据,请稍等')
                    worksheet.write(AvailableLine, 0, UserName)
                    worksheet.write(AvailableLine, 1, AddUserPhone)
                    worksheet.write(AvailableLine, 2, AddMeiJiaTime)
                    worksheet.write(AvailableLine, 3, AddHuFuTime)
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
        elif (UserInput == '护肤'):
            print('设定成功！用户类型为“护肤”')
            time.sleep(1)
            os.system("cls")
            print('请输入用户名')
            UserName = input()
            if (UserName in FileUserName):
                #搜索用户所在行
                index = df[df["姓名"]== UserName].index.tolist()[0]
                # 获取用户名所在行
                all_row_values = sheet1_object.row_values(rowx=index+1)
                print('找到了该用户的历史！')
                print(all_row_values)
                print('确定？（y/n）')
                if (input() == 'y'):
                    print('正在写入日期，请稍后')
                    time.sleep(0.7)
                    Info = [i for i in all_row_values if i != '']    #删除空值
                    AvailableRow = len(Info)
                    localtime = time.strftime("%Y-%m-%d", time.localtime())
                    worksheet.write(index+1,AvailableRow,localtime)    #写入日期
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
 
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
            else:
                print('未找到该用户的历史！')
                print('是否新建该用户信息？（y/n）')
                if (input() == 'y'):
                    print('请输入用户',UserName,'的电话号码')
                    AddUserPhone = input()
                    print('请输入该用户的可用护肤余额')
                    AddHuFuTime = input()
                    print('请输入该用户的可用美甲余额')
                    AddMeiJiaTime = input()
                    print('正在新建用户数据,请稍等')
                    worksheet.write(AvailableLine, 0, UserName)
                    worksheet.write(AvailableLine, 1, AddUserPhone)
                    worksheet.write(AvailableLine, 2, AddMeiJiaTime)
                    worksheet.write(AvailableLine, 3, AddHuFuTime)
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
        else:
            print('未找到该指令!')
            time.sleep(1)
            os.system("cls")
    elif (Input == '2'):
        print('已进入库存管理模式！')
        time.sleep(1)
        os.system("cls")
        print('---------正在初始化，请稍后---------')
        #读表(pandas)
        df=pd.read_excel('库存表.xls')
        AvailableLine = len(df)+1 #可用行数

        #读表(xlrd)
        data = xlrd.open_workbook('库存表.xls',formatting_info=True)

        #获取表格数目
        sheet1 = data.sheet_by_index(0) 
 
        """ 获取所有或某个sheet对象"""
        # 通过index获取第一个sheet对象
        sheet1_object = data.sheet_by_index(0)
        print('sheet对象：',sheet1_object)

        """ 判断某个sheet是否已导入"""
        # 通过index判断sheet1是否导入
        sheet1_is_load = data.sheet_loaded(sheet_name_or_index=0)
        print('导入sheet布尔值：',sheet1_is_load)

        """ 对sheet对象中的列执行操作："""
        # 获取sheet1中的有效列数
        ncols = sheet1_object.ncols
        print('sheet中有效列数',ncols)             

        nrows = sheet1.nrows  #获取该sheet中的有效行数
        print('sheet中有效行数',nrows)

        # 获取sheet1中第colx+0列的数据（名字）
        ObjectsName = sheet1_object.col_values(colx=0)
        print(ObjectsName)           

        # 获取sheet1中第colx+1列的数据（库存）
        RestNumber = sheet1_object.col_values(colx=1)
        print(RestNumber)                

        # 使用xlutils将xlrd读取的对象转为xlwt可操作对象
        workbook = copy(data)# 完成xlrd对象向xlwt对象转换
        writebook = xlwt.Workbook()
        worksheet = workbook.get_sheet(0) # 获得要操作的页
        table = data.sheets()[0]

        print('---------初始化完成，准备进入程序---------')
        time.sleep(0.8)
        os.system("cls")
        print('请输入商品名')
        InputName = input()
        if (InputName in ObjectsName):
            #搜索商品所在行
            index = df[df["商品名"]== InputName].index.tolist()[0]
            #获取商品名所在行
            all_row_values = sheet1_object.row_values(rowx=index+1)
            print('找到了该商品！')
            print(all_row_values)
            print('确定？（y/n）')
            if (input() == 'y'):
                # 获取库存
                ObjectNumber = df.iloc[index, 1]
                print('商品',InputName,'的库存量为：',ObjectNumber,'件')
                print('请选择功能：【1】：进货')
                print('           【2】：使用')
                Input2 = input()
                if (Input2 == '1'):
                    print('已进入进货模式！')
                    time.sleep(1.0)
                    os.system("cls")
                    print('请输入进货数量')
                    NewObjectNumber = input()
                    print('处理中...')
                    time.sleep(0.5)
                    AfterObjectNumber = int(ObjectNumber) + int(NewObjectNumber)
                    worksheet.write(index+1,1,str(AfterObjectNumber))
                    workbook.save('库存表.xls')
                    print('剩余库存:',AfterObjectNumber)
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('已进入使用库存模式！')
                    time.sleep(1.0)
                    os.system("cls")
                    print('请输入使用数量')
                    NewObjectNumber = input()
                    print('处理中...')
                    time.sleep(0.5)
                    AfterObjectNumber = int(ObjectNumber) - int(NewObjectNumber)
                    print('剩余库存:',AfterObjectNumber)
                    worksheet.write(index+1,1,str(AfterObjectNumber))
                    workbook.save('库存表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")              
            else:
                print('退出')
                time.sleep(1)
                os.system("cls")
        else:
            print('未找到该商品，是否新增该商品信息？（y/n）')
            if (input() == 'y'):
                print('请输入商品',InputName,'的库存量')
                AddObjectNumber = input()
                print('添加中，请稍后....')
                worksheet.write(AvailableLine, 0, InputName)
                worksheet.write(AvailableLine, 1, AddObjectNumber)
                time.sleep(1)
                workbook.save('库存表.xls')
                print('完成')
                time.sleep(1)
                os.system("cls")
    elif (Input == '3'):
        print('已进入用户充值模式！')
        time.sleep(1)
        os.system("cls")
        print('---------正在初始化，请稍后---------')
        #读表(pandas)
        df=pd.read_excel('用户表.xls')
        AvailableLine = len(df)+1 #可用行数

        #读表(xlrd)
        data = xlrd.open_workbook('用户表.xls',formatting_info=True)

        #获取表格数目
        sheet1 = data.sheet_by_index(0) 
 
        """ 获取所有或某个sheet对象"""
        # 通过index获取第一个sheet对象
        sheet1_object = data.sheet_by_index(0)
        print('sheet对象：',sheet1_object)

        """ 判断某个sheet是否已导入"""
        # 通过index判断sheet1是否导入
        sheet1_is_load = data.sheet_loaded(sheet_name_or_index=0)
        print('导入sheet布尔值：',sheet1_is_load)

        """ 对sheet对象中的列执行操作："""
        # 获取sheet1中的有效列数
        ncols = sheet1_object.ncols
        print('sheet中有效列数',ncols)             

        nrows = sheet1.nrows  #获取该sheet中的有效行数
        print('sheet中有效行数',nrows)

        # 获取sheet1中第colx+0列的数据（名字）
        FileUserName = sheet1_object.col_values(colx=0)
        print(FileUserName)           

        # 获取sheet1中第colx+1列的数据（电话）
        UsersPhoneNumber = sheet1_object.col_values(colx=1)
        print(UsersPhoneNumber)          

        # 获取sheet1中第colx+2列的数据（次数）
        HuFuTimeAfterHuFuTimes = sheet1_object.col_values(colx=2)
        print(HuFuTimeAfterHuFuTimes)          

        # 获取sheet1中第colx+1列的数据（护肤）
        HuFuTimes = sheet1_object.col_values(colx=3)
        print(HuFuTimes)           

        # 使用xlutils将xlrd读取的对象转为xlwt可操作对象
        workbook = copy(data)# 完成xlrd对象向xlwt对象转换
        writebook = xlwt.Workbook()
        worksheet = workbook.get_sheet(0) # 获得要操作的页
        table = data.sheets()[0]
        print('---------初始化完成，准备进入程序---------')
        time.sleep(0.8)
        os.system("cls")
        
        print('请输入服务类型（美甲/护肤）')
        UserInput = input()
        if (UserInput == '美甲'):
            print('设定成功！用户类型为“美甲”')
            time.sleep(1)
            os.system("cls")
            print('请输入用户名')
            UserName = input()
            if (UserName in FileUserName):
                #搜索用户所在行
                index = df[df["姓名"]== UserName].index.tolist()[0]
                # 获取用户名所在行
                all_row_values = sheet1_object.row_values(rowx=index+1)
                print('找到了该用户的历史！')
                print(all_row_values)
                print('确定？（y/n）')
                if (input() == 'y'):
                    # 获取余额
                    HuFuTime = df.iloc[index, 2]
                    print('请输入充值金额（美甲）')
                    AddMoney = input()
                    AfterHuFuTime = int(HuFuTime) + int(AddMoney)
                    print('成功!剩余余额:',AfterHuFuTime)
                     # 写入一个值，括号内分别为行数、列数、内容
                    worksheet.write(index+1,2,str(AfterHuFuTime))
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
            else:
                print('未找到该用户的历史！')
                print('是否新建该用户信息？（y/n）')
                if (input() == 'y'):
                    print('请输入用户',UserName,'的电话号码')
                    AddUserPhone = input()
                    print('请输入该用户的可用护肤余额')
                    AddHuFuTime = input()
                    print('请输入该用户的可用美甲余额')
                    AddMeiJiaTime = input()
                    print('正在新建用户数据,请稍等')
                    worksheet.write(AvailableLine, 0, UserName)
                    worksheet.write(AvailableLine, 1, AddUserPhone)
                    worksheet.write(AvailableLine, 2, AddMeiJiaTime)
                    worksheet.write(AvailableLine, 3, AddHuFuTime)
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
        elif (UserInput == '护肤'):
            print('设定成功！用户类型为“护肤”')
            time.sleep(1)
            os.system("cls")
            print('请输入用户名')
            UserName = input()
            if (UserName in FileUserName):
                #搜索用户所在行
                index = df[df["姓名"]== UserName].index.tolist()[0]
                # 获取用户名所在行
                all_row_values = sheet1_object.row_values(rowx=index+1)
                print('找到了该用户的历史！')
                print(all_row_values)
                print('确定？（y/n）')
                if (input() == 'y'):
                    # 获取余额
                    HuFuTime = df.iloc[index, 3]
                    print('请输入充值金额（护肤）')
                    ShouldMoney = input()
                    AfterHuFuTime = int(HuFuTime) + int(ShouldMoney)
                    print('成功!剩余余额:',AfterHuFuTime)
                    # 写入一个值，括号内分别为行数、列数、内容
                    worksheet.write(index+1,2,str(AfterHuFuTime))
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
            else:
                print('未找到该用户的历史！')
                print('是否新建该用户信息？（y/n）')
                if (input() == 'y'):
                    print('请输入用户',UserName,'的电话号码')
                    AddUserPhone = input()
                    print('请输入该用户的可用护肤余额')
                    AddHuFuTime = input()
                    print('请输入该用户的可用美甲余额')
                    AddMeiJiaTime = input()
                    print('正在新建用户数据,请稍等')
                    worksheet.write(AvailableLine, 0, UserName)
                    worksheet.write(AvailableLine, 1, AddUserPhone)
                    worksheet.write(AvailableLine, 2, AddMeiJiaTime)
                    worksheet.write(AvailableLine, 3, AddHuFuTime)
                    workbook.save('用户表.xls')
                    print('完成')
                    time.sleep(1)
                    os.system("cls")
                else:
                    print('退出')
                    time.sleep(1)
                    os.system("cls")
    elif(Input == '4'):
        print('读取表格中，请稍后')
        time.sleep(0.4)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        pd.set_option('display.width', 180) # 设置打印宽度(**重要**)
        df = pd.read_excel('用户表.xls')  #pd
        print('-----------------------------------------')
        print(df)
        print('-----------------------------------------')
        print('请按任意键退出')
        input()
        os.system("cls")
    elif(Input == '5'):
        print('读取表格中，请稍后')
        time.sleep(0.4)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        pd.set_option('display.width', 180) # 设置打印宽度(**重要**)
        df=pd.read_excel('库存表.xls')
        print('-------------------')
        print(df)
        print('-------------------')
        print('请按任意键退出')
        input()
        os.system("cls")
    else:
            print('未找到该指令!')
            time.sleep(1)
            os.system("cls")
        
    
        
                
                
                
