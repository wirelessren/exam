#!/usr/bin/python3.7
# -*- coding: utf-8 -*-
import pandas as pd
import configparser
from datetime import datetime
import os
import sys
#配置文件读取
##pyinstaller -F -w -i 2.ico out_for_system.py
##pyinstaller -D -w -i 2.ico out_for_system.py#多文件
filepath = os.path.join(os.getcwd(),'config.ini')
cp = configparser.ConfigParser()
cp.read(filepath, encoding="utf-8-sig")
#设置
阅卷系统 = int(cp.get('系统设置','阅卷系统'))

#欧码
欧码科类代码 = cp.get('考场设置','欧码科类代码')
欧码学校代码 = cp.get('考场设置','欧码学校代码')
欧码学生类别 = cp.get('考场设置','欧码学生类别')
#爱云校
云校学校代码 = cp.get('考场设置','云校学校代码')

#df排序
df = pd.read_excel("走班考试安排.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,'物理考场':str,'物理座号':str,"化学考场":str,"化学座号":str,"生物考场":str,"生物座号":str,"政治考场":str,"政治座号":str,"历史考场":str,"历史座号":str,"地理考场":str,"地理座号":str,})#导入走班考试安排,设置字符串

if 阅卷系统 == 1:#欧码系统
    #语数英考场导出
    df1=df.copy()
    df1.insert(9, "欧码科类代码", 欧码科类代码, allow_duplicates=False)
    df1.insert(10, "欧码学校代码", 欧码学校代码, allow_duplicates=False)
    df1.insert(11, "班级代码",  0, allow_duplicates=False)#插入
    df1['班级代码']= df1['班级名称']#赋值
    df1['备注'] = df1['身份证号']  # 备注写入身份证
    df1.insert(12, "欧码学生类别", 欧码学生类别, allow_duplicates=False)
    colNameDict = {'语数英考场':'考场号','语数英座号':'座号'} #df列重命名字典
    df1.rename(columns=colNameDict, inplace=True)
    df11=df1[['考生号','姓名','欧码科类代码','考场号','座号','欧码学校代码','学籍号','班级代码','班级名称','欧码学生类别','备注','身份证号']]#筛选需要的信息
    # 保存到相对目录
    if os.path.exists("导入阅卷系统文件") == False:
        os.mkdir("导入阅卷系统文件")
    lujing = os.getcwd() + r'/导入阅卷系统文件/'
    # lujing=os.path.abspath(os.path.dirname(os.path.abspath(__file__)))+ r'/导入阅卷系统文件/'#导出有误
    df11.to_excel(lujing  + '语数英.xls', sheet_name="语数英", index=False)  # 保存
    #df11.to_excel('语数英.xls', sheet_name="语数英",index=False) # 保存
    colNameDict = {'考场号':'语数英考场','座号':'语数英座号'} #恢复原设置
    df1.rename(columns=colNameDict, inplace=True)

    #6选3考场导出
    考试科目 = ["物理", "化学", "生物", "政治", "历史", "地理", ]
    for 学科 in 考试科目:
        #选出改科考试的学生DF2
        df2=df1.copy()
        df2 = df2.loc[df1[学科 + '考否'].str.contains('考试')]
        df2.sort_values(学科+"考场",)  # 按照考场排序
        colNameDict = {学科+'考场': '考场号', 学科+'座号': '座号'}  # 重命名字典
        df2.rename(columns=colNameDict, inplace=True)
        df2 = df2[['考生号', '姓名', '欧码科类代码', '考场号', '座号', '欧码学校代码', '学籍号', '班级代码', '班级名称', '欧码学生类别', '备注', '身份证号']]
        #保存到相对目录
        if os.path.exists("导入阅卷系统文件")==False:#判断文件是否创建
            os.mkdir("导入阅卷系统文件")
        lujing=os.path.dirname(__file__)+r'/导入阅卷系统文件/'#生成文件路径
        df2.to_excel(lujing+学科+'.xls', sheet_name=学科, index=False)  # 保存

if 阅卷系统 == 2:#2爱云校系统
    #语数英考场导出
    df1=df.copy()
    df1.insert(9, "学校代码", 云校学校代码, allow_duplicates=False)
    # df1.insert(10, "云校学校代码", 云校学校代码, allow_duplicates=False)
    # df1.insert(11, "班级代码",  0, allow_duplicates=False)#插入
    df1['班级代码']= df1['班级名称']#赋值
    df1['备注'] = df1['身份证号']  # 备注写入身份证号
    # df1.insert(12, "云校学生类别", 云校学生类别, allow_duplicates=False)
    colNameDict = {'语数英考场':'考场号','语数英座号':'座号'} #df列重命名字典
    df1.rename(columns=colNameDict, inplace=True)
    df11=df1[['考生号','姓名','考场号','座号','学校代码','班级代码','班级名称','备注',]]#筛选需要的信息
    # 保存到相对目录
    if os.path.exists("导入阅卷系统文件") == False:
        os.mkdir("导入阅卷系统文件")
    lujing= os.getcwd()+r'/导入阅卷系统文件/'
    # lujing = os.path.dirname(__file__) + r'/导入阅卷系统文件/'  # 生成文件路径
    df11.to_excel(lujing  + '语数英.xls', sheet_name="语数英", index=False)  # 保存
    colNameDict = {'考场号':'语数英考场','座号':'语数英座号'} #恢复原设置
    df1.rename(columns=colNameDict, inplace=True)

    #6选3考场导出
    考试科目 = ["物理", "化学", "生物", "政治", "历史", "地理", ]
    for 学科 in 考试科目:
        #选出改科考试的学生DF2
        df2=df1.copy()
        df2 = df2.loc[df1[学科 + '考否'].str.contains('考试')]
        df2.sort_values(学科+"考场",)  # 按照考场排序
        colNameDict = {学科+'考场': '考场号', 学科+'座号': '座号'}  # 重命名字典
        df2.rename(columns=colNameDict, inplace=True)
        df2 = df2[['考生号','姓名','考场号','座号','学校代码','班级代码','班级名称','备注',]]
        #保存到相对目录
        if os.path.exists("导入阅卷系统文件")==False:#判断文件是否创建
            os.mkdir("导入阅卷系统文件")
        lujing = os.getcwd() + r'/导入阅卷系统文件/'
        # lujing=os.path.dirname(__file__)+r'/导入阅卷系统文件/'#生成文件路径
        df2.to_excel(lujing+学科+'.xls', sheet_name=学科, index=False)  # 保存


# df=df[['班级名称','姓名','学籍号','学号','身份证号','组合','备注','名次',]]##数据删除
dt = datetime.now()
#dt= dt.strftime( '%Y-%m-%d %H:%M:%S %f' )
dt= dt.strftime( '%Y%m%d %H_%M' )
# wb.save(kao_name+dt+'.xlsx')
df.to_excel('走班考试安排备份'+dt+'.xls', sheet_name="走班考试安排", index=False)  # 保存
