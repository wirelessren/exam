#!/usr/bin/python3.7
# -*- coding: utf-8 -*-
import pandas as pd
import configparser
import os
import math

#pyinstaller -F -w -i 3.ico out_for_print2.py
#pyinstaller -D -w -i 3.ico out_for_print2.py
#配置文件读取
filepath = os.path.join(os.getcwd(),'config.ini')
cp = configparser.ConfigParser()
cp.read(filepath, encoding="utf-8-sig")

#从配置文件读取相关设置
考试名称 = cp.get('系统设置','考试名称')
kc_rs = int(cp.get('辅助设置','考场人数'))
阅卷系统 = int(cp.get('系统设置','阅卷系统'))

#df排序
df = pd.read_excel("走班考试安排.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,"语数英考座":str,'物理考场':str,'物理座号':str,"物理考座":str,"化学考场":str,"化学座号":str,"化学考座":str,"生物考场":str,"生物座号":str,"生物考座":str,"政治考场":str,"政治座号":str,"政治考座":str,"历史考场":str,"历史座号":str,"历史考座":str,"地理考场":str,"地理座号":str,"地理考座":str,})#导入走班考试安排,设置字符串
df.sort_values("考生号",inplace=True)#按照名次排序
df.reset_index(inplace = True,drop = True)#重置索引

# 保存到相对目录
if os.path.exists("打印文件") == False:
    os.mkdir("打印文件")
lujing = os.getcwd() + r'/打印文件/'

# lujing = os.path.dirname(__file__) + r'/打印文件/'  # 生成一个文件路径

###1班级打印
#获取班级列表
df1=df.copy()
z=list(set(df1['班级名称']))#所有考场列表
z.sort()#考场排序
# z=dict(enumerate(z))#考场转换为字典

#筛选班级
#班级打印总表
df1= df1[['班级名称', '姓名', '考生号', '语数英考场', '语数英座号', '物理考场', '物理座号', '化学考场', '化学座号', '生物考场', '生物座号', '政治考场', '政治座号', '历史考场','历史座号', '地理考场', '地理座号']]  # 筛选需要的信息
df1.sort_values("班级名称",inplace=True)#按照名次排序
df1.to_excel(lujing+"班级打印.xlsx", index=False, sheet_name="班级打印总表")#保存到文件夹，备用追加数据

writer = pd.ExcelWriter(lujing+'班级打印.xlsx', mode='a',engine="openpyxl")#a=append#writer独立
for v in z:#追加内容到工作簿，文件夹必须有xlsx文件
    df1=df.copy()
    df1=df1.loc[df1['班级名称'] == v]
    df1 = df1[['班级名称', '姓名', '考生号', '语数英考场','语数英座号','物理考场','物理座号','化学考场','化学座号','生物考场','生物座号','政治考场','政治座号','历史考场','历史座号','地理考场','地理座号']]  # 筛选需要的信息
    # writer = pd.ExcelWriter(lujing+'班级打印.xls', mode='a',engine="openpyxl")#a=append
    df1.to_excel(writer,str(v),index=False)
    del df1
writer.save()
writer.close()

## 考场打印
#生成所有考座

out_system=阅卷系统 #1为云校，2为教科院
if out_system == 1:
    ks_num = cp.get('考场设置', '欧码考生号')
    kc_num = cp.get('考场设置', '欧码考场号')
if out_system == 2:
    ks_num = cp.get('考场设置', '云校考生号')
    kc_num = cp.get('考场设置', '云校考场号')
df1=df.copy()
all_ksrs = len(df1)
kc_sl=math.ceil((all_ksrs/kc_rs)+1)
kczuohao=list(range(1,kc_sl+1))#生成考场座号列表
kc_sl列表=[]
for i in kczuohao:#两位补充
    i=kc_num+str(i).zfill(2)
    kc_sl列表.append(i)
所有考场号=kc_sl列表*kc_rs
所有考场号.sort(reverse = False)
#
单个考场座号容量=list(range(1,kc_rs+1))
kc_rs=[]
for i in 单个考场座号容量:
    i=str(i).zfill(2)
    kc_rs.append(i)
所有座号=kc_rs*kc_sl

df21 = pd.DataFrame()#新建空白df
df21['所有考场号'] = 所有考场号
df21['所有座号']=所有座号
df21['考座']=df21['所有考场号']+df21['所有座号']
df55=df21.copy()#为自习考场删选生成df备用
#
# ##生成各科df
ks_xk=['语数英','物理','化学','生物','政治','历史','地理',]
for i in ks_xk:
    df22=df1[[i+'考座','班级名称','姓名','考生号']]
    df22.columns = ['考座',i+'班级名称',i+'姓名','考生号']
    df22.sort_values(by='考座')
    df21=pd.merge(df21,df22,on='考座',how='outer')
    del df22

del df21['考座']#删除辅助考座
df21.columns = ['考场','座号','语数英班级','语数英姓名','考生号',
                '物理班级','物理姓名','考生号',
                '化学班级','化学姓名','考生号',
                '生物班级','生物姓名','考生号',
                '政治班级','政治姓名','考生号',
                '历史班级','历史姓名','考生号',
                '地理班级','地理姓名','考生号',]#更改行名称
#
#分裂表格按照考场输出
# all_kc=list(set(df21['考场']))#所有考场列表
# all_kc.sort()#考场排序
# z=dict(enumerate(all_kc))#考场转换为字典
all_kc = df21['考场'].unique()
all_kc.sort()#考场排序

#筛选班级
df21.to_excel(lujing+"考场打印.xlsx", index=False, sheet_name="考场打印总表")#总表写入
# writer = pd.ExcelWriter(lujing + '考场打印.xls',)
writer = pd.ExcelWriter(lujing+'考场打印.xlsx', mode='a',engine="openpyxl")
for v in all_kc:##追加内容到总表
    df22=df21.copy
    df22 = df21[df21['考场'] == v]
    df22.to_excel(writer,str(v),index=False)
writer.save()
writer.close()
del df22
del df21

##自习打印
df5=df.copy()
zx_km=["物理","化学","生物","政治","历史","地理",]
for zxxk in zx_km:
    df51=df5.loc[df5[zxxk+'考否'].str.contains('自习')]
    df51 = df51[[zxxk + '考座', '班级名称', '姓名', ]]
    df51.columns = ['考座', zxxk + '班级名称', zxxk + '姓名', ]
    df51.sort_values(by='考座')
    df55 = pd.merge(df55, df51, on='考座', how='outer')
del df55['考座']  # 删除辅助考座

df55.columns = ['考场','座号','物理班级', '物理姓名', '化学班级', '化学姓名', '生物班级', '生物姓名','政治班级', '政治姓名', '历史班级', '历史姓名', '地理班级', '地理姓名',]  # 更改行名称
df55.to_excel(lujing + '自习打印.xlsx', sheet_name="自习打印总表", index=False)  # 保存

writer = pd.ExcelWriter(lujing + '自习打印.xlsx', mode='a', engine="openpyxl")
# writer = pd.ExcelWriter(lujing + '自习打印.xlsx', mode='a', engine="xlsxwriter")
for v in all_kc:  # 追加内容到工作簿，文件夹必须有xls文件
    ss=df55  #继承上部计算结果
    ss = df55.loc[df55['考场'] == v]
    ss.to_excel(writer, str(v), index=False)
writer.save()
writer.close()
