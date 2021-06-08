#!/usr/bin/python3.7
# -*- coding: utf-8 -*-
import pandas as pd
import random
import math
import configparser
import os
#pyinstaller -F -w -i 1.ico exam_system.py
#pyinstaller -D -w -i 1.ico exam_system.py 多文件
#df排序
# df1 = pd.read_excel("走班考试安排.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,'物理考场':str,'物理座号':str,"化学考场":str,"化学座号":str,"生物考场":str,"生物座号":str,"政治考场":str,"政治座号":str,"历史考场":str,"历史座号":str,"地理考场":str,"地理座号":str,})#导入走班考试安排,设置字符串
df = pd.read_excel("走班考试安排.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,})#导入走班考试安排,设置字符串
df1=df.copy()
df1.sort_values("名次",inplace=True)#按照名次排序
df1.reset_index(inplace = True,drop = True)#重置索引
#插入列名称
if '分层数值' not in df1.columns:
    input_cl_num=(df1.shape[1])
    # y=["分层数值","考生号",'语数英考场','语数英座号','物理考否','物理考场','物理座号','化学考否','化学考场','化学座号','生物考否','生物考场','生物座号','政治考否','政治考场','政治座号','历史考否','历史考场','历史座号','地理考否','地理考场','地理座号']
    y=["分层数值","考生号",'语数英考场','语数英座号','语数英考座','物理考否','物理考场','物理座号','物理考座','化学考否','化学考场','化学座号','化学考座','生物考否','生物考场','生物座号','生物考座','政治考否','政治考场','政治座号','政治考座','历史考否','历史考场','历史座号','历史考座','地理考否','地理考场','地理座号','地理考座',]

    s=list(range(input_cl_num, input_cl_num+len(y)))
    z = dict(zip(s,y[:len(y)]))
    for k, v in z.items():
        # print(type(k))
        # print(type(v))
        df1.insert(k, str(v), '', )
#考试信息统计
df4 = pd.DataFrame({
    '项目':['考试人数','考场数量','结束考场人数','自习人数','自习开始考场号','自习结束考场号',],
    '语数英':[0,0,0,0,0,0,],
    '物理':[0,0,0,0,0,0,],
    '化学': [0, 0, 0, 0, 0, 0, ],
    '生物': [0, 0, 0, 0, 0, 0, ],
    '历史': [0, 0, 0, 0, 0, 0, ],
    '地理': [0, 0, 0, 0, 0, 0, ],
})

#配置外置文件读取
filepath = os.path.join(os.getcwd(),'config.ini')
cp = configparser.ConfigParser()
cp.read(filepath, encoding="utf-8-sig")


# 设置区
总人数 = len(df1)
分层数目 = int(cp.get('辅助设置','分层数值'))
# 分层数目 = 15
每层人数 = int(总人数 / 分层数目)
混编标记分层值 = (list(range(1, 总人数, 每层人数)))
混编标记分层值[-1]=总人数  #最后一个键值改为总人数
混编分层数目列表 = (list(range(1, int(分层数目) + 1, 1)))
混编辅助数值 = (list(range(1, int(总人数 + 1), 1)))
混编分层字典 = dict(zip(混编分层数目列表,混编标记分层值[1:分层数目 + 1]))
考场人数 = int(cp.get('辅助设置','考场人数'))
阅卷系统 = int(cp.get('系统设置','阅卷系统'))


#生成分层值+随机数值 写入分层数值并排序
混编分层数值=[]
for i in range(1, 总人数+1):
    for k, v in 混编分层字典.items():        # 前提是新版的py，字典有序
        if i <= v:
            sj = random.random()
            混编分层数值.append(k+sj)
            break
df1["分层数值"] = 混编分层数值  #写入分层数值
df1.sort_values("分层数值",inplace=True)#按照名次排序

##########################################考生号


out_system=阅卷系统 #1为云校，2为教科院
if out_system == 1:
    ks_num = cp.get('考场设置', '欧码考生号')
    kc_num = cp.get('考场设置', '欧码考场号')
if out_system == 2:
    ks_num = cp.get('考场设置', '云校考生号')
    kc_num = cp.get('考场设置', '云校考场号')

考生号 = []

for i in 混编辅助数值:
    i=str(i).zfill(4)
    考生号1=str(ks_num+i)
    考生号.append(考生号1)
df1["考生号"] = 考生号  # 写入分层数值
#########################################################语数英座号
# 生成单个考场座号

单个考场座号容量=list(range(1,考场人数+1))
考场人数=[]
for i in 单个考场座号容量:
    i=str(i).zfill(2)
    考场人数.append(i)
#先生成所需的所有考场座号，取列表中考试人数的座号
考场数量=math.ceil(总人数/len(考场人数))    #计算考场数量，上入
生成所有考场座号=list(考场人数*考场数量)
考试座号=生成所有考场座号[0:总人数]
df1["语数英座号"] = 考试座号  # 写入分层数值



################################语数英考场号
kczuohao=list(range(1,考场数量+1))#生成考场座号列表
考场数量列表=[]
for i in kczuohao:#两位补充
    i=kc_num+str(i).zfill(2)
    考场数量列表.append(i)

生成所有考场考场号=list(考场数量列表*len(考场人数))#生成所有考场号

生成所有考场考场号.sort()#考场号列表排序
考场号=生成所有考场考场号[0:总人数]#取考试人数相符合的考场号
df1["语数英考场"] = 考场号  # 写入考场号
df1["语数英考座"] = df1["语数英考场"]+df1["语数英座号"]#######################合成考座
##############################################考否
# 判断考试还是自习
选科组合=df1["组合"]
考试科目=["物理","化学","生物","政治","历史","地理",]
# 判断是否考试
for 学科 in 考试科目:
    考否数列 = []
    for i in 选科组合:
        if 学科 in i:
            考否数列.append("考试")
        else:
            考否数列.append("自习")
    df1[(学科 + "考否")] = 考否数列

###########################################
#考试和自习筛选

#裂表
考试科目=["物理","化学","生物","政治","历史","地理",]
# 考试科目=["生物"]
for 学科 in 考试科目:
    df2=df1.loc[df1[学科+'考否'].str.contains('考试')]
    df3=df1.loc[df1[学科+'考否'].str.contains('自习')]
    df2.to_csv('考试.csv', index=False)  # 保存
    df3.to_csv('自习.csv', index=False)  # 保存
    #考试考场号
    # df2 = pd.read_excel("考试.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,"语数英考座":str,'物理考场':str,'物理座号':str,"物理考座":str,"化学考场":str,"化学座号":str,"化学考座":str,"生物考场":str,"生物座号":str,"生物考座":str,"政治考场":str,"政治座号":str,"政治考座":str,"历史考场":str,"历史座号":str,"历史考座":str,"地理考场":str,"地理座号":str,"地理考座":str,})#导入走班考试安排,设置字符串
    df2 = pd.read_csv("考试.csv",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,"语数英考座":str,'物理考场':str,'物理座号':str,"物理考座":str,"化学考场":str,"化学座号":str,"化学考座":str,"生物考场":str,"生物座号":str,"生物考座":str,"政治考场":str,"政治座号":str,"政治考座":str,"历史考场":str,"历史座号":str,"历史考座":str,"地理考场":str,"地理座号":str,"地理考座":str,})#导入走班考试安排,设置字符串
    df2.sort_values("考生号")#按照名次排序
    六选三考试人数=(len(df2))
    df2[学科+'考场']=生成所有考场考场号[0:六选三考试人数]#前边声明该变量
    #座号
    df2[学科+'座号']=生成所有考场座号[0:六选三考试人数]#座号
    #考试考座合成
    df2[学科+'考座']=df2[学科+'考场']+df2[学科+'座号']

    # 自习考场号
    # df3 = pd.read_excel("自习.xls",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,"语数英考座":str,'物理考场':str,'物理座号':str,"物理考座":str,"化学考场":str,"化学座号":str,"化学考座":str,"生物考场":str,"生物座号":str,"生物考座":str,"政治考场":str,"政治座号":str,"政治考座":str,"历史考场":str,"历史座号":str,"历史考座":str,"地理考场":str,"地理座号":str,"地理考座":str,})#导入走班考试安排,设置字符串
    df3 = pd.read_csv("自习.csv",converters={"学籍号":str,"学号":str,"身份证号":str,"考生号":str,"语数英考场":str,"语数英座号":str,"语数英考座":str,'物理考场':str,'物理座号':str,"物理考座":str,"化学考场":str,"化学座号":str,"化学考座":str,"生物考场":str,"生物座号":str,"生物考座":str,"政治考场":str,"政治座号":str,"政治考座":str,"历史考场":str,"历史座号":str,"历史考座":str,"地理考场":str,"地理座号":str,"地理考座":str,})#导入走班考试安排,设置字符串
    df3.sort_values("班级名称",inplace=True)#按照名次排序
    已用考场=math.ceil(len(df2)/len(考场人数))#统计考试已用考场
    自习考场号列表 = list(range(已用考场 + 1, 考场数量 + 2))
    自习考场号 = []
    for i in 自习考场号列表:
        i = kc_num + str(i).zfill(2)
        自习考场号.append(i)
    所有自习考场号=自习考场号*len(考场人数)
    所有自习考场号.sort()#考场号列表排序
    自习人数 = (len(df3))
    考场号=所有自习考场号[0:自习人数]#取考试人数相符合的考场号

    df3[学科+'考场']=考场号

    #座号
    df3[学科+'座号']=生成所有考场座号[0:自习人数]
    #自习考座合成
    df3[学科 + '考座']=df3[学科+'考场']+df3[学科+'座号']
    df3.to_csv('自习.csv', index=False)  # 保存
    df1=pd.concat([df2,df3])
    df1.sort_values("考生号",inplace=True)#按照名次排序

    os.remove('考试.csv')
    os.remove('自习.csv')

    df1.to_excel('走班考试安排.xls', sheet_name='走班考试安排',index=False)  # 保存

#输出考试信息
#1语数英信息输出
#考场人数
df4.loc[0, '语数英'] = 总人数 #考试人数
#考场数量
语数英考场数 =len(set(df1['语数英' + '考场']))
df4.loc[1, '语数英'] = 语数英考场数
#结束考场人数
语数英尾考场=list(df1['语数英' + '座号'])
df4.loc[2, '语数英'] = int(语数英尾考场[-1])
#2六选三信息输出

for xk in 考试科目:
    #6选3考试人数
    df2 = df1.loc[df1[xk + '考否'].str.contains('考试')] #提取考试信息
    ksrs=len(df2)#统计考试人数
    df4.loc[0, xk] = ksrs #写入考试人数
    #考试考场数量
    xk_kc_num = len(set(df2[xk + '考场']))
    df4.loc[1, xk] = xk_kc_num
    #结束考场人数
    # 语数英尾考场 = list(df2[xk + '座号'])
    选考尾考场=list(df2[xk+'座号'])
    # 尾考场人数整理=int(选考尾考场)+'人'
    df4.loc[2, xk] = int(选考尾考场[-1])

    #6选3自习人数
    df2 = df1.loc[df1[xk + '考否'].str.contains('自习')] #提取考试信息
    ksrs=len(df2)#统计考试人数
    df4.loc[3, xk] = ksrs #写入考试人数
    #自习开始考场
    zx_kc_num = list(df2[xk + '考场'])
    zx_kc_num.sort()
    df4.loc[4, xk] = zx_kc_num[0]
    #自习结束考场
    df4.loc[5, xk] = zx_kc_num[-1]
df4.to_excel('考试信息统计.xls',sheet_name='考试信息统计',index=False)
