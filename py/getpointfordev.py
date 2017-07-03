#coding:utf-8

import requests
import os
import re
import xlwt
import time
import datetime
from xlutils.copy import copy
from xlrd import open_workbook
import json
from collections import defaultdict
# import sys
# reload(sys)
# sys.setdefaultencoding('GBK')







# def getbugs():
#     #add cookie
#     cookies = {}
#     cookies[sid]='acg0moqbt9iht8b5guids4ev81'
#     f = xlwt.Workbook(encoding='utf-8')
#     sheet1 = f.add_sheet("sheet1",cell_overwrite_ok=True)
#     row0 = ['id','产品线','分支','项目','标题','严重度','责任人',  ,'由谁创建','创建日期','状态','解决人','解决方案','解决日期','激活次数','解决花费时间(小时)','验证花费时间(小时)','bug总周期(小时)']
#     for j in range(len(row0)):
#         sheet1.write(0,j,row0[j])

#     k = 1
#     for i in range(int(start_id),int(end_id)):
#         r = requests.get('http://chandao.com/zentao/bug-view-'+str(i)+'.html',cookies=cookies).content
#         # print r
#         if re.search("3.9.3",r):
#             id = re.findall("(?<=BUG #)[0-9]{4}",r)[0]
#             print id
#             # print re.findall("(?<=所属分支</th>).*?(?=</a>)",r,re.S)
#             # for App
#             if(re.findall("(?<=所属分支</th>).*?(?=</a>)",r,re.S) == []):
#                 continue
#             else:
#                 branch = re.findall("(?<=html' >).*",re.findall("(?<=所属分支</th>).*?(?=</a>)",r,re.S)[0])[0]
#                 # print re.findall("(?<=html' >).*",re.findall("(?<=所属分支</th>).*?(?=</a>)",r,re.S)[0])

#             # for IOS
#             # if (re.findall("(?<=所属分支</th>).*?(?=</a>)", r, re.S) == []):
#             #     continue
#             # else:
#             #     if (re.findall("(?<=html' >).*",re.findall("(?<=所属分支</th>).*?(?=</a>)",r,re.S)[0]) == ['iOS']):
#             #         branch = re.findall("(?<=html' >).*", re.findall("(?<=所属分支</th>).*?(?=</a>)", r, re.S)[0])[0]
#             #     else:
#             #         continue
#             title = re.findall("(?<=#[0-9]{4} ).*?(?=- )",r)[0]
#             responsible = re.findall("(?<=td>).*",re.findall("(?<=操作系统</th>).*?(?=</td>)",r,re.S)[0])[0]
#             tester = re.findall("(?<=td>).*",re.findall("(?<=由谁创建</th>).*?(?=于)",r,re.S)[0])[0]
#             create_time = re.findall("(?<=于 ).*",re.findall("(?<=由谁创建).*?(?=</td>)",r,re.S)[0])[0]
#             status = re.findall("(?<=strong>).*",re.findall("(?<=Bug状态</th>).*?(?=</strong>)",r,re.S)[0])[0]
#             if status == "已关闭":
#                 solver = re.findall("(?<=td>).*",re.findall("(?<=由谁解决</th>).*?(?=于)",r,re.S)[0])[0]
#                 resolution = re.findall("(?<=td>).*(?=</td>)", re.findall("(?<=解决方案</th>).*?(?<=</td>)", r, re.S)[0], re.S)[0].strip()
#                 print resolution
#                 resolved_time = re.findall("(?<=于 ).*",re.findall("(?<=由谁解决).*?(?=</td>)",r,re.S)[0])[0]
#                 opened_times = re.findall("(?<=td>).*",re.findall("(?<=激活次数).*?(?=</td>)",r,re.S)[0])[0]
#                 closed_time = re.findall("(?<=于 ).*",re.findall("(?<=由谁关闭).*?(?=</td>)",r,re.S)[0])[0]
#                 time_resolved = datetime.datetime.strptime(resolved_time,"%Y-%m-%d %H:%M:%S")
#                 time_create = datetime.datetime.strptime(create_time,"%Y-%m-%d %H:%M:%S")
#                 time_closed = datetime.datetime.strptime(closed_time,"%Y-%m-%d %H:%M:%S")
#                 if time_resolved > time_create:
#                     resolve_costs = int((time_resolved - time_create).seconds/3600)+1
#                 else:
#                     resolve_costs = "error"
#                 if time_closed > time_resolved:
#                     verify_costs = int((time_closed - time_resolved).seconds/3600)+1
#                 else:
#                     verify_costs = "error"
#                 if time_closed > time_create:
#                     bug_costs = int((time_closed - time_create).seconds/3600)+1
#                 else:
#                     bug_costs = "error"
#             if status == "已解决":
#                 solver = re.findall("(?<=td>).*",re.findall("(?<=由谁解决</th>).*?(?=于)",r,re.S)[0])[0]
#                 resolution = re.findall("(?<=td>).*(?=</td>)", re.findall("(?<=解决方案</th>).*?(?<=</td>)", r, re.S)[0], re.S)[0].strip()
#                 resolved_time = re.findall("(?<=于 ).*",re.findall("(?<=由谁解决).*?(?=</td>)",r,re.S)[0])[0]
#                 opened_times = re.findall("(?<=td>).*",re.findall("(?<=激活次数).*?(?=</td>)",r,re.S)[0])[0]
#                 closed_time = ""
#                 verify_costs = ""
#                 bug_costs = ""
#                 time_resolved = datetime.datetime.strptime(resolved_time,"%Y-%m-%d %H:%M:%S")
#                 time_create = datetime.datetime.strptime(create_time,"%Y-%m-%d %H:%M:%S")
#                 time_closed = ""
#                 if time_resolved > time_create:
#                     resolve_costs = int((time_resolved - time_create).seconds/3600)+1
#                 else:
#                     resolve_costs = "error"
#             if status == "激活":
#                 solver = ""
#                 resolution = ""
#                 # resolution = re.findall("(?<=td>).*", re.findall("(?<=解决方案</th>).*?(?=于)", r, re.S)[0])[0]
#                 print resolution
#                 resolved_time = ""
#                 opened_times = re.findall("(?<=td>).*",re.findall("(?<=激活次数).*?(?=</td>)",r,re.S)[0])[0]
#                 closed_time = ""
#                 verify_costs = ""
#                 bug_costs = ""
#                 time_closed = ""
#                 resolve_costs = ""


#             # else:
#             #     solver = "error"
#             #     resolve_time = "error"
#             #     opened_times = "error"
#             #     resolve_costs = "error"

#             # tester_data.append(tester)
#             # print type(id),type(title),type(responsible),type(tester)
#             print id,branch,title,responsible,tester,create_time,status,solver,resolution,resolved_time,opened_times,resolve_costs,verify_costs,bug_costs
#             row = [id,branch,title,responsible,tester,create_time,status,solver,resolution,resolved_time,opened_times,resolve_costs,verify_costs,bug_costs]
#             for j in range(len(row)):
#                 sheet1.write(k,j,row[j])
#                 # print f.sheet1[j]
#             k += 1

#     f.save("3.9.3demoiOS.xls")
# class dev(object) :
#     def __set(self, key, value):
#         setattr(self, key, value)
#     def __get(self, key):
#         return getattr(self, key)



# 研发  兼容性--研发 产品    设计缺陷--产品    测试  重复--测试  其他  其他
#       代码错误--研发       需求建议--产品        测试问题--测试        第三方sdk
#       界面错误--研发       需求错误--产品        无法重现        需求争议--争议
#       安全--研发           用户体验--产品        不是问题        调试阶段问题
#       性能--研发           需求变更--产品        重复问题        对接争议--争议
#       需求实现错误--研发   产品问题              测试问题        测试环境异常
#       已解决               需求变更                非研发中心问题
#       延期处理                        null
#       历史问题                        




def getbugspoint(serid,reoid,onlid):
    serious={1:2,2:1.2,3:1,4:0.7}
    reopen={1:1.1,2:1.5,3:2.5,4:4,0:1,5:10}
    #bugtype1={'兼容性--研发':1,'代码错误--研发':1,'界面错误--研发':1,'安全--研发':1,'性能--研发':1,'需求实现错误--研发':1}
    online={1:10,0:1}
    c=serious[serid]*reopen[reoid]*online[onlid]
    return c

def getdevpoint(exceltab):
    #初始化参数
    devbugs=[]
    probugs=[]
    testbugs=[]
    outsidebugs=[]
    devfixbugs=[]
    profixbugs=[]
    testfixbugs=[]
    outsidefixbugs=[]



    excelload='D:\\noo\\_内容\\rebuild.xlsx'
    excelsave='D:\\noo\\_内容\\save.xls'

    excelsavedev='devs'
    listsavedev=['姓名','总计','待处理','已修复','评分','历史问题数','平均bug关闭时长','关闭bug数量','平均bug解决时长','解决bug数量']
    excelsavetest='test'
    listsavetest=['姓名','总计','重复--测试','测试问题--测试','无法重现','不是问题','重复问题','测试问题','待处理问题','总提交bug']
    excelsavepro='pro'
    listsavepro=['项目','设计缺陷--产品','需求建议--产品','需求错误--产品','用户体验--产品','需求变更--产品','产品问题','需求变更','产品处理中','','研发汇总','测试汇总','产品总计','外部汇总','异常bug汇总','项目事项总计']
    excelsaveother='other'
    listsaveother=['bugid','bugcreater','bugsolver','bugtpye','','bugtpye','count']
    excelsetting='setting'
    excelsaveproject='count'
    excel=open_workbook(excelload)


    list_dev=defaultdict(dict)
    list_test=defaultdict(dict)
    list_pro=defaultdict(dict)   
    list_other=defaultdict(dict)   

    dev=['万其','尚小健','徐涛','刘骏','顾飞','林建有','程星','陈博鹏','刘涛','李杰','李江东','','邱计','解宝龙','姬武超','马冯鑫','吴文超','宋乃银','刘祖荣','张继文','','包杰','郭雄飞','刘太华','杨扬','丁庆','刘兴旺','吴正兵','','陈丽能','李永京','冯瑞','周奇','李诚杰','朱博','','谈开凯','李文灏','魏可森','张超','陆俊杰','沈兴山','蒋双喜','','王骏','陈险梅','赵钱龙','刘帅帅','王爽','单杰','张孟伟','徐惠芳','修亦池','吴照永','沈天忆','陈祥凯','马骥雄','','王宏玉','范其乐','韩武杰','程显玮','叶永发','谭东杰','','郭懿欣','陈涛','吴宇才','范伟','任伟栋','顾光益','薛闯','孙绍兵','黄昳彬','']

    test=['陈亮','付正国','叶春兰','史传倩','华岳琴','许小红','钟晓星','王吉祥','吴荣春','贺志平','荆莹','蒋震','陈凌霄','admin','王施民','李志荣','']

    pro=['管昊俊','董方','王晓芸','李梦麒','连宏坤','赵鸿','王林豪','唐豫','钦达莫尼','胡东','栗君','朱润','李煜','']

    projectall=[]

    #初始化第一行及用户列表 
    for i in range (len(listsavedev)):
        list_dev[0][i]=listsavedev[i]

    for i in range (len(listsavetest)):
        list_test[0][i]=listsavetest[i]

    for i in range (len(listsaveother)):
        list_other[0][i]=listsaveother[i]

    for i in range (len(listsavepro)):
        list_pro[0][i]=listsavepro[i]

    for j in range (len(dev)):
        for i in range (len(listsavedev)) :
            if (i==0):
                list_dev[j+1][i]=dev[j]
            else :
                list_dev[j+1][i]=0

    for j in range (len(test)):
        for i in range (len(listsavetest)) :
            if (i==0):
                list_test[j+1][i]=test[j]
            else :
                list_test[j+1][i]=0





    # for j in range (len(dev_api)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=dev_api[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[dev_api[j]]=nameline
    #     print ('set','|',nameline,'|',dev_api[j],'|row',nameline)

    # for j in range (len(dev_web)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=dev_web[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[dev_web[j]]=nameline
    #     print ('set','|',nameline,'|',dev_web[j],'|row',nameline)

    # for j in range (len(dev_yy)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=dev_yy[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[dev_yy[j]]=nameline
    #     print ('set','|',nameline,'|',dev_yy[j],'|row',nameline)
        
    # for j in range (len(dev_video)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=dev_video[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[dev_video[j]]=nameline
    #     print ('set','|',nameline,'|',dev_video[j],'|row',nameline)

    # for j in range (len(test)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=test[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[test[j]]=nameline
    #     print ('set','|',nameline,'|',test[j],'|row',nameline)

    # for j in range (len(pro)):
    #     nameline=nameline+1
    #     list_dev[nameline]['name']=pro[j]
    #     list_dev[nameline]['waitbug']=0
    #     list_dev[nameline]['fixbug']=0 
    #     list_dev[nameline]['point']=0
    #     namelist[pro[j]]=nameline
    #     print ('set','|',nameline,'|',pro[j],'|row',nameline)   

    #处理bug范围
    for i in range (15):
        if (excel.sheet_by_name(excelsetting).cell(i, 1).value!=''):
            devbugs.append(excel.sheet_by_name(excelsetting).cell(i, 1).value)
        if (excel.sheet_by_name(excelsetting).cell(i, 3).value!=''):   
            probugs.append(excel.sheet_by_name(excelsetting).cell(i, 3).value)    
        if (excel.sheet_by_name(excelsetting).cell(i, 5).value!=''):
            testbugs.append(excel.sheet_by_name(excelsetting).cell(i, 5).value)    
        if (excel.sheet_by_name(excelsetting).cell(i, 7).value!=''):
            outsidebugs.append(excel.sheet_by_name(excelsetting).cell(i, 7).value)  
    for i in range (excel.sheet_by_name(excelsetting).nrows-20):
        if (excel.sheet_by_name(excelsetting).cell(i+20, 1).value!=''):
            devfixbugs.append(excel.sheet_by_name(excelsetting).cell(i+20, 1).value)
        if (excel.sheet_by_name(excelsetting).cell(i+20, 3).value!=''):   
            profixbugs.append(excel.sheet_by_name(excelsetting).cell(i+20, 3).value)    
        if (excel.sheet_by_name(excelsetting).cell(i+20, 5).value!=''):
            testfixbugs.append(excel.sheet_by_name(excelsetting).cell(i+20, 5).value)    
        if (excel.sheet_by_name(excelsetting).cell(i+20, 7).value!=''):
            outsidefixbugs.append(excel.sheet_by_name(excelsetting).cell(i+20, 7).value)  


    for j in range (len(outsidebugs)):
        indexX=j+2
        list_other[indexX][5]=outsidebugs[j]    
        list_other[indexX][6]=0
    for j in range (len(outsidefixbugs)):
        indexX=j+2+len(outsidebugs)
        list_other[indexX][5]=outsidefixbugs[j]    
        list_other[indexX][6]=0

    list_other[1][5]='error'
    list_other[1][6]=0
    #print ('devbugs=>>',devbugs)
    #print ("probugs=>>",probugs)
    #print ("testbugs=>>",testbugs)
    #print ("outsidebugs=>>",outsidebugs)
    #处理数据
    for i in range(excel.sheet_by_name(exceltab).nrows-1):
        print ('-------处理第',i,'条事项中------')
        bugid=excel.sheet_by_name(exceltab).cell(i+1, 0).value#bug编号
        productline=excel.sheet_by_name(exceltab).cell(i+1, 1).value#所属产品
        project=excel.sheet_by_name(exceltab).cell(i+1, 4).value#所属项目
        bugtitle=excel.sheet_by_name(exceltab).cell(i+1, 7).value#bug标题
        bugserious=excel.sheet_by_name(exceltab).cell(i+1, 9).value#严重程度
        bugtpye=excel.sheet_by_name(exceltab).cell(i+1, 11).value#bug类型
        bugsummary=excel.sheet_by_name(exceltab).cell(i+1, 14).value#重现步骤
        bugstatus=excel.sheet_by_name(exceltab).cell(i+1, 15).value#bug状态
        bugreopentime=excel.sheet_by_name(exceltab).cell(i+1, 16).value#激活次数
        bugcreater=excel.sheet_by_name(exceltab).cell(i+1, 19).value#由谁创建
        bugcreatedate=excel.sheet_by_name(exceltab).cell(i+1, 20).value#创建日期
        bugassign=excel.sheet_by_name(exceltab).cell(i+1, 22).value#指派给
        bugassigndate=excel.sheet_by_name(exceltab).cell(i+1, 23).value#指派日期
        bugsolver=excel.sheet_by_name(exceltab).cell(i+1, 24).value#解决者
        bugsolution=excel.sheet_by_name(exceltab).cell(i+1, 25).value#解决方案
        bugsolvedate=excel.sheet_by_name(exceltab).cell(i+1, 27).value#解决日期
        bugcloser=excel.sheet_by_name(exceltab).cell(i+1, 28).value#关闭者
        bugclosedate=excel.sheet_by_name(exceltab).cell(i+1, 29).value#关闭日期  
        if (bugclosedate!='0000-00-00'):   
            bugclosetime=bugclosedate-bugcreatedate#关闭日长
        else:
            bugclosetime=-1
        if (bugsolvedate!='0000-00-00'):                                           
            bugsolvetime=bugsolvedate-bugcreatedate#解决日长
        else:
            bugsolvetime=-1


        if (project not in projectall):
            projectall.append(project)
            for i in range (len(listsavepro)):
                if (i==0):
                    list_pro[len(projectall)][i]=project
                else:
                    list_pro[len(projectall)][i]=0


        #处理未加入列表用户
        # if (bugsolver!='' and not namelist.get(bugsolver,False)):
        #     nameline=nameline+1 
        #     list_dev[nameline]['name']=bugsolver
        #     list_dev[nameline]['waitbug']=0
        #     list_dev[nameline]['fixbug']=0 
        #     list_dev[nameline]['point']=0
        #     namelist[bugsolver]=nameline
        #     print ('set','|',nameline,'|',bugsolver,'|row',i)
        # elif (bugassign!=''and not namelist.get(bugassign,False)):
        #     nameline=nameline+1
        #     namelist[bugassign]=nameline
        #     list_dev[nameline]['name']=bugassign
        #     list_dev[nameline]['waitbug']=0
        #     list_dev[nameline]['fixbug']=0 
        #     list_dev[nameline]['point']=0
        #     namelist[bugassign]=nameline
        #     print ('set','|',nameline,'|',bugassign,'|row',i)
        #初始化第一行

        # sheet.write(0,0,'人员')     #_name
        # sheet.write(0,1,'总计bug名称数量')  #_countbug
        # sheet.write(0,2,'待处理bug')    #_waitbug
        # sheet.write(0,3,'已处理bug')    #_fixbug
        # sheet.write(0,4,'总积分')       #_point

        #处理bug分类---其他

        if (bugtpye in outsidebugs):
            print ('add to ==>>>未处理的outside') 
            indexX=len(list_other)
            list_other[indexX][0]=bugid
            list_other[indexX][1]=bugcreater
            list_other[indexX][2]=bugsolver
            list_other[indexX][3]=bugtpye

            indexX=projectall.index(project)+1 
            list_pro[indexX][13]+=1          

            indexY=outsidebugs.index(bugtpye)+2
            list_other[indexY][6]+=1           


        #处理bug分类---fix的其他
        elif(bugsolution in outsidefixbugs):
            print ('add to ==>>>fix的outside')            
            indexX=len(list_other)
            list_other[indexX][0]=bugid
            list_other[indexX][1]=bugcreater
            list_other[indexX][2]=bugsolver
            list_other[indexX][3]=bugsolution

            indexX=projectall.index(project)+1 
            list_pro[indexX][13]+=1        

            indexY=outsidefixbugs.index(bugsolution)+2+len(outsidebugs)
            list_other[indexY][6]+=1   

        #处理研发问题 -- 解决者
        elif ((bugsolver in dev) and (bugsolution in devfixbugs)and (bugtpye in devbugs)):
            print ('add to ==>>>解决者')            
            indexX=dev.index(bugsolver)+1 
            list_dev[indexX][3]+=1
            list_dev[indexX][4]+=getbugspoint(bugserious,bugreopentime,0) 
            if (bugsolution=='历史问题'):
                list_dev[indexX][5]+=1
            #研发关闭日长
            if (bugclosetime!=-1):
                list_dev[indexX][6]+=bugclosetime
                list_dev[indexX][7]+=1
            #研发解决日长
            if (bugsolvetime!=-1):
                list_dev[indexX][8]+=bugsolvetime
                list_dev[indexX][9]+=1


            indexX=projectall.index(project)+1 
            list_pro[indexX][10]+=1  

        #处理研发问题 -- 指派给
        elif (bugsolver=='' and (bugassign in dev) and (bugtpye in devbugs)):
            print ('add to ==>>>指派者')
            indexX=dev.index(bugassign)+1 
            list_dev[indexX][2]+=1
            list_dev[indexX][4]+=getbugspoint(bugserious,bugreopentime,0)
            #研发解决日长
            if (bugsolvetime!=-1):
                list_dev[indexX][8]+=bugsolvetime
                list_dev[indexX][9]+=1

            indexX=projectall.index(project)+1 
            list_pro[indexX][10]+=1  

        #处理测试问题 -- 指派给
        elif ((bugsolution in testfixbugs) and (bugcreater in test)):
            print ('add to ==>>>test指派')
            indexX=test.index(bugcreater)+1 
            indexY=testfixbugs.index(bugsolution)+4
            list_test[indexX][indexY]+=1
            list_test[indexX][1]+=1

            indexX=projectall.index(project)+1 
            list_pro[indexX][11]+=1  

        #处理测试问题 -- 分类
        elif ((bugtpye in testbugs)and (bugcreater in test)):
            print ('add to ==>>>test提交')
            indexX=test.index(bugcreater)+1 
            indexY=testbugs.index(bugtpye)+1
            list_test[indexX][indexY]+=1
            list_test[indexX][1]+=1

            indexX=projectall.index(project)+1 
            list_pro[indexX][11]+=1  

        #处理产品问题 -- 解决方案
        elif ((bugsolution in profixbugs) ):
            print ('add to ==>>>pro解决方案')
            indexX=projectall.index(project)+1 
            indexY=profixbugs.index(bugsolution)+5
            list_pro[indexX][indexY]+=1
            list_pro[indexX][12]+=1
        #处理产品问题 -- 分类
        elif ((bugtpye in probugs)):
            print ('add to ==>>>pro分类')
            indexX=projectall.index(project)+1 
            indexY=probugs.index(bugtpye)+1
            list_pro[indexX][indexY]+=1
            list_pro[indexX][12]+=1
        #处理产品问题 -- 处理中
        elif ((bugassign in pro)):
            print ('add to ==>>>pro处理中')
            indexX=projectall.index(project)+1 
            list_pro[indexX][8]+=1
            list_pro[indexX][12]+=1        


        else :
            print ('------error bugs --------')            
            indexX=len(list_other)
            list_other[indexX][0]=bugid
            list_other[indexX][1]=bugcreater
            list_other[indexX][2]=bugsolver
            list_other[indexX][3]='error'

            list_other[1][6]+=1

        #测试待处理问题
        if (bugassign in test and bugcloser ==''):
            indexX=test.index(bugcreater)+1 
            list_test[indexX][8]+=1         
        #测试提交bug数
        if (bugcreater in test):
            indexX=test.index(bugcreater)+1 
            list_test[indexX][9]+=1    

        #项目bug汇总
        indexX=projectall.index(project)+1 
        list_pro[indexX][15]+=1            


    #保存dev数据
    a=xlwt.Workbook(encoding='utf-8')
    sheet=a.add_sheet(excelsavedev,cell_overwrite_ok=True)
    for i in range (len(dev)):
        if (i >0):
            list_dev[i][1]=list_dev[i][2]+list_dev[i][3]
            if ((list_dev[i][7])>0):
                list_dev[i][6]=list_dev[i][6]/list_dev[i][7]   
            if ((list_dev[i][9])>0):
                list_dev[i][8]=list_dev[i][8]/list_dev[i][9]               
        for j in range (len(listsavedev)):
            sheet.write(i,j,list_dev[i][j])


    #保存test数据
    sheet1=a.add_sheet(excelsavetest,cell_overwrite_ok=True)
    for i in range (len(list_test)):
        for j in range (len(listsavetest)):
            try:
                sheet1.write(i,j,list_test[i][j])
            except:
                thy=1

    #保存pro数据
    sheet1=a.add_sheet(excelsavepro,cell_overwrite_ok=True)
    for i in range (len(list_pro)):
        for j in range (len(listsavepro)):
            if (i>0):
                list_pro[i][14]=list_pro[i][15]-list_pro[i][13]-list_pro[i][12]-list_pro[i][11]-list_pro[i][10]
            sheet1.write(i,j,list_pro[i][j])
            # except:
            #     thy=1



    #保存other数据
    sheet1=a.add_sheet(excelsaveother,cell_overwrite_ok=True)
    for i in range (len(list_other)):
        for j in range (len(listsaveother)):
            try:
                sheet1.write(i,j,list_other[i][j])
            except:
                thy=1

    a.save(excelsave)

a=input ('请输入需要处理的bug tab：')
getdevpoint(str(a))
print ('----------Mission done!----------') 
