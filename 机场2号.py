'''通过对进度及分级信息的读取、分析，
实现提示当月卫生监督任务、具体单位卫生监督情况查询等功能
'''
'''其中cur_mm表示当前月份是2018年12月后的第几个月，
如当前月份为2019年5月，cur_mm即为5；
如为2020年3月，则对cur_mm的赋值需+12，2021年则需+24，以此类推
'''
import os,openpyxl,datetime,time

#读取进度表及分级情况
form_nm=openpyxl.load_workbook('00卫生监督、病媒生物进度V6.0.xlsx',data_only=True)
sht1=form_nm['卫生监督进度表']

#将co作为列表，co的每一项嵌套一个子列表，一个子列表为卫生监督进度表的一个单位的条目
r=3
i=0
co=[]
while sht1.cell(r,1).value != None:
  co.append([])
  co[i].append(sht1.cell(r,1).value)
  co[i].append(sht1.cell(r,2).value)
  co[i].append(sht1.cell(r,3).value)
  for l in range(0,12):
    co[i].append(sht1.cell(r,(4+l*7)).value)
  r+=1
  i+=1

#遍历每个单位，根据评级填好‘-’
for co_i in range(0,len(co)):
  cur_mm=int(time.strftime("%m",time.localtime()))
  if co[co_i][2]=='A级':
    while co[co_i][cur_mm+2]!='√'and type(co[co_i][cur_mm+2])!='int':
      cur_mm-=1
    finish_mm=cur_mm
    for i in range(1,6):
      sht1.cell(row=(co_i+3),column=((finish_mm+i-1)*7+4)).value='-'
  elif co[co_i][2]=='B级':
    while co[co_i][cur_mm+2]!='√'and type(co[co_i][cur_mm+2])!='int':
      cur_mm-=1
    finish_mm=cur_mm
    for i in range(1,3):
      sht1.cell(row=(co_i+3),column=((finish_mm+i-1)*7+4)).value='-'
  elif co[co_i][2]=='未定级':
    while co[co_i][cur_mm+2]!='√'and type(co[co_i][cur_mm+2])!='int':
      cur_mm-=1
      if cur_mm==0:
        cur_mm_for_change=int(time.strftime("%m",time.localtime()))
        sht1.cell(row=(co_i+3),column=((cur_mm_for_change-1)*7-3)).value='-'
        sht1.cell(row=(co_i+3),column=((cur_mm_for_change-2)*7-3)).value='√'
        break
    finish_mm=cur_mm
    for i in range(1,2):
      sht1.cell(row=(co_i+3),column=((finish_mm+i-1)*7+4)).value='-'

#将co作为列表，co的每一项嵌套一个子列表，一个子列表为卫生监督进度表的一个单位的条目
#重新读取加入‘-’后的数据以便打印
r=3
i=0
co=[]
while sht1.cell(r,1).value != None:
  co.append([])
  co[i].append(sht1.cell(r,1).value)
  co[i].append(sht1.cell(r,2).value)
  co[i].append(sht1.cell(r,3).value)
  for l in range(0,12):
    co[i].append(sht1.cell(r,(4+l*7)).value)
  r+=1
  i+=1

#从每一个子列表中读取当前月份的记录是否为√，如是，则说明本月已做卫生监督，如否，则向前逐月读取，直至读到√，并确定月份
co_i_todo=[]#未做
co_i_nottodo=[]#无需做
co_i_finish=[]#已完成
for co_i in range(0,len(co)):
  cur_mm=int(time.strftime("%m",time.localtime()))
  if co[co_i][cur_mm+2]=='√'or type(co[co_i][cur_mm+2])=='int':
    co_i_finish.append(co_i)
  elif co[co_i][cur_mm+2]=='-':
    co_i_nottodo.append(co_i)
  elif co[co_i][cur_mm+2]==None:
    co_i_todo.append(co_i)

print('本月卫生监督情况如下：')
print('不必监管：')
for i in co_i_nottodo:
  print(co[i][0])
print('--------------')
print('本月已完成：')
for i in co_i_finish:
  print(co[i][0])
print('--------------')
print('本月需监管：')
for i in co_i_todo:
  print(co[i][0])

form_nm.save('卫生监督进度V6.0.xlsx')
input("按回车键退出")
