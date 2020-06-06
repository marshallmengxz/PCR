import h5py
import csv
import numpy as np
from PIL import Image
import pytesseract
import xlrd
import openpyxl
from xlutils.copy import copy
def bossID(bossname):
    if bossname=='双足飞龙':
        return(1)
    elif bossname=='野性狮鹫':
        return(2)
    elif bossname=='兽人头目':
        return(3)
    elif bossname=='灵魂角鹿':
        return(4)
    elif bossname=='弥诺陶洛斯':
        return (5)
pytesseract.pytesseract.tesseract_cmd = r'/usr/local/Cellar/tesseract/4.1.1/bin/tesseract'
pytesseract.pytesseract.tessdata_cmd = r'/usr/local/Cellar/tesseract/4.1.1/share/tessdata'
imageObject=Image.open('/Users/mengzixia/Desktop/9.jpg')
savePath='/Users/mengzixia/Downloads/工会战报刀表.xlsx'
date=2
present_round=1
#读取excel文件
myexcel=xlrd.open_workbook(savePath)
table=myexcel.sheets()[0]

rows=table.nrows
columns=table.ncols
table_content=[]
for i in range(columns):
    table_content.append([])
print(rows,columns)
for j in range(columns):
    for i in range(rows):
        table_content[j].append(table.cell_value(i, j) ) # 返回单元格中的数据

#ID 列表
id_set=table_content[0]


#修改excel文件
data = openpyxl.load_workbook(savePath) # 可读可写
ws=data['Sheet1']


#图片识别
print(imageObject.size)
# cropped=imageObject.crop((1990,300,2350,950))
cropped=imageObject.crop((1493,220,1750,700))
# cropped=imageObject.crop((1435,220,1700,700))
cropped.show()
text=pytesseract.image_to_string(cropped, lang='chi_sim+eng+chi_tra',config='--psm 1')
# print(text)
text=text.replace(' ','')
text=text.replace('\n','')
print(text)

if(1):
    while(text and len(text)>5):
        playerid=text[0:text.find('对')]
        damage=text[text.find('造成了')+3:text.find('伤害')]
        bossid=bossID(text[text.find('对')+1:text.find('造成了')])
        print(playerid,bossid,damage)
        if playerid in id_set:
            id_index=id_set.index(playerid)
            date_index=date+(date-2)*3
            if(table_content[date_index][id_index] and [ord(c) for c in str(table_content[date_index +1][id_index])]!=[ord(c) for c in damage]):
                id_index=id_index+1
                if (table_content[date_index][id_index] and [ord(c) for c in table_content[date_index +1][id_index]]!=[ord(c) for c in damage]):
                    id_index=id_index+1

            ws.cell(row=id_index+1, column=date_index + 1, value=bossid)
            table_content[date_index ][id_index ]=bossid
            ws.cell(row=id_index+1, column=date_index+2,value=damage)
            ws.cell(row=id_index+1, column=date_index,value=present_round)
            text=text[text.find('伤害')+2:]

        else:
            print(playerid,'not in the ID list')
            text = text[text.find('伤害') + 2:]

    data.save(savePath)


#column0=table.col_slice(0, start_rowx=0, end_rowx=None)
#print(column0)
