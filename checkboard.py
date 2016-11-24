#!/usr/bin/python3
# -*- coding: utf-8 -*-
# filename:checkboard.py


import sys
import xlrd
import xlwt
import psycopg2 as pg2
import re
import time
from mod.getcfg import getcfg


# 定义贴吧检查函数
def checktieba(tbname, tburl):
    namelen = len(tbname)
    namelast = tbname[namelen - 1]
    urlfmt = re.match('^http://tieba\.baidu\.com/f\?kw=[A-za-z0-9%]+$', tburl)
    try:
        if namelast != '吧':
            a = '吧名不规范，请修改后再试'
        else:
            if urlfmt:
                a = 1
            else:
                a = 'URL不规范，请修改后再试'
    except Exception as e:
        a = '判断过程中出错'
        return a
    return a


cfg = getcfg('checkboard.cfg')

print('Connect to datebase ...')

# 连接数据库
try:
    conn = pg2.connect(
        database=cfg['database'],
        user=cfg['user'],
        password=cfg['password'],
        host=cfg['host'],
        port=cfg['port']
    )
    cur = conn.cursor()
except Exception as e:
    print('Connect server failed.')
    anyenter = input('Press ENTER to quit.')
else:
    print('Successfully connect to datebase.')

# 读取要检查的 board
openboard = xlrd.open_workbook('board.xls')
boardsheet1 = openboard.sheets()[0]
srclist = []

for row in range(boardsheet1.nrows)[1:]:
    values = []
    for col in range(boardsheet1.ncols):
        values.append(boardsheet1.cell(row, col).value)
    srclist.append(values)

print('There are {} boards need to be checked ..'.format(len(srclist)))
print('*' * 78)

# 初始化输出文件，设置标题，列宽
stylecol = 1
namecol = 2
urlcol = 3
fidcol = 4
idcol = 5
countcol = 6
statuscol = 7
srcnamecol = 8
init = xlwt.Workbook(encoding='utf-8')
sheet1 = init.add_sheet('sheet1')
sheet1.write(0, stylecol, 'Style')
sheet1.write(0, namecol, 'Name')
sheet1.write(0, urlcol, 'URL')
sheet1.write(0, fidcol, 'Fid')
sheet1.write(0, idcol, 'Bid/Nsid')
sheet1.write(0, countcol, 'Count')
sheet1.write(0, statuscol, 'Status')
sheet1.write(0, srcnamecol, '信息源Name')
sheet1.col(stylecol).width = 256 * 5
sheet1.col(namecol).width = 256 * 15
sheet1.col(urlcol).width = 256 * 50
sheet1.col(statuscol).width = 256 * 20
sheet1.col(srcnamecol).width = 256 * 15

# 开始检查
rows = 1
for boardname, boardurl in srclist:
    urlblank = re.match('\s+.*|.*\s+', boardurl)
    nameblank = re.match('\s+.*|.*\s+', boardname)
    istieba = re.match('http://tieba\.baidu\.com.*', boardurl)
    isweibo = re.match('http://(e\.)?weibo\.com/.*', boardurl)
    if nameblank or urlblank:
        sheet1.write(rows, namecol, boardname)
        sheet1.write(rows, urlcol, boardurl)
        sheet1.write(rows, statuscol, 'Name或URL含有空格，请修改后再试')
    elif istieba:
        cktieba = checktieba(boardname, boardurl)
        if cktieba != 1:
            sheet1.write(rows, namecol, boardname)
            sheet1.write(rows, urlcol, boardurl)
            sheet1.write(rows, statuscol, cktieba)
        else:
            findtieba = "select fid,name,url,bid from board where is_active=1 \
            and fid=101 and name='%s' order by bid" % (boardname)
            cur.execute(findtieba)
            # returnlist:查询结果列表
            # data[0]:fid
            # data[1]:name
            # data[2]:url
            # data[3]:bid
            returnlist = cur.fetchall()
            if len(returnlist) == 0:
                # 若结果为0，则保存count为0
                sheet1.write(rows, namecol, boardname)
                sheet1.write(rows, urlcol, boardurl)
                sheet1.write(rows, countcol, 0)
            else:
                # 反之，则保存结果中第一条记录的各字段和结果总数
                data = returnlist[0]
                sheet1.write(rows, stylecol, 1)
                sheet1.write(rows, fidcol, data[0])
                sheet1.write(rows, namecol, data[1])
                sheet1.write(rows, urlcol, data[2])
                sheet1.write(rows, idcol, data[3])
                sheet1.write(rows, countcol, len(returnlist))
    elif isweibo:
        findweibo = "select 4 as style,b.fid,b.name,b.url,b.bid,b.is_active,\
        wu.uid from weibo_user wu full join board b on wu.bid=b.bid where \
        wu.name='%s' order by b.bid" % (boardname)
        cur.execute(findweibo)
        returnlist = cur.fetchall()
        if len(returnlist) == 0:
            sheet1.write(rows, stylecol, 4)
            sheet1.write(rows, namecol, boardname)
            sheet1.write(rows, urlcol, boardurl)
            sheet1.write(rows, countcol, 0)
        else:
            data = returnlist[0]
            if data[5] == 1:
                sheet1.write(rows, stylecol, 4)
                sheet1.write(rows, fidcol, data[1])
                sheet1.write(rows, namecol, data[2])
                sheet1.write(rows, urlcol, data[3])
                sheet1.write(rows, idcol, data[4])
                weibouid = 'uid:'+str(data[6])
                sheet1.write(rows, statuscol, weibouid)
                sheet1.write(rows, countcol, len(returnlist))
            else:
                sheet1.write(rows, stylecol, 4)
                sheet1.write(rows, fidcol, data[1])
                sheet1.write(rows, namecol, data[2])
                sheet1.write(rows, urlcol, data[3])
                sheet1.write(rows, idcol, data[4])
                weibouid = 'uid:'+str(data[6])+' uid已存在，对应版面已被停止，请检\
                查'
                sheet1.write(rows, statuscol, weibouid)
                sheet1.write(rows, countcol, len(returnlist))
    else:
        findboard = "select w.style,b.fid,b.name,b.url,b.bid from board b left \
        join website w on b.fid=w.fid where b.is_active=1 and b.url='%s' order \
        by b.bid" % (boardurl)
        cur.execute(findboard)
        returnlist = cur.fetchall()
        if len(returnlist) == 0:
            findnews = "select style,fid,name,url,nsid from news_site where \
            is_active=1 and url='%s'" % (boardurl)
            cur.execute(findnews)
            returnlist = cur.fetchall()
            if len(returnlist) == 0:
                sheet1.write(rows, namecol, boardname)
                sheet1.write(rows, urlcol, boardurl)
                sheet1.write(rows, countcol, 0)
            else:
                data = returnlist[0]
                data = returnlist[0]
                sheet1.write(rows, stylecol, data[0])
                sheet1.write(rows, fidcol, data[1])
                sheet1.write(rows, namecol, data[2])
                sheet1.write(rows, urlcol, data[3])
                sheet1.write(rows, idcol, data[4])
                sheet1.write(rows, countcol, len(returnlist))
                if data[2] != boardname:
                    sheet1.write(rows, statuscol, '版面已配置但名称不同')
                    sheet1.write(rows, srcnamecol, boardname)
                else:
                    pass
        else:
            data = returnlist[0]
            sheet1.write(rows, stylecol, data[0])
            sheet1.write(rows, fidcol, data[1])
            sheet1.write(rows, namecol, data[2])
            sheet1.write(rows, urlcol, data[3])
            sheet1.write(rows, idcol, data[4])
            sheet1.write(rows, countcol, len(returnlist))
            if data[2] != boardname:
                sheet1.write(rows, statuscol, '版面已配置但名称不同')
                sheet1.write(rows, srcnamecol, boardname)
            else:
                pass
    print('complete {}/{}.'.format(rows, len(srclist)))
    rows += 1

# 关闭连接
# conn.commit()
conn.close()

print('*' * 78)

# 保存结果
filename = time.strftime('checkboard_result_%Y%m%d_%H%M%S.xls')
init.save(filename)

time.sleep(1)

anyenter = input('Check result has been saved.Press Enter to quit.')
if anyenter:
    sys.exit()
else:
    sys.exit()
