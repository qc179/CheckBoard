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
init = xlwt.Workbook(encoding='utf-8')
isheet1 = init.add_sheet('sheet1')
isheet1.write(0, 0, ' ')
isheet1.write(0, 1, 'Name')
isheet1.write(0, 2, 'URL')
isheet1.write(0, 3, 'Fid')
isheet1.write(0, 4, 'Bid')
isheet1.write(0, 5, 'Count')
isheet1.write(0, 6, 'Status')
isheet1.write(0, 7, '信息源Name')
isheet1.col(1).width = 256 * 15
isheet1.col(2).width = 256 * 50
isheet1.col(6).width = 256 * 20
isheet1.col(7).width = 256 * 15

# 开始检查
rows = 1
for boardname, boardurl in srclist:
    urlblank = re.match('\s+.*|.*\s+', boardurl)
    istieba = re.match('http://tieba\.baidu\.com.*', boardurl)
    nameblank = re.match('\s+.*|.*\s+', boardname)
    if nameblank or urlblank:
        isheet1.write(rows, 1, boardname)
        isheet1.write(rows, 2, boardurl)
        isheet1.write(rows, 6, 'Name或URL含有空格，请修改后再试')
    elif istieba:
        cktieba = checktieba(boardname, boardurl)
        if cktieba != 1:
            isheet1.write(rows, 1, boardname)
            isheet1.write(rows, 2, boardurl)
            isheet1.write(rows, 6, cktieba)
        else:
            findtieba = "select fid,name,url,bid from board where is_active=1 \
            and fid=101 and name='" + boardname + "' order by bid"
            cur.execute(findtieba)
            # returnlist:查询结果列表
            # data[0]:fid
            # data[1]:name
            # data[2]:url
            # data[3]:bid
            returnlist = cur.fetchall()
            if len(returnlist) == 0:
                isheet1.write(rows, 1, boardname)
                isheet1.write(rows, 2, boardurl)
                isheet1.write(rows, 5, 0)
            else:
                # 若查询结果不为0，则保存结果中第一条记录的各字段
                data = returnlist[0]
                isheet1.write(rows, 3, data[0])
                isheet1.write(rows, 1, data[1])
                isheet1.write(rows, 2, data[2])
                isheet1.write(rows, 4, data[3])
                # 保存查询结果的数量
                isheet1.write(rows, 5, len(returnlist))
    else:
        findboard = "select fid,name,url,bid from board where is_active=1 and \
        url='" + boardurl + "' order by bid"
        cur.execute(findboard)
        returnlist = cur.fetchall()
        if len(returnlist) == 0:
            isheet1.write(rows, 1, boardname)
            isheet1.write(rows, 2, boardurl)
            isheet1.write(rows, 5, 0)
        else:
            data = returnlist[0]
            isheet1.write(rows, 3, data[0])
            isheet1.write(rows, 1, data[1])
            isheet1.write(rows, 2, data[2])
            isheet1.write(rows, 4, data[3])
            isheet1.write(rows, 5, len(returnlist))
            if data[1] != boardname:
                isheet1.write(rows, 6, '版面已配置但名称不同')
                isheet1.write(rows, 7, boardname)
            else:
                pass
    print('complete {}/{}.'.format(rows, len(srclist)))
    rows += 1

# 保存结果
filename = time.strftime('checkboard_result_%Y%m%d_%H%M%S.xls')
init.save(filename)

# 关闭连接
# conn.commit()
conn.close()
print('*' * 78)

anyenter = input('Check result has been saved.Press Enter to quit.')
if anyenter:
    sys.exit()
else:
    sys.exit()
