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
        if namelast != u'吧':
            a = 1
            # 吧名不规范，请检查
        else:
            if urlfmt:
                a = 0
                # 吧名和URL正确
            else:
                a = 2
                # URL不规范，请检查
    except Exception as e:
        a = 9
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
isheet1.write(0, 7, 'SourceName')
isheet1.col(1).width = 256 * 15
isheet1.col(2).width = 256 * 50
isheet1.col(6).width = 256 * 20
isheet1.col(7).width = 256 * 15

# 开始检查
rows = 1
for eachsrc in srclist:
    # eachsrc[0]:board name
    # eachsrc[1]:board url
    istieba = re.match('http://tieba\.baidu\.com.*', eachsrc[1])
    nameblank = re.match('\s+.*|.*\s+', eachsrc[0])
    urlblank = re.match('\s+.*|.*\s+', eachsrc[1])
    if nameblank or urlblank:
        isheet1.write(rows, 1, eachsrc[0])
        isheet1.write(rows, 2, eachsrc[1])
        isheet1.write(rows, 6, u'Name或URL含有空格，请修改')
    elif istieba:
        cktieba = checktieba(eachsrc[0], eachsrc[1])
        if cktieba == 9:
            isheet1.write(rows, 1, eachsrc[0])
            isheet1.write(rows, 2, eachsrc[1])
            isheet1.write(rows, 6, u'判断过程中出错')
        elif cktieba == 1:
            isheet1.write(rows, 1, eachsrc[0])
            isheet1.write(rows, 2, eachsrc[1])
            isheet1.write(rows, 6, u'吧名不规范，请修改')
        elif cktieba == 2:
            isheet1.write(rows, 1, eachsrc[0])
            isheet1.write(rows, 2, eachsrc[1])
            isheet1.write(rows, 6, u'URL不规范，请修改')
        else:
            sqltieba = "select fid,name,url,bid from board where is_active=1 and fid=101 and name='" + eachsrc[
                0] + "' order by bid"
            cur.execute(sqltieba)
            anslist = cur.fetchall()
            if len(anslist) == 0:
                isheet1.write(rows, 1, eachsrc[0])
                isheet1.write(rows, 2, eachsrc[1])
                isheet1.write(rows, 5, len(anslist))
            else:
                ans = anslist[0]
                isheet1.write(rows, 3, ans[0])
                isheet1.write(rows, 1, ans[1])
                isheet1.write(rows, 2, ans[2])
                isheet1.write(rows, 4, ans[3])
                # isheet1.write(rows,6,'OK')
                isheet1.write(rows, 5, len(anslist))
    else:
        sql0 = "select fid,name,url,bid from board where is_active=1 and url='" + eachsrc[1] + "' order by bid"
        cur.execute(sql0)
        anslist = cur.fetchall()
        # anslist:查询结果列表
        if len(anslist) == 0:
            # isheet1.write(rows,3,'NONE')
            isheet1.write(rows, 1, eachsrc[0])
            isheet1.write(rows, 2, eachsrc[1])
            # isheet1.write(rows,4,'NONE')            
            isheet1.write(rows, 5, len(anslist))
        else:
            ans = anslist[0]
            isheet1.write(rows, 3, ans[0])
            isheet1.write(rows, 1, ans[1])
            isheet1.write(rows, 2, ans[2])
            isheet1.write(rows, 4, ans[3])
            # isheet1.write(rows,6,'OK')
            isheet1.write(rows, 5, len(anslist))
            if ans[1] != eachsrc[0]:
                isheet1.write(rows, 6, u'版面已配置但名称不同')
                isheet1.write(rows, 7, eachsrc[0])
            else:
                pass
    print('complete {}/{}.'.format(rows, len(srclist)))
    rows = rows + 1

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
