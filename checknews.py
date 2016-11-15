#!/usr/bin/python3
# -*- coding: utf-8 -*-
# filename:checknews.py


import sys
import xlrd
import xlwt
import psycopg2 as pg2
import re
import time
from mod.getcfg import getcfg

cfg = getcfg('checkboard.cfg')

print('Connect to datebase ...')

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

# 读取将要检查的 news
opennews = xlrd.open_workbook('news.xls')
newssheet1 = opennews.sheets()[0]
srclist = []

for row in range(newssheet1.nrows)[1:]:
    values = []
    for col in range(newssheet1.ncols):
        values.append(newssheet1.cell(row, col).value)
    srclist.append(values)

print('There are {} news need to be checked ..'.format(len(srclist)))
print('*' * 78)

# 初始化输出文件，设置标题，列宽
init = xlwt.Workbook(encoding='utf-8')
isheet1 = init.add_sheet('sheet1')
isheet1.write(0, 0, ' ')
isheet1.write(0, 1, 'Name')
isheet1.write(0, 2, 'URL')
isheet1.write(0, 3, 'Fid')
isheet1.write(0, 4, 'Nsid')
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
    # eachsrc[0]:news name
    # eachsrc[1]:news url
    nameblank = re.match('\s+.*|.*\s+', eachsrc[0])
    urlblank = re.match('\s+.*|.*\s+', eachsrc[1])
    if nameblank or urlblank:
        isheet1.write(rows, 1, eachsrc[0])
        isheet1.write(rows, 2, eachsrc[1])
        isheet1.write(rows, 6, 'Name或URL含有空格，请修改')
    else:
        sql0 = "select fid,name,url,nsid from news_site where is_active=1 and \
        url='" + eachsrc[1] + "' order by nsid"
        cur.execute(sql0)
        anslist = cur.fetchall()
        # anslist:a list of select results
        if len(anslist) == 0:
            # isheet1.write(rows,3,'NONE')
            isheet1.write(rows, 1, eachsrc[0])
            isheet1.write(rows, 2, eachsrc[1])
            # isheet1.write(rows,4,'NONE')
            isheet1.write(rows, 6, '未查到这个URL')
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
                isheet1.write(rows, 6, '版面已配置但名称不同')
                isheet1.write(rows, 7, eachsrc[0])
            else:
                pass
    print('complete {}/{}.'.format(rows, len(srclist)))
    rows = rows + 1

# 保存结果
filename = time.strftime('checknews_result_%Y%m%d_%H%M%S.xls')
init.save(filename)

# conn.commit()
# 关闭连接
conn.close()

print('*' * 78)

anyenter = input('Check result has been saved.Press Enter to quit.')
if anyenter:
    sys.exit()
else:
    sys.exit()
