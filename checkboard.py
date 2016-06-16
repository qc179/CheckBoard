#!/usr/bin/env python
#-*- coding: utf-8 -*-
#filename:checkboard.py


import sys
sys.path.append('mod')
import xlrd
import xlwt
import psycopg2 as pg2
import getcfg
import re

cfg = getcfg.getcfg('checkboard.cfg')

#try connect database
try:
    conn = pg2.connect(
    database=cfg['database'],
    user=cfg['user'],
    password=cfg['password'],
    host=cfg['host'],
    port=cfg['port']
    )
    cur = conn.cursor()
except Exception,e:
    #print e
    print 'Connect server failed.'
    anyenter = raw_input('Press ENTER to quit.')

#make board source into a list from sheet1
board = xlrd.open_workbook('board.xls')
bsheet1 = board.sheets()[0]
srclist = []
for row in range(bsheet1.nrows)[1:]:
    values = []
    for col in range(bsheet1.ncols):
        values.append(bsheet1.cell(row,col).value)
    srclist.append(values)

#init board_result.xls such as title and column width
initsave = xlwt.Workbook(encoding='utf-8')
isheet1 = initsave.add_sheet('sheet1')
isheet1.write(0,0,' ')
isheet1.write(0,1,'Name')
isheet1.write(0,2,'URL')
isheet1.write(0,3,'Fid')
isheet1.write(0,4,'Bid')
isheet1.write(0,5,'Count')
isheet1.write(0,6,'Status')
isheet1.write(0,7,'SourceName')
isheet1.col(1).width = 256*15
isheet1.col(2).width = 256*50
isheet1.col(6).width = 256*20
isheet1.col(7).width = 256*15

def checktieba(tbname,tburl):
    namelen = len(tbname)
    namelast = tbname[namelen-1]
    urlfmt = re.match('^http://tieba\.baidu\.com/f\?kw=[A-za-z0-9%]+$',\
                      tburl)
    try:
        if namelast != u'吧' :
            a = 1
            # tieba.write(rows,2,u'吧名不规范，请检查')
        else :
            if urlfmt:
                a = 0
                # tieba.write(rows,2,u'吧名和URL正确')
            else:
                a = 2
                # tieba.write(rows,2,u'URL不规范，请检查')
    except Exception, e:
        a = 9
        return a
    return a

#judge each source whether has been inserted,then save them
print 'There are {} boards need to be checked ..'.format(len(srclist))
print '*'*78
rows = 1
for eachsrc in srclist:
    #eachsrc[0]:board name
    #eachsrc[1]:board url
    istieba = re.match('http://tieba\.baidu\.com.*',eachsrc[1])
    nameblank = re.match('\s+(?u)\w+|(?u)\w+\s+',eachsrc[0])
    urlblank = re.match('\s+.*|.*\s+',eachsrc[1])
    if nameblank or urlblank:
        isheet1.write(rows,1,eachsrc[0])
        isheet1.write(rows,2,eachsrc[1])
        isheet1.write(rows,6,u'Name或URL含有空格，请修改')
    elif istieba:
        cktieba = checktieba(eachsrc[0],eachsrc[1])
        if cktieba == 9:
            isheet1.write(rows,1,eachsrc[0])
            isheet1.write(rows,2,eachsrc[1])
            isheet1.write(rows,6,u'判断过程中出错')
        elif cktieba == 1:
            isheet1.write(rows,1,eachsrc[0])
            isheet1.write(rows,2,eachsrc[1])
            isheet1.write(rows,6,u'吧名不规范，请修改')
        elif cktieba == 2:
            isheet1.write(rows,1,eachsrc[0])
            isheet1.write(rows,2,eachsrc[1])
            isheet1.write(rows,6,u'URL不规范，请修改')
        else:
            sqltieba = "select fid,name,url,bid from board where \
            is_active=1 and fid=101 and name='"+eachsrc[0]+"' order by bid"
            cur.execute(sqltieba)
            anslist = cur.fetchall()   
            if len(anslist) == 0:
                isheet1.write(rows,1,eachsrc[0])
                isheet1.write(rows,2,eachsrc[1])                
                isheet1.write(rows,5,len(anslist))
            else:
                ans = anslist[0]
                isheet1.write(rows,3,ans[0])
                isheet1.write(rows,1,ans[1])
                isheet1.write(rows,2,ans[2])
                isheet1.write(rows,4,ans[3])
                #isheet1.write(rows,6,'OK')
                isheet1.write(rows,5,len(anslist))            
    else:
        sql0 = "select fid,name,url,bid from board where is_active=1 \
        and url='"+eachsrc[1]+"' order by bid"
        cur.execute(sql0)
        anslist = cur.fetchall()
        #anslist:a list of select results    
        if len(anslist) == 0:
            #isheet1.write(rows,3,'NONE')
            isheet1.write(rows,1,eachsrc[0])
            isheet1.write(rows,2,eachsrc[1])
            # isheet1.write(rows,4,'NONE')            
            isheet1.write(rows,5,len(anslist))
        else:
            ans = anslist[0]
            isheet1.write(rows,3,ans[0])
            isheet1.write(rows,1,ans[1])
            isheet1.write(rows,2,ans[2])
            isheet1.write(rows,4,ans[3])
            #isheet1.write(rows,6,'OK')
            isheet1.write(rows,5,len(anslist))            
            unicodename=ans[1].decode('utf-8')
            if unicodename != eachsrc[0]:
                isheet1.write(rows,6,u'版面已配置但名称不同')
                isheet1.write(rows,7,eachsrc[0])
            else:
                pass
    print 'complete {}/{}.'.format(rows,len(srclist))
    rows = rows+1

initsave.save('board_result.xls')

conn.commit()
conn.close()
print '*'*78
anyenter = raw_input('Check result has been saved.Press Enter to quit.')
