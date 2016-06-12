#!/usr/bin/env python
#-*- coding: utf-8 -*-
#filename:getcfg.py

def getcfg(filename):
    try:
        with open(filename,'rb') as cfg:
            cfglist = cfg.readlines()
            a = []
            for i in range(len(cfglist)):
                if cfglist[i] == '\r\n':
                    pass
                elif cfglist[i] == '\n':
                    pass
                elif cfglist[i] == '\r':
                    pass
                else:
                    cfglist[i] = cfglist[i].replace('\r','')
                    cfglist[i] = cfglist[i].replace('\n','')
                    cfglist[i] = cfglist[i].replace(' ','')
                    cfglist[i] = cfglist[i].split('=')
                    a.append((cfglist[i][0],cfglist[i][1]))
        a = dict(a)
    except Exception,e:
        print e
        print 'Please check checkboard.cfg,looks like something configured \
wrong.'
        anyenter = raw_input('Press ENTER to confirm.')
    return a

if __name__ == '__main__':
    a = getcfg('checkboard.cfg')
    for i in a.keys():
        print i+'='+a[i]
    quit = raw_input('Press ENTER to quit')
