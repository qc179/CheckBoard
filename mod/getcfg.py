#!/usr/bin/env python
# -*- coding: utf-8 -*-
# filename:getcfg.py

def getcfg(filename):
    try:
        with open(filename, 'r') as cfg:
            readlist = cfg.readlines()
            cfglist = []
            for i in range(len(readlist)):
                if readlist[i] == '\r\n':
                    pass
                elif readlist[i] == '\n':
                    pass
                elif readlist[i] == '\r':
                    pass
                else:
                    readlist[i] = readlist[i].replace('\r', '')
                    readlist[i] = readlist[i].replace('\n', '')
                    readlist[i] = readlist[i].replace(' ', '')
                    readlist[i] = readlist[i].split('=')
                    cfglist.append((readlist[i][0], readlist[i][1]))
        cfgdict = dict(cfglist)
    except Exception as e:
        print(e)
        print('Please check checkboard.cfg,looks like something configured wrong.')
        anyenter = input('Press ENTER to confirm.')
    return cfgdict


if __name__ == '__main__':
    a = getcfg('checkboard.cfg')
    for i in a.keys():
        print(i + '=' + a[i])
    quit = input('\nPress ENTER to quit.')
