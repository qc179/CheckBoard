#!/usr/bin/env python
#-*- coding: utf-8 -*-
#filename:checkmod.py
#You need root power when install on linux

import os,time

cwd = os.getcwd()

print '\nChecking module xlrd ..'

try:
    import xlrd
except Exception, e:
    print '\nStart to setup module xlrd ...'
    os.chdir(cwd+'\\xlrd-0.9.4')
    os.system('python setup.py install')
    print '\nFinished.'
else:
    print '\nThe module xlrd was already exists.'
time.sleep(1)

print '\nChecking module xlwt ..'

try:
    import xlwt
except Exception, e:
    print '\nStart to setup module xlwt ...'
    os.chdir(cwd+'\\xlwt-1.0.0')
    os.system('python setup.py install')
    print '\nFinished.'
else:
    print '\nThe module xlwt was already exists.'
time.sleep(1)

print '\nChecking module psycopg2 ..'

try:
    import psycopg2
except Exception, e:
    print '\nStart to setup module psycopg2 ...'
    os.chdir(cwd)
    os.system('pip install psycopg2-2.6.1-cp27-none-win32.whl')
    print '\nFinished.'
else:
    print '\nThe module psycopg2 was already exists.'
time.sleep(1)

getany = raw_input('\npress Enter to quit. ')