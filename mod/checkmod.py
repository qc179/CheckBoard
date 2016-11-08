#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys

try:
    print('Checking module psycopg2 ...')
    import psycopg2
except Exception as e:
    print('Start to install module psycopg2 ...')
    os.system('pip install psycopg2==2.6.2 -q -i https://pypi.doubanio.com/simple/')
    print('Successfully installed psycopg2-2.6.2\n')
else:
    print('The module psycopg2 already exists.\n')

try:
    print('Checking module xlrd ...')
    import xlrd
except Exception as e:
    print('Start to install module xlrd ...')
    os.system('pip install xlrd==1.0.0 -q -i https://pypi.doubanio.com/simple/')
    print('Successfully installed xlrd-1.0.0\n')
else:
    print('The module xlrd already exists.\n')

try:
    print('Checking module xlwt ...')
    import xlwt
except Exception as e:
    print('Start to install module xlwt ...')
    os.system('pip install xlwt==1.0.0 -q -i https://pypi.doubanio.com/simple/')
    print('Successfully installed xlwt-1.0.0\n')
else:
    print('The module xlwt already exists.\n')

getany = input('Check complete.Press Enter to quit. ')
if getany:
    sys.exit()
else:
    sys.exit()
