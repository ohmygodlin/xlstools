#version: 1.0
#constants and util functions shared by all
import datetime
import re
#import xlrd
import os
import shutil
from time import sleep

FORMAT_DATE = '%Y%m%d'
TODAY_DATETIME = datetime.date.today() # - datetime.timedelta(days=1)
TODAY_STR = format(TODAY_DATETIME, FORMAT_DATE)
SUFFIX_XLSX = '.xlsx'
SUFFIX_XLS = '.xls'
PREFIX_BEFORE = 'before_'
PREFIX_AT = 'at_'
PREFIX_AFTER = 'after_'
PREFIX_DIVIDES = [PREFIX_BEFORE, PREFIX_AT, PREFIX_AFTER]

#https://www.bbsmax.com/A/1O5EP0qbJ7/
def read_cell(ws, row, column):
  try:
    cell = ws.cell(row, column)
    value = cell.value
    #ctype:0 empty, 1 string, 2 number, 3 date, 4 boolean, 5 error
    if cell.ctype == 2 and value % 1 == 0.0:
      value = int(value)
    if not value:
      value = ''
    return value
  except:
    return ''

#NOTICE, must pass the max_column, if use ws.max_column directly, it would be too slow to cal that value again and again!!!
def read_row(ws, row, ncols):
  array = []
  for j in range(ncols):
    array.append(read_cell(ws, row, j))
  return array

def load_array(file):
  f = open(file, 'r')
  array = f.readlines()
  return [i.strip() for i in array]

def make_dir(out_dir):
  if os.path.exists(out_dir):
    c = raw_input(out_dir + " directory would be clean up, press N to stop, or any key to continue.")
    if c.lower() == 'n':
      exit(0)
    shutil.rmtree(out_dir)
    sleep(1)
  
  os.makedirs(out_dir)

def parse_date(s):
  if not s:
    return None
  m = re.search('(\d+)[/\-](\d+)[/\-](\d+)', s)
  if not m:
    return None
  return datetime.date(int(m.group(1)),int(m.group(2)),int(m.group(3)))