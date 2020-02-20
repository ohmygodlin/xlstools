# -*- coding: gbk -*-
import sys
import common
import argparse
import xlrd
import openpyxl
import os

reload(sys)
sys.setdefaultencoding('gbk')

def get_index(array, target):
  for j in range(len(array)):
    if target == array[j]:
      return j
  return None

def get_group_index(groups, s):
  ret = len(groups)
  if s is not None:
    for i in range(len(groups)):
      if groups[i] in s:
        ret = i
        break
  return ret

def write_file(out_dir, groups, caption_array, out_array):
  common.make_dir(out_dir)
    
  for i in range(len(out_array)):
    if len(out_array[i]) <= 0:
      continue
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(caption_array)
    for data in out_array[i]:
      ws.append(data)
    if i < len(groups):
      file = os.path.join(out_dir, groups[i])
    else:
      if ws.max_row <= 1:
        break
      file = os.path.join(out_dir, 'unknown')
    file += common.SUFFIX_XLSX
    wb.save(file)

def group(input, group_file, caption, out_dir, caption_row, data_row):
  ws = xlrd.open_workbook(input).sheet_by_index(0)
  ncols = ws.ncols
  caption_array = common.read_row(ws, caption_row, ncols) 
  caption_column = get_index(caption_array, caption)
  if not caption_column:
    print "Could not find a column with caption: ", caption
    exit(0)
  
  print 'caption_column: ', caption_column
  groups = common.load_array(group_file)
  out_array = [[] for i in range(len(groups)+1)]
  
  for i in range(data_row, ws.nrows):
    if i % 1000 == 0:
      print i
    data_array = common.read_row(ws, i, ncols)
    index = get_group_index(groups, data_array[caption_column])
    out_array[index].append(data_array)
  
  write_file(out_dir, groups, caption_array, out_array)

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='group one .xlsx according to group file')
  parser.add_argument("input", help="input .xlsx file", type=str)
  parser.add_argument("group", help="group file, one group in one line", type=str)
  parser.add_argument("caption", help="caption for that column needed to group by", type=str)
  parser.add_argument("-o", "--out_dir", help="output directory", type=str, default='out')
  parser.add_argument("-c", "--caption_row", help="caption row index", type=int, default=0)
  parser.add_argument("-d", "--data_row", help="data start row index", type=int, default=1)
  args = parser.parse_args()
  print "caption: ", args.caption
  
  group(args.input, args.group, args.caption, args.out_dir, args.caption_row, args.data_row)