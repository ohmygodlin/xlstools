import common
import argparse
import xlrd
import openpyxl
import os
import timeit
import sys

reload(sys)
sys.setdefaultencoding('gbk')

def merge(in_dir, output, caption_row, data_row):
  out_wb = openpyxl.Workbook()
  out_ws = out_wb.active
  caption_array = []
  for i in os.listdir(in_dir):
    if (not i.endswith(common.SUFFIX_XLS)) and (not i.endswith(common.SUFFIX_XLSX)):
      continue
    file = os.path.join(in_dir, i)
    print file
    in_ws = xlrd.open_workbook(file).sheet_by_index(0)
    ncols = in_ws.ncols
    if len(caption_array) <= 0:
      caption_array = common.read_row(in_ws, caption_row, ncols)
      out_ws.append(caption_array)
    
    for i in range(data_row, in_ws.nrows):
      data_array = common.read_row(in_ws, i, ncols)
      if all([not i or i == '' for i in data_array]):
        continue
      out_ws.append(data_array)
      
  out_wb.save(output)

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='merge all input .xlsx into one')
  parser.add_argument("in_dir", help="input directory", type=str)
  parser.add_argument("-o", "--output", help="output file", type=str, default='out.xlsx')
  parser.add_argument("-c", "--caption_row", help="caption row index", type=int, default=0)
  parser.add_argument("-d", "--data_row", help="data start row index", type=int, default=1)
  args = parser.parse_args()
  print args
  
  start = timeit.default_timer()
  merge(args.in_dir, args.output, args.caption_row, args.data_row)
  end = timeit.default_timer()
  print str(end-start)