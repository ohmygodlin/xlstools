import common
import argparse
import xlrd
import openpyxl
import os
import timeit
import sys

reload(sys)
sys.setdefaultencoding('gbk')

def get_key_indexes(caption_array, keys):
  array = []
  for key in keys:
    for i in range(len(caption_array)):
      if key in caption_array[i]:
        array.append(i)
        break
  if len(array) != len(keys):
    print "NOT enough column for total file, please include at least: "
    for key in keys:
      print key,
    exit(0)
  return array

def add_key(key_sets, data_array, key_indexes):
  for j in range(len(key_indexes)):
    value = data_array[key_indexes[j]]
    if value != '':
      key_sets[j].add(str(value))

def add_total(data_array, total_indexes, total_ws):
  array = []
  for i in total_indexes:
    array.append(data_array[i])
  total_ws.append(array)

def total_to_sets(total, keys, out_total_ws, total_captions):
  print "start loading total: ", total
  in_total_ws = xlrd.open_workbook(total).sheet_by_index(0)
  ncols = in_total_ws.ncols
  caption_array = common.read_row(in_total_ws, 0, ncols)
  for i in caption_array:
    total_captions.append(i)
  out_total_ws.append(caption_array)
  key_indexes = get_key_indexes(caption_array, keys)
  print key_indexes
  key_sets = [set() for j in range(len(key_indexes))]
  for i in range(1, in_total_ws.nrows):
    if i%1000 == 0:
      print i
    data_array = common.read_row(in_total_ws, i, ncols)
    out_total_ws.append(data_array)
    add_key(key_sets, data_array, key_indexes)
  
  print "successful loading total: ", total
  return key_sets

def set_has_key(data_array, key_indexes, key_sets):
  ret = None
  for j in range(len(key_indexes)):
    value = data_array[key_indexes[j]]
    if value != '':
      if str(value) in key_sets[j]:
        return True
      else:
        ret = False
  return ret #all blank would retrun None!
  
def dedup_one(input, output, keys, key_sets, total_ws, total_captions, caption_row, data_row):
  in_ws = xlrd.open_workbook(input).sheet_by_index(0)
  ncols = in_ws.ncols
  
  caption_array = common.read_row(in_ws, caption_row, ncols)
  key_indexes = get_key_indexes(caption_array, keys)
  total_indexes = get_key_indexes(caption_array, total_captions)
  print key_indexes, total_indexes
  
  out_wb = openpyxl.Workbook()
  out_ws = out_wb.active
  out_ws.append(caption_array)

  for i in range(data_row, in_ws.nrows):
    if i%1000 == 0:
      print i
    data_array = common.read_row(in_ws, i, ncols)
    if set_has_key(data_array, key_indexes, key_sets) is False:
      out_ws.append(data_array)
      add_key(key_sets, data_array, key_indexes)
      add_total(data_array, total_indexes, total_ws)
  
  out_wb.save(output)

def dedup(in_dir, total, key_file, out_dir, caption_row, data_row, out_total):
  keys = common.load_array(key_file)
  
  total_wb = openpyxl.Workbook()
  total_ws = total_wb.active
  total_captions = []
  key_sets = total_to_sets(total, keys, total_ws, total_captions)
  print total_captions
  
  common.make_dir(out_dir)
  for i in os.listdir(in_dir):
    if (not i.endswith(common.SUFFIX_XLS)) and (not i.endswith(common.SUFFIX_XLSX)):
      continue
    input = os.path.join(in_dir, i)
    output = os.path.join(out_dir, i)
    if output.endswith(common.SUFFIX_XLS):
      output += 'x'
    print input, output
    dedup_one(input, output, keys, key_sets, total_ws, total_captions, caption_row, data_row)
  
  total_wb.save(out_total)

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='check duplicate from total in target file')
  parser.add_argument("in_dir", help="input target directory", type=str)
  parser.add_argument("total", help="total input file", type=str)
  parser.add_argument("key_file", help="key column file, deduplicate column captions", type=str)
  parser.add_argument("-o", "--out_dir", help="output directory", type=str, default='out')
  parser.add_argument("-c", "--caption_row", help="caption row index in .xlsx", type=int, default=0)
  parser.add_argument("-d", "--data_row", help="data row index in .xlsx", type=int, default=1)
  parser.add_argument("-t", "--out_total", help="output total file", type=str)
  args = parser.parse_args()
  print args
  out_total = args.out_total
  if not out_total:
    out_total = 'total_' + common.TODAY_STR + common.SUFFIX_XLSX
  start = timeit.default_timer()
  dedup(args.in_dir, args.total, args.key_file, args.out_dir, args.caption_row, args.data_row, out_total)
  end = timeit.default_timer()
  print str(end-start)