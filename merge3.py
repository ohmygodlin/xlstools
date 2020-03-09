#python3 merge3.py input -c 1 -d 2
import common
import argparse
import pandas
import os
import timeit
import sys

def merge(in_dir, output, caption_row, data_row):
  is_first = True
  startrow = 0 #help to implement append excel with startrow
  with pandas.ExcelWriter(output) as writer:
    for i in os.listdir(in_dir):
      if (not i.endswith(common.SUFFIX_XLS)) and (not i.endswith(common.SUFFIX_XLSX)):
        continue
      file = os.path.join(in_dir, i)
      print(file)

      if is_first:
        df = pandas.read_excel(file, header=caption_row)
        df.to_excel(writer, index=False, startrow=startrow)
        is_first = False
        startrow += df.shape[0] + 1
        print(startrow)
        continue

      df = pandas.read_excel(file, header=None, skiprows=data_row)
      df.to_excel(writer, header=False, index=False, startrow=startrow)
      startrow += df.shape[0]
      print(startrow)

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='merge all excel file into one')
  parser.add_argument("in_dir", help="input directory", type=str)
  parser.add_argument("-o", "--output", help="output file", type=str, default='out.xlsx')
  parser.add_argument("-c", "--caption_row", help="caption row index", type=int, default=0)
  parser.add_argument("-d", "--data_row", help="data start row index", type=int)
  args = parser.parse_args()
  data_row = args.data_row
  if data_row is None:
    data_row = args.caption_row + 1
  print(args)
  
  start = timeit.default_timer()
  merge(args.in_dir, args.output, args.caption_row, data_row)
  end = timeit.default_timer()
  print(str(end-start))