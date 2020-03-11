#python3 merge3.py input -s Sheet2 -k sheet2|test1
import argparse
import pandas
import os
import timeit

def extract_one(in_df, keywords, writer, sheet, startrow):
  out_df = pandas.DataFrame(columns=in_df.columns)
  for index, row in in_df.iterrows():
    if row.astype(str).str.contains(keywords).sum() >= 1:
      out_df = out_df.append(row, ignore_index=True)
  
  if out_df.empty:
    return startrow
  
  is_first = (startrow==0)
  out_df.to_excel(writer, sheet_name=sheet if sheet is not None else 'Sheet1', header=is_first, index=False, startrow=startrow)
  
  startrow += out_df.shape[0]
  if is_first:
    startrow += 1
  return startrow

def extract(in_dir, sheet, keywords, output):
  startrow = 0 #help to implement append excel with startrow
  with pandas.ExcelWriter(output) as writer:
    for i in os.listdir(in_dir):
      if (not i.endswith('.xls')) and (not i.endswith('.xlsx')):
        continue
      
      file = os.path.join(in_dir, i)
      print(file)
      
      try:
        excel = pandas.read_excel(file, sheet_name=sheet)
        if isinstance(excel, dict):
          for k, df in excel.items():
            startrow = extract_one(df, keywords, writer, sheet, startrow)
        else:
          startrow = extract_one(excel, keywords, writer, sheet, startrow)
      
      except Exception as e:
        print(e)

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='extract keyword lines into one excel')
  parser.add_argument("in_dir", help="input directory", type=str)
  parser.add_argument("-s", "--sheet", help="specific sheet", type=str)
  parser.add_argument("-k", "--keywords", help="search keywords, delimitered by |", type=str)
  parser.add_argument("-o", "--output", help="output file", type=str, default='out.xlsx')
  args = parser.parse_args()
  print(args)
  
  start = timeit.default_timer()
  extract(args.in_dir, args.sheet, args.keywords, args.output)
  end = timeit.default_timer()
  print(str(end-start))