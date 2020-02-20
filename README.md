# xlstools
Python2 process xls/xlsx, reading with xlrd, writing with openpyxl

### 使用说明

默认标题行是第1行，数据行是第2行，如果标题行是第2行，数据行是第3行，则指定`-c 1 -d 2`（python读行从0开始，所以指定-c 1则表示标题行为第2行）

#### 功能1：去重

```shell
python ./dedup.py input total_20200205.xlsx config/key.txt
```

参数1：要去重表的目录（可放一张或多张表），参数2：total文件，参数3：作为key的字段（当前用证件和手机）。默认输出为当前out目录，如果需要更改目录，指定`-o c:\`，默认新的total文件为`total_[今天日期].xlsx`，如果需要更改，指定`-t total_xxxxx.xlsx`。

#### 功能2：分类

```shell
python ./group.py input.xlsx config/city.txt 市县
```

参数1：要分类的表，参数2：类别文件（每行一个类别），参数3：为分类的字段字段名。

#### 功能3：合并相同格式表

```shell
python ./merge.py input
```

参数1：需要合并表存放的目录。默认输出为当前`out.xlsx`，如果需要更改输出文件，指定`-o out_xxxxx.xlsx`，注意：必须是`xlsx`后缀。
