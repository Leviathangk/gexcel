# gexcel

一个 excel 便捷读取插入工具，有如下三个模块

- ExcelReader：读取数据（自适应 xlsx、xls）
- ExcelWriter：保存数据（可能会存在丢失长度问题）
- ExcelWriterPerfect：保存数据（不会丢失）

# 安装

```
pip install gexcel
```

# 示例

## ExcelReader

自适应 xlsx、xls  
打开失败很大概率是文件自身的类型和后缀不符

```
from gexcel import ExcelReader

reader = ExcelReader('Result.xlsx')
for i in reader.read_lines():   # 可以输入 sheet 获取 sheet索引 来输出
    print(i)
```

## ExcelWriter

```
from gexcel import ExcelWriter

excel = ExcelWriter()
writer = excel.writer(sheetname='Sheet')    # sheetname 可以不给

for i in range(10):
    writer.write_line([i])

excel.save('Result.xlsx')
```

# ExcelWriterPerfect

```
from gexcel import ExcelWriterPerfect

excel = ExcelWriterPerfect('Result.xlsx')
writer = excel.writer(sheetname='Sheet')    # sheetname 可以不给

for i in range(10):
    writer.write_line([i])

excel.save()
```

# 注意：

这里的 writer 是线程安全的，一个 writer 代表一个 sheet
