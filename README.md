# Introduction
This lib came in between **pandas** and **xlsxwriter**, it helps to create complex sheets (separated tables, tables with no digits). 
In fast summary, the idea is to think that an excel sheet in one way or other will be built with multiple tables that each one can contain:
- Title for the whole table.
- Columns (head) for the data.
- Index.
- Data.

and each of those elements has its own style ("text-align", "font_size", "reading_direction", etc...).

let take a demonstration of how it works:

# Data creation
```
import random 
import pandas as pd
index = ["Net income", "Minority rights","Net revenue", "Net loans", "Total debt", "Earnings", "consumption", "Treasury stocks"]
random.shuffle(index)
data = {
    "2013": [random.randint(1, 100) for _ in range(0, 8)],
    "2014": [random.randint(1, 100) for _ in range(0, 8)],
    "2015": [random.randint(1, 100) for _ in range(0, 8)],
    "2016": [random.randint(1, 100) for _ in range(0, 8)],
}
```

** Pandas is used for validating the data**<br>
`df = pd.DataFrame(data=data, index=index)`

# Preview
`print(df)`
```
2013  2014  2015  2016
Net income        100    91    87    86
Net revenue        99    87    22    54
Net loans          10    51    35    93
consumption        41    46    12    70
Treasury stocks    52     8    11    48
Earnings           61    64    98     8
Total debt         31    74    37   100
Minority rights    36    77    79    98
```
# Use of the library

```
from core.manager import VirtualSheet, VirtualTable, WorkBookManager

excel_file = WorkBookManager("test.xlsx") # The initialisation of the Workbook 
v_sheet = excel_file.add_worksheet("Company sheet", worksheet_class=VirtualSheet) # Add a Worksheet to the book
```
# Styles preparation
**Style is a <a href="https://xlsxwriter.readthedocs.io/format.html"> xlsxwriter Format object </a>**
```
table_index_style = excel_file.add_format({"bold": True, "border": 1})
title_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16, "reading_order": 2})
shape_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16})
```
# Table creation
```
easy_table = VirtualTable(df, 0, 0,
                          display_index=True,
                          display_head=True,
                          title="Financial analysis",
                          title_style=title_style,
                          index_style=table_index_style,
                          shape_style=shape_style,
                          head_style=shape_style,
                          to_xls_table=True
                          )
```
# Add it to the worksheet
`v_sheet.add_virtual_table(easy_table)`

# Previewing table object
**Coordinates**<br>
```
for k,v in easy_table.coordinates.items():
    print(k, "=>", v["start"], v["end"])
``` 
```
title => (0, 1) (0, 4)
head => (1, 1) (1, 4)
index => (2, 0) (9, 0)
shape => (2, 1) (9, 4)
```
**Cursor**

`print(easy_table.cr)`

`(9, 4)`

# Build the file
`
excel_file.build_all()
`
# Result
**Result with `to_xls_table=True`** <br>
<img src="https://github.com/bguernouti/easy_table_to_excel/blob/master/to_xls_table.png" width="350" alt="to_xls_table enabled" />

**Result with out `to_xls_table`**<br>
<img src="https://github.com/bguernouti/easy_table_to_excel/blob/master/simple.png" width="350" alt="to_xls_table disabled" />

**Idea of complex sheet**<br>
<img src="https://github.com/bguernouti/easy_table_to_excel/blob/master/complex.png" alt="to_xls_table disabled" />
> check **complex.py**

Unfortunately, this kind of complexity can not be done using **pandas** and **xlsxwriter** only.

# Upcomming features
1- For sure, charts generation.
