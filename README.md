# Easy Table to xls sheet

# Creating data
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
df = pd.DataFrame(data=data, index=index)
```
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
*creating a file and adding a sheet*
```
from core.manager import VirtualSheet, VirtualTable, WorkBookManager
excel_file = WorkBookManager("test.xlsx")
v_sheet = excel_file.add_worksheet("Campany sheet", worksheet_class=VirtualSheet)
```
```
table_index_style = excel_file.add_format({"bold": True, "border": 1})
title_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16, "reading_order": 2})
shape_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16})
```

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

`
v_sheet.add_virtual_table(easy_table)
`

`
excel_file.build_all()
`
