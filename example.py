import random
import pandas as pd
from core.manager import VirtualSheet, VirtualTable, WorkBookManager

index = [
    "Net income",
    "Minority rights",
    "Net revenue",
    "Net loans",
    "Total debt",
    "Earnings",
    "consumption",
    "Treasury stocks"
]
random.shuffle(index)

data = {
    "2013": [random.randint(1, 100) for _ in range(0, 8)],
    "2014": [random.randint(1, 100) for _ in range(0, 8)],
    "2015": [random.randint(1, 100) for _ in range(0, 8)],
    "2016": [random.randint(1, 100) for _ in range(0, 8)],
}

df = pd.DataFrame(data=data, index=index)

excel_file = WorkBookManager("test.xlsx")  # The initialisation of the Workbook
v_sheet = excel_file.add_worksheet("Company sheet", worksheet_class=VirtualSheet)  # Added a Worksheet

table_index_style = excel_file.add_format({"bold": True, "border": 1})
title_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16, "reading_order": 2})
shape_style = excel_file.add_format({"bold": True, "align": "center", "font_size": 16})

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

v_sheet.add_virtual_table(easy_table)

excel_file.build_all()
