from core.manager import VirtualSheet, WorkBookManager, VirtualTable
import pandas as pd

excel_file = WorkBookManager("complex.xlsx")
complex_sheet = excel_file.add_worksheet("Multi format", worksheet_class=VirtualSheet)
complex_sheet.set_column("A:G", 30)

# Let assume both share same title style
title_style = excel_file.add_format({"align": "center", "font_size": 16})

# Left to right styles
head_style_en = excel_file.add_format({"align": "center", "font_size": 12})
data_style_en = excel_file.add_format({"align": "left", "font_size": 11, "border": 1})

# Left to right styles
head_style_ar = excel_file.add_format({"align": "center", "font_size": 12})
data_style_ar = excel_file.add_format({"align": "right", "font_size": 11, "border": 1})

# English data
en_df = pd.read_csv("en.csv")
table_en = VirtualTable(en_df, 0, 0,
                        display_head=True,
                        head_style=head_style_en,
                        title="List of largest companies by revenue",
                        title_style=title_style,
                        shape_style=data_style_en,
                        to_xls_table=True
                        )
complex_sheet.add_virtual_table(table_en)
end_row, end_col = table_en.cr  # Recover table ending coordinates, to use it next

# Arabic data
ar_df = pd.read_csv("ar.csv")
ar_df = ar_df[ar_df.columns[::-1]]
table_ar = VirtualTable(ar_df, end_row+3, 0,
                        display_head=True,
                        head_style=head_style_ar,
                        title="قائمة أكبر الشركات حسب الإيرادات",
                        title_style=title_style,
                        shape_style=data_style_ar,
                        to_xls_table=True,
                        )
complex_sheet.add_virtual_table(table_ar)

# Building
excel_file.build_all()
