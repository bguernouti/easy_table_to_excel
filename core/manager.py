from typing import List, Dict

from pandas import DataFrame
from xlsxwriter import Workbook
from xlsxwriter.workbook import Format
from xlsxwriter.worksheet import Worksheet


class VirtualTable:

    """
    Independent table class, it hold all necessary information about the data frame
    (
    title => a text that describe a full table (used a lot),
    index => is like a header of rows (this add an extra column to the beginning),
    header => data header (this add an extra row at the top),
    shape "the data" => the indispensable data, this is validated by 'pandas'
    )
    Parameters
    ----------
    df : DataFrame (pandas DataFrame object)
    start_row : int (at which row index in the excel sheet, the table starts)
    start_col : int (at which col index in the excel sheet, the table starts)
    title : str
    title_style : Format (xlsxwriter.format object) style applied to the title
    display_head : bool (either display the head or don't)
    head_style : Format (xlsxwriter.format object) style applied to the head
    display_index : bool (either display the index or don't)
    index_style : Format (xlsxwriter.format object) style applied to the index
    shape_style : Format (xlsxwriter.format object) style applied to the shape
    to_xls_table: bool (either the table should be an excel table or not)
    --------

    All the above parameters will be treated in "__prepare()" method,
    and that last returns a dict holding every element with there coordinates:
    ex. {"head": {
                    "start": (1, 0), # row 1 col 0 (the presence of title)
                    "end": (1, 3), # row 1 col 3
                    "data": ["2010","2011","2012","2013"]
                    "style": {"border": 1, "align": "center"}
                    "merge": False # this is used for title only
                }
         "title": {
                    "start": (0, 0), #  row 0 col 0
                    "end": (0, 3), # row 0 col 3
                    "data": "Financial income"
                    "style": {"bold": True, "align": "center", "font_size":16}
                    "merge": True
                }
        "shape": {
                    "start" : (2, 0) : presence of title and header, the start row become 2
                    "end" : (9,3) : shape size is (8x4)
                    "data" : [[...], [...]] a (8x4) data
                    "style" : {"reading_order": 1,"align": "center"}
                    "merge" : False
        }

    """

    def __init__(self,
                 df: DataFrame,
                 start_row: int,
                 start_col: int,
                 title: str = None,
                 title_style: Format = None,
                 display_head: bool = False,
                 head_style: Format = None,
                 display_index: bool = False,
                 index_style: Format = None,
                 shape_style: Format = None,
                 to_xls_table: bool = False,
                 ):

        if type(df) is not DataFrame:
            raise TypeError("Data must be a valid DataFrame parsed by pandas library !")

        if shape_style and type(shape_style) is not Format:
            raise TypeError("Invalid style format !")

        if head_style and type(head_style) is not Format:
            raise TypeError("Invalid style format !")

        if index_style and type(index_style) is not Format:
            raise TypeError("Invalid style format !")

        self.df: DataFrame = df
        self.start_row: int = start_row
        self.start_col: int = start_col
        self.title: str = title
        self.title_style: Format = title_style
        self.display_head: bool = display_head
        self.head_style: Format = head_style
        self.display_index: bool = display_index
        self.index_style: Format = index_style
        self.shape_style: Format = shape_style
        self.to_xls_table: bool = to_xls_table

        self.__shape_rows, self.__shape_cols = self.df.shape
        self.cr: List[int] = [self.start_row, self.start_col]

        self.coordinates: Dict = self.__prepare()

    def __prepare(self) -> Dict:

        res: Dict = {}

        if (self.display_index and self.display_head) or (self.display_index and self.title):
            self.cr[1] += 1

        elif self.display_index:
            self.cr[1] += 1

        if self.title:

            start_title: tuple = (self.cr[0], self.cr[1])
            end_title: tuple = (start_title[0], start_title[1] + self.__shape_cols-1)
            self.cr[0] += 1

            res["title"]: Dict = {
                "start": start_title,
                "end": end_title,
                "data": self.title,
                "style": self.title_style if self.title_style else {},
                "merge": True
            }

        if self.display_head:

            start_head: tuple = (self.cr[0], self.cr[1])
            end_head: tuple = (start_head[0], start_head[1] + self.__shape_cols-1)
            self.cr[0] += 1
            res["head"]: Dict = {
                "start": start_head,
                "end": end_head,
                "data": [list(self.df.columns)],
                "style": self.head_style if self.head_style else {},
                "merge": False
            }

        if self.display_index:

            reshaped_index = [[self.df.index[idx]] for idx in range(0, len(self.df.index))]

            start_index: tuple = (self.cr[0], self.start_col)
            end_index: tuple = (start_index[0] + self.__shape_rows-1, start_index[1])
            res["index"]: Dict = {
                "start": start_index,
                "end": end_index,
                "data": reshaped_index,
                "style": self.index_style if self.index_style else {},
                "merge": False
            }

        shape_start: tuple = (self.cr[0], self.cr[1])
        shape_end: tuple = (shape_start[0] + self.__shape_rows - 1, self.cr[1] + self.__shape_cols - 1)

        res["shape"]: Dict = {
            "start": shape_start,
            "end": shape_end,
            "data": self.df.values,
            "style": self.shape_style if self.shape_style else {},
            "merge": False
        }

        self.cr = shape_end
        return res


class VirtualSheet(Worksheet):

    def __init__(self):
        super(VirtualSheet, self).__init__()
        self.virtual_tables: List[VirtualTable] = []

    def add_virtual_table(self, table: VirtualTable):

        if type(table) is not VirtualTable:
            raise TypeError("Invalid data frame type, expected VirtualTable !")

        self.virtual_tables.append(table)

    def coordinates_writer(self, coordinates: Dict):

        start = coordinates.get("start")
        end = coordinates.get("end")
        data = coordinates.get("data")
        style = coordinates.get("style")
        merge = coordinates.get("merge")

        s_format = style if style else {}
        if merge:
            self.merge_range(start[0], start[1], end[0], end[1], data, s_format)
            return

        try:
            data.__getattribute__("__iter__")
        except:
            raise ValueError("Data must be iterable !")

        data_row_counter = 0

        for i in range(start[0], end[0]+1):
            data_col_counter = 0
            for j in range(start[1], end[1]+1):
                cell = data[data_row_counter][data_col_counter]
                self.write(i, j, cell, s_format)
                data_col_counter += 1

            data_row_counter += 1

        return
    
    def table_writer(self, table: VirtualTable):
        if "head" not in table.coordinates.keys():
            raise TypeError("Header is required to build as xls table !")
        f_row, f_col = table.coordinates.get("head").get("start")
        l_row, l_col = table.coordinates.get("shape").get("end")

        self.add_table(f_row, f_col, l_row, l_col)

        for ele_name, ele in table.coordinates.items():
            self.coordinates_writer(ele)

    def build(self):
        for v_table in self.virtual_tables:
            if v_table.to_xls_table:
                self.table_writer(v_table)
            else:
                for ele_name, coordinates in v_table.coordinates.items():
                    self.coordinates_writer(coordinates)


class WorkBookManager(Workbook):

    def __init__(self, file: str = None, options: Dict = None) -> None:
        super(WorkBookManager, self).__init__(file, options)

    def build_all(self) -> None:

        for sheet in self.worksheets():
            sheet.build()

        self.close()
