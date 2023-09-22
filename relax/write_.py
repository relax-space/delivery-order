import os
import pandas as pd
from datetime import datetime, timedelta
from xlsxwriter import Workbook
from xlsxwriter.format import Format
from enum import Enum

from relax.util import get_row_height, stamp


class CssFormatKey(Enum):
    Row1 = 1
    Row2_title = 2
    Row2_value = 3
    Content_title = 4
    Content_1 = 5
    Content_2 = 6
    Sum_title = 7
    Sum_value = 8
    Row_last = 9


def start(
    input_file_name: str,
    input_sheet_name: str,
    input_column: str,
    output_folder_path,
    params: dict,
):
    sheet_name_list = input_sheet_name.replace("，", ",").split(",")
    column_list = input_column.replace("，", ",").split(",")
    df = read_all(input_file_name, sheet_name_list, column_list)
    write_all(df, output_folder_path, params)
    pass


def read_all(xlsx_path: str, sheet_name_list: list, column_list: list) -> pd.DataFrame:
    if len(column_list) != 6:
        return None
    rename_list = {
        column_list[0]: "A",
        column_list[1]: "B",
        column_list[2]: "C",
        column_list[3]: "D",
        column_list[4]: "E",
        column_list[5]: "F",
    }
    dic_df = pd.read_excel(
        xlsx_path,
        usecols=column_list,
        dtype={column_list[0]: str},
        sheet_name=sheet_name_list,
    )
    df = pd.concat(list(dic_df.values()), ignore_index=True)
    df = df.rename(columns=rename_list)
    return df


def write_all(df: pd.DataFrame, output_folder_path, params: dict):
    df["A"] = df["A"].apply(lambda x: x.replace(" 00:00:00", ""))
    date_set = set(df["A"].values)
    for v in date_set:
        # if v != "2023-07-17":
        #     continue
        df2 = df.query("A == @v")
        df2 = df2.reset_index(drop=True)
        write_one(df2, output_folder_path, v, params)
        pass


def write_one(df: pd.DataFrame, output_folder_path, date_str: str, params: dict):
    file_name = date_str
    wb1 = Workbook(
        os.path.join(output_folder_path, f"{file_name}.xlsx"),
        options={
            "strings_to_numbers": True,
            "constant_memory": True,
            "encoding": "utf-8",
        },
    )

    ws1 = wb1.add_worksheet()
    ws1.center_horizontally()
    # set column width
    ws1.set_column("A:A", 10.67)
    ws1.set_column("B:B", 23)
    ws1.set_column("C:C", 8.89)
    ws1.set_column("D:D", 10)
    ws1.set_column("E:E", 10.44)
    ws1.set_column("F:F", 13)
    ws1.set_column("G:G", 7.33)

    date = datetime.strptime(date_str, "%Y-%m-%d")
    d1 = (date + timedelta(days=-1)).strftime("%Y-%m-%d")
    d2 = date.strftime("%Y-%m-%d")

    row_height_list = [
        params["row1"]["height"],
        params["row2"]["height"],
        params["row3"]["height"],
        params["row4"]["height"],
    ]
    # set 4 row width and value
    s1 = params["row1"]["contents"][0]
    s2 = params["row2"]["contents"][0]
    s3 = params["row2"]["contents"][1]
    s4 = params["row3"]["contents"][0]
    s5 = params["row3"]["contents"][1]

    s6 = params["row_sum"]["contents"][0]
    s7 = params["row_last"]["contents"][0]
    s8 = params["row_last"]["contents"][1]

    ws1.set_row(0, row_height_list[0])
    ws1.merge_range(0, 0, 0, 6, s1, get_css_format(wb1, CssFormatKey.Row1, params))

    ws1.set_row(1, row_height_list[1])
    ws1.write(1, 0, s2, get_css_format(wb1, CssFormatKey.Row2_title, params))
    ws1.write(1, 1, d1, get_css_format(wb1, CssFormatKey.Row2_value, params))
    ws1.merge_range(
        1, 4, 1, 5, s3, get_css_format(wb1, CssFormatKey.Row2_title, params)
    )

    ws1.set_row(2, row_height_list[2])
    ws1.write(2, 0, s4, get_css_format(wb1, CssFormatKey.Row2_title, params))
    ws1.write(2, 1, d2, get_css_format(wb1, CssFormatKey.Row2_value, params))
    ws1.merge_range(
        2, 4, 2, 6, s5, get_css_format(wb1, CssFormatKey.Row2_title, params)
    )

    ws1.set_row(3, row_height_list[3])
    column_name_list = (
        params["row4"]["contents"][0].replace("，", ",").replace(" ", "").split(",")
    )
    for i, v in enumerate(column_name_list):
        ws1.write(3, i, v, get_css_format(wb1, CssFormatKey.Content_title, params))

    css_content1 = get_css_format(wb1, CssFormatKey.Content_1, params)
    css_content2 = get_css_format(wb1, CssFormatKey.Content_2, params)
    sum = 0
    # write content
    for i, row in df.iterrows():
        index = i + 1
        row_index = i + 4
        row_b = row["B"]
        height = get_row_height(row_b)
        row_height_list.append(height)
        ws1.set_row(row_index, height)

        row_f = row["F"]
        sum += row_f

        ws1.write(row_index, 0, index, css_content1)
        ws1.write(row_index, 1, row_b, css_content1)
        ws1.write(row_index, 2, row["C"], css_content1)
        ws1.write(row_index, 3, row["D"], css_content2)
        ws1.write(row_index, 4, row["E"], css_content2)
        ws1.write(row_index, 5, row_f, css_content2)
        ws1.write(row_index, 6, "", css_content1)

    height_end_list = [params["row_sum"]["height"], params["row_last"]["height"]]
    row_height_list.extend(height_end_list)
    # 倒数第二行
    row_index = len(df) + 4
    ws1.set_row(row_index, height_end_list[0])
    ws1.merge_range(
        row_index,
        0,
        row_index,
        4,
        s6,
        get_css_format(wb1, CssFormatKey.Sum_title, params),
    )
    ws1.write(row_index, 5, sum, get_css_format(wb1, CssFormatKey.Sum_value, params))
    ws1.write(row_index, 6, "", get_css_format(wb1, CssFormatKey.Sum_value, params))

    # 最后一行
    row_index += 1
    ws1.set_row(row_index, height_end_list[1])
    s7 = f"     {s7}：                                            {s8}："
    ws1.merge_range(
        row_index,
        0,
        row_index,
        6,
        s7,
        get_css_format(wb1, CssFormatKey.Row_last, params),
    )

    stamp_json = params.get("stamp")
    if stamp_json and stamp_json["enable"]:
        img_path = stamp_json.get("path")
        img_column = stamp_json.get("column")
        page_height = stamp_json.get("page_height")
        stamp(ws1, row_height_list, img_path, img_column, page_height)

    wb1.close()
    pass


def get_css_format(wb1: Workbook, css_format: CssFormatKey, dic_format: dict) -> Format:
    fmt = None
    match css_format:
        case CssFormatKey.Row1:
            fmt = wb1.add_format(dic_format["row1"]["formats"][0])
        case CssFormatKey.Row2_title:
            fmt = wb1.add_format(dic_format["row2"]["formats"][0])
        case CssFormatKey.Row2_value:
            fmt = wb1.add_format(dic_format["row2"]["formats"][1])
        case CssFormatKey.Content_title:
            fmt = wb1.add_format(dic_format["row4"]["formats"][0])
        case CssFormatKey.Content_1:
            fmt = wb1.add_format(dic_format["row5"]["formats"][0])
        case CssFormatKey.Content_2:
            fmt = wb1.add_format(dic_format["row5"]["formats"][1])
        case CssFormatKey.Sum_title:
            fmt = wb1.add_format(dic_format["row_sum"]["formats"][0])
        case CssFormatKey.Sum_value:
            fmt = wb1.add_format(dic_format["row_sum"]["formats"][1])
        case CssFormatKey.Row_last:
            fmt = wb1.add_format(dic_format["row_last"]["formats"][0])
    return fmt
