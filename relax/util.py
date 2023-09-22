from xlsxwriter.worksheet import Worksheet
import json
import os


def get_menu_list() -> tuple[list, list, int, str]:
    with open(os.path.join("base", "menu.json"), mode="r", encoding="utf8") as f:
        lst: list[dict] = json.load(f)
        menus = []
        current_index = 0
        current_key = 0
        for i, value in enumerate(lst):
            for k, v in value.items():
                if k == "value":
                    menus.append(v)
                elif k == "checked" and value["checked"]:
                    current_index = i
                    current_key = value["key"]
        return lst, menus, current_index, current_key


def write_menu_list(raw: dict):
    with open(os.path.join("base", "menu.json"), mode="w", encoding="utf8") as f:
        json.dump(raw, f, ensure_ascii=False)


def get_base_data() -> list[dict]:
    with open(os.path.join("base", "base.json"), mode="r", encoding="utf8") as f:
        dic = json.load(f)
        return dic


def get_template_data() -> list[dict]:
    with open(
        os.path.join("base", "base_template.json"), mode="r", encoding="utf8"
    ) as f:
        dic = json.load(f)
        return dic


def write_base_data(raw: dict):
    with open(os.path.join("base", "base.json"), mode="w", encoding="utf8") as f:
        json.dump(raw, f, ensure_ascii=False)


def get_current_data(raw: dict, index: int) -> dict:
    return raw[f"{index}"]


def get_row_height_byrow(
    row_count: int,
    height1: int = 24,
    height2: int = 7,
    height3: int = 16,
) -> int:
    total_height = 0
    match row_count:
        case 0 | 1:
            total_height = height1
        case 2:
            total_height = height1 + height2
        case 3:
            total_height = height1 + height2 + (row_count - 2) * height3
        case _:
            total_height = height1
    return total_height


def get_row_count(raw: str, byte_per_row: int = 19) -> int:
    raw_count = int(int(len(raw.encode("utf8")) * 2 / 3) / byte_per_row)
    if raw_count % byte_per_row != 0:
        raw_count += 1
    return raw_count


def get_row_height(
    raw: str,
    byte_per_row: int = 19,
    height1: int = 24,
    height2: int = 7,
    height3: int = 16,
):
    row_count = get_row_count(raw, byte_per_row)
    row_height = get_row_height_byrow(row_count, height1, height2, height3)
    return row_height


def stamp(ws1: Worksheet, row_height_list: list, img_path, column, page_height=760):
    page_count = 1
    total_height = 0
    break_list = []
    current_height = 0
    for height in row_height_list:
        total_height += height

    i = 0
    row_count = len(row_height_list)
    while i < row_count:
        v = row_height_list[i]
        current_height += v
        if current_height > page_height * page_count:
            current_height -= v
            i -= 1
            break_list.append(i)
            page_count += 1
        i += 1

    # 保证最后一页，至少有两行数据
    rest_height = total_height - current_height
    if rest_height == row_height_list[-1]:
        if break_list:
            break_list.pop()
        break_list.append(row_count - 2)

    dic_img = {"x_scale": 4.5 / 4.73, "y_scale": 3 / 3.01, "y_offset": 2}
    ws1.insert_image(
        f"{column}2",
        img_path,
        dic_img,
    )

    for i in break_list:
        ws1.insert_image(
            f"{column}{i+1}",
            img_path,
            dic_img,
        )
    ws1.set_h_pagebreaks(break_list)
