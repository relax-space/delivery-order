from relax.util import get_base_data, get_current_data
from relax.write_ import start
import os


def main():
    input_file_name = os.path.join("asset", "zj3.xlsx")
    input_sheet_name = "5,6,7,8"
    input_column = "收货日期,商品名称,单位,数量,单价,金额"
    output_folder_path = "data"
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    base_data = get_base_data()
    current_data = get_current_data(base_data, "1")
    params = current_data["output"]
    start(input_file_name, input_sheet_name, input_column, output_folder_path, params)
    pass


if __name__ == "__main__":
    main()
