import os
import tkinter as tk
from tkinter import (
    Button,
    Checkbutton,
    Listbox,
    Menu,
    Scrollbar,
    Tk,
    Toplevel,
    filedialog,
    Frame,
    Label,
    messagebox,
    Entry,
)
from PIL import Image, ImageTk
from relax.util import (
    get_menu_list,
    get_base_data,
    get_current_data,
    get_template_data,
    write_base_data,
    write_menu_list,
)
import tkinter.font as tkFont
import pandas as pd
import copy

from relax.write_ import start


def center_window(root: Tk, w, h):
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = int((ws / 2) - (w / 2))
    y = int((hs / 2) - (h / 2))
    root.geometry(f"{w}x{h}+{x}+{y}")


def load_base(current_data):
    input = current_data["input"]
    _widgets["ety_column"].delete(0, tk.END)
    _widgets["ety_column"].insert(0, input["column_name"])
    _widgets["ety_sheet_exclude"].delete(0, tk.END)
    _widgets["ety_sheet_exclude"].insert(0, input["sheet_exclude"])

    output = current_data["output"]
    column_width = output["column_width"]
    for i in ["A", "B", "C", "D", "E", "F", "G"]:
        _widgets[f"col_{i}"].delete(0, tk.END)
        _widgets[f"col_{i}"].insert(0, column_width[i])
        pass

    _widgets["row1_height"].delete(0, tk.END)
    _widgets["row1_0"].delete(0, tk.END)
    _widgets["row2_height"].delete(0, tk.END)
    _widgets["row2_0"].delete(0, tk.END)
    _widgets["row2_1"].delete(0, tk.END)
    _widgets["row3_height"].delete(0, tk.END)
    _widgets["row3_0"].delete(0, tk.END)
    _widgets["row3_1"].delete(0, tk.END)
    _widgets["row4_height"].delete(0, tk.END)
    _widgets["row4_0"].delete(0, tk.END)
    _widgets["row5_height"].delete(0, tk.END)
    _widgets["row_sum_height"].delete(0, tk.END)
    _widgets["row_sum_0"].delete(0, tk.END)
    _widgets["row_last_height"].delete(0, tk.END)
    _widgets["row_last_0"].delete(0, tk.END)
    _widgets["row_last_1"].delete(0, tk.END)

    _widgets["row1_height"].insert(0, output["row1"]["height"])
    _widgets["row1_0"].insert(0, output["row1"]["contents"])
    _widgets["row2_height"].insert(0, output["row2"]["height"])
    _widgets["row2_0"].insert(0, output["row2"]["contents"][0])
    _widgets["row2_1"].insert(0, output["row2"]["contents"][1])
    _widgets["row3_height"].insert(0, output["row3"]["height"])
    _widgets["row3_0"].insert(0, output["row3"]["contents"][0])
    _widgets["row3_1"].insert(0, output["row3"]["contents"][1])
    _widgets["row4_height"].insert(0, output["row4"]["height"])
    _widgets["row4_0"].insert(0, output["row4"]["contents"][0])
    _widgets["row5_height"].insert(0, output["row5"]["height"])
    _widgets["row_sum_height"].insert(0, output["row_sum"]["height"])
    _widgets["row_sum_0"].insert(0, output["row_sum"]["contents"])
    _widgets["row_last_height"].insert(0, output["row_last"]["height"])
    _widgets["row_last_0"].insert(0, output["row_last"]["contents"][0])
    _widgets["row_last_1"].insert(0, output["row_last"]["contents"][1])

    stamp = output["stamp"]
    _enable_stamp.set(stamp["enable"])
    _widgets["ety_stamp"].delete(0, tk.END)
    _widgets["ety_stamp"].insert(0, stamp["path"])
    _widgets["ety_page_height"].delete(0, tk.END)
    _widgets["ety_page_height"].insert(0, stamp["page_height"])
    pass


def read_from_view(current_data):
    ety_sheet_exclude = _widgets["ety_sheet_exclude"].get()
    ety_file_output = _widgets["ety_file_output"].get()
    ety_column = _widgets["ety_column"].get()
    col_A = _widgets["col_A"].get()
    col_B = _widgets["col_B"].get()
    col_C = _widgets["col_C"].get()
    col_D = _widgets["col_D"].get()
    col_E = _widgets["col_E"].get()
    col_F = _widgets["col_F"].get()
    col_G = _widgets["col_G"].get()

    row1_height = _widgets["row1_height"].get()
    row1_0 = _widgets["row1_0"].get()
    row2_height = _widgets["row2_height"].get()
    row2_0 = _widgets["row2_0"].get()
    row2_1 = _widgets["row2_1"].get()
    row3_height = _widgets["row3_height"].get()
    row3_0 = _widgets["row3_0"].get()
    row3_1 = _widgets["row3_1"].get()
    row4_height = _widgets["row4_height"].get()
    row4_0 = _widgets["row4_0"].get()
    row5_height = _widgets["row5_height"].get()
    row_sum_height = _widgets["row_sum_height"].get()
    row_sum_0 = _widgets["row_sum_0"].get()
    row_last_height = _widgets["row_last_height"].get()
    row_last_0 = _widgets["row_last_0"].get()
    row_last_1 = _widgets["row_last_1"].get()

    enable_stamp = _enable_stamp.get()
    ety_stamp = _widgets["ety_stamp"].get()
    ety_page_height = _widgets["ety_page_height"].get()

    input = current_data["input"]
    input["column_name"] = ety_column
    input["sheet_exclude"] = ety_sheet_exclude

    output = current_data["output"]
    output["folder_path"] = ety_file_output

    column_width = output["column_width"]
    column_width["A"] = int(col_A)
    column_width["B"] = int(col_B)
    column_width["C"] = int(col_C)
    column_width["D"] = int(col_D)
    column_width["E"] = int(col_E)
    column_width["F"] = int(col_F)
    column_width["G"] = int(col_G)

    output["row1"]["height"] = int(row1_height)
    output["row1"]["contents"] = [row1_0]
    output["row2"]["height"] = int(row2_height)
    output["row2"]["contents"] = [row2_0, row2_1]
    output["row3"]["height"] = int(row3_height)
    output["row3"]["contents"] = [row3_0, row3_1]
    output["row4"]["height"] = int(row4_height)
    output["row4"]["contents"] = [row4_0]
    output["row5"]["height"] = int(row5_height)
    output["row_sum"]["height"] = int(row_sum_height)
    output["row_sum"]["contents"] = [row_sum_0]
    output["row_last"]["height"] = int(row_last_height)
    output["row_last"]["contents"] = [row_last_0, row_last_1]

    stamp = output["stamp"]
    stamp["enable"] = enable_stamp
    stamp["path"] = ety_stamp
    stamp["page_height"] = int(ety_page_height)


def save_template():
    lst_menu: Listbox = _widgets["lst_menu"]
    indexs = lst_menu.curselection()
    current_data = {}
    if not indexs:
        messagebox.showinfo("", "请先选择左边的模板")
        return current_data
    index: int = lst_menu.curselection()[0]
    data = get_current_data_index(index)
    if not data:
        return
    read_from_view(data)

    current_key = _menu_list[index]["key"]
    _base_data[current_key] = data
    write_base_data(_base_data)
    messagebox.showinfo("提示", "保存成功")

    pass


def render_save_btn(fr):
    btn_save = Button(fr, text="保存到模板", command=save_template)
    # btn_save.bind("<Button-1>", save_template)
    btn_save.pack(fill="x", side="left", ipadx=0)
    _widgets["btn_save"] = btn_save


def upload_file_output():
    lst_menu: Listbox = _widgets["lst_menu"]
    if not lst_menu.curselection():
        messagebox.showinfo("提示", "请先在左边选择一个模板")
        return

    file_path = filedialog.askdirectory()
    if not file_path:
        return
    ety_file_output: Entry = _widgets["ety_file_output"]
    ety_file_output.delete(0, tk.END)
    ety_file_output.insert(tk.END, file_path)
    pass


def render_column_width(fr, dic_column_width: dict, folder_path):
    lbl_1 = Label(fr, text="输出文件夹路径：*", fg="red")
    lbl_1.grid(row=0, column=0, sticky="e")
    ety_file_output = Entry(
        fr,
        width=60,
        name="ety_file_output",
    )
    ety_file_output.insert(0, folder_path)
    _widgets["ety_file_output"] = ety_file_output
    ety_file_output.grid(row=0, column=1, columnspan=8)

    btn_file = Button(fr, text="选择", command=upload_file_output)
    btn_file.grid(row=0, column=9)

    lbl2 = Label(fr, text="列宽：", width=10)
    lbl2.grid(row=1, column=0, sticky="e")
    index = 0
    for k, v in dic_column_width.items():
        lbl = Label(fr, text=f"{k}：")
        index += 1
        lbl.grid(row=1, column=index, sticky="w", padx=(0, 3))
        entry = Entry(fr, width=6, justify="right", name=f"col_{k}")
        _widgets[f"col_{k}"] = entry
        entry.insert(0, v)
        index += 1
        entry.grid(row=1, column=index, padx=(0, 3))


def render_image(fr121):
    img_path_list = [
        "base/trash.png",
        "base/edit.png",
        "base/new.png",
    ]

    def get_command(name):
        match name:
            case "base/new.png":
                return on_add
            case "base/edit.png":
                return on_edit
            case "base/trash.png":
                return on_delete
        return None

    for i in img_path_list:
        img1 = Image.open(i)
        img1 = img1.resize((20, 20))
        img1 = ImageTk.PhotoImage(img1)
        event = get_command(i)
        btn1 = Button(fr121, image=img1, borderwidth=0, command=event)
        # if event:
        #     btn1.bind("<Button-1>", event)
        btn1.pack(fill="x", side="right", pady=0)
        # 重新引用，防止被垃圾回收
        btn1.image = img1


def upload_stamp():
    file_path = filedialog.askopenfilename()
    if not file_path:
        return
    ety_stamp: Entry = _widgets["ety_stamp"]
    ety_stamp.delete(0, tk.END)
    ety_stamp.insert(0, file_path)
    pass


def stamp_click():
    if _enable_stamp:
        pass
    else:
        pass

    pass


def render_stamp(fr: Frame, dic_v: dict):
    global _enable_stamp
    enable = dic_v["enable"]
    _enable_stamp = tk.BooleanVar()
    chk = Checkbutton(
        fr, text="是否盖章", variable=_enable_stamp, name="chk_stamp", command=stamp_click
    )
    _widgets["chk_stamp"] = chk
    if enable:
        chk.select()
    chk.grid(row=0, column=0)
    lbl_remark = Label(fr, text="(如果不需要，下面的不用设置)")
    lbl_remark.grid(row=0, column=1, sticky="w")

    lbl_1 = Label(fr, text="图片路径：")
    lbl_1.grid(row=1, column=0, sticky="e")
    ety_stamp = Entry(fr, width=60, name="ety_stamp")
    _widgets["ety_stamp"] = ety_stamp
    ety_stamp.insert(0, dic_v["path"])

    ety_stamp.grid(row=1, column=1, columnspan=2, sticky="w")
    btn_stamp = Button(fr, text="选择", command=upload_stamp, name="btn_stamp")
    _widgets["btn_stamp"] = btn_stamp
    btn_stamp.grid(row=1, column=3)

    lbl_2 = Label(fr, text="页面高度：")
    lbl_2.grid(row=2, column=0, sticky="e")
    ety_page_height = Entry(fr, width=6, name="ety_page_height")
    _widgets["ety_page_height"] = ety_page_height
    ety_page_height.insert(0, dic_v["page_height"])
    ety_page_height.grid(row=2, column=1, sticky="w")
    pass


def upload_file():
    lst_menu: Listbox = _widgets["lst_menu"]
    if not lst_menu.curselection():
        messagebox.showinfo("提示", "请先在左边选择一个模板")
        return

    file_path = filedialog.askopenfilename()
    if not file_path:
        return
    ety_file: Entry = _widgets["ety_file"]
    ety_file.delete(0, tk.END)
    ety_file.insert(tk.END, file_path)
    df = pd.ExcelFile(file_path)
    sheets = df.sheet_names
    ety_sheet_exclude: Entry = _widgets["ety_sheet_exclude"]
    exclude_str: str = ety_sheet_exclude.get()
    exclude_list = exclude_str.replace("，", ",").replace(" ", "").split(",")
    in_sheets = [i for i in sheets if i not in exclude_list]
    ety_sheet: Entry = _widgets["ety_sheet"]
    ety_sheet.delete(0, tk.END)
    ety_sheet.insert(0, ",".join(in_sheets))
    pass


def start_click():
    input_file: Entry = _widgets["ety_file"]
    input_file_name = input_file.get()
    if not input_file_name:
        messagebox.showinfo("提示", "请选择文件路径")
        input_file.focus_set()
        return
    if not os.path.exists(input_file_name):
        messagebox.showinfo("提示", "文件路径不存在！")
        input_file.focus_set()
        return
    ety_sheet: Entry = _widgets["ety_sheet"]
    input_sheet_name = ety_sheet.get()
    if not input_sheet_name:
        messagebox.showinfo("提示", "请输入sheet名")
        ety_sheet.focus_set()
        return
    lst_menu: Listbox = _widgets["lst_menu"]
    if not lst_menu.curselection():
        messagebox.showinfo("提示", "请先在左边选择一个模板")
        return

    current_data = get_current_data_kong()
    if not current_data:
        return
    read_from_view(current_data)
    input_column = current_data["input"]["column_name"]
    ety_file_output: Entry = _widgets["ety_file_output"]
    output_path = ety_file_output.get()
    if output_path == "data":
        if not os.path.exists(output_path):
            os.makedirs(output_path)
    elif not os.path.exists(output_path):
        messagebox.showinfo("提示", "输出文件夹路径不存在！")
        ety_file_output.focus_set()
        return

    params = current_data["output"]
    start(input_file_name, input_sheet_name, input_column, output_path, params)
    result = messagebox.askquestion("提示", "成功，你需要打开文件夹吗？")
    if result == "yes":
        os.startfile(output_path)


def render_file(fr: Frame):
    lbl_1 = Label(fr, text="文件：*", fg="red")
    lbl_1.grid(row=0, column=0, sticky="e")
    ety_file = Entry(
        fr,
        width=60,
        name="ety_file",
    )

    _widgets["ety_file"] = ety_file
    ety_file.grid(row=0, column=1, columnspan=8)
    btn_file = Button(fr, text="选择", command=upload_file)
    btn_file.grid(row=0, column=9, ipady=0)

    lbl_2 = Label(fr, text="sheet名(多个用逗号隔开)：*", fg="red")
    lbl_2.grid(row=1, column=0, sticky="e")
    ety_sheet = Entry(fr, width=60, name="ety_sheet")
    _widgets["ety_sheet"] = ety_sheet
    ety_sheet.grid(row=1, column=1, columnspan=8)

    lbl_3 = Label(fr, text="sheet名(排除)：")
    lbl_3.grid(row=2, column=0, sticky="e")
    ety_sheet_exclude = Entry(fr, width=60, name="ety_sheet_exclude")
    _widgets["ety_sheet_exclude"] = ety_sheet_exclude
    ety_sheet_exclude.grid(row=2, column=1, columnspan=8)

    lbl_4 = Label(fr, text="提取的列名：")
    lbl_4.grid(row=3, column=0, sticky="e")
    ety_column = Entry(fr, width=60, name="ety_column")
    _widgets["ety_column"] = ety_column
    ety_column.grid(row=3, column=1, columnspan=8)

    btn_start = Button(fr, text="生成", width=15, command=start_click)
    btn_start.grid(
        row=0, column=10, rowspan=4, padx=(40, 0), sticky="W" + "E" + "N" + "S"
    )

    pass


def adjust_width(event):
    # entry.config(width=len(entry.get()))
    pass


def render_row(fr: Frame, k, v, row_index):
    s = k.replace("row", "")
    lbl_text = ""
    if s == "_sum":
        lbl_text = "合计"
    elif s == "_last":
        lbl_text = "最后一行"
    else:
        lbl_text = f"第{s}行"
        pass
    lbl = Label(fr, text=f"{lbl_text}，行高：")
    lbl.grid(row=row_index, column=0, padx=(0, 3), sticky="e")
    entry1 = Entry(fr, width=6, justify="right", name=f"{k}_height")
    _widgets[f"{k}_height"] = entry1
    entry1.grid(row=row_index, column=1, padx=(0, 3), sticky="w")
    entry1.insert(0, v["height"])
    lbl2 = Label(fr, text=f"内容：", name=f"{k}_content")
    lbl2.grid(row=row_index, column=2, padx=(0, 3), sticky="w")
    contents = v["contents"]
    column_index = 3

    for i, v2 in enumerate(contents):
        if row_index == 3:
            entry = Entry(fr, width=51, name=f"{k}_{i}")
            entry.grid(
                row=row_index,
                column=column_index,
                columnspan=2,
                padx=(0, 3),
                sticky="w",
            )
        else:
            entry = Entry(fr, width=25, name=f"{k}_{i}")
            entry.grid(row=row_index, column=column_index, padx=(0, 3), sticky="w")

        _widgets[f"{k}_{i}"] = entry
        entry.insert(0, v2)
        column_index += 1


def get_current_data_kong():
    lst_menu: Listbox = _widgets["lst_menu"]
    indexs = lst_menu.curselection()
    current_data = {}
    if not indexs:
        return current_data
    index: int = lst_menu.curselection()[0]
    return get_current_data_index(index)


def get_current_data_index(index):
    current_key = _menu_list[index]["key"]
    current_data = get_current_data(_base_data, current_key)
    return copy.deepcopy(current_data)


def left_click(e):
    data = get_current_data_kong()
    if not data:
        return
    load_base(data)


def right_click(event):
    lst_menu: Listbox = _widgets["lst_menu"]
    index = lst_menu.nearest(event.y)
    if index == -1:
        return

    current_data = get_current_data_index(index)
    load_base(current_data)

    lst_menu.select_clear(0, tk.END)
    lst_menu.select_set(index)

    right_click_menu = _widgets["right_click_menu"]
    right_click_menu.post(event.x_root, event.y_root)


def on_delete():
    lst_menu: Listbox = _widgets["lst_menu"]
    selected_indices = lst_menu.curselection()
    if not selected_indices:
        messagebox.showinfo("", "请先选择要删除的模板")
        return
    result = messagebox.askquestion("提示", "您确定要删除吗？")
    if result != "yes":
        return

    index = selected_indices[0]
    lst_menu.delete(index)

    current_key = _menu_list[index]["key"]
    del _menu_list[index]
    del _base_data[current_key]

    write_menu_list(_menu_list)
    write_base_data(_base_data)


def on_default():
    lst_menu: Listbox = _widgets["lst_menu"]
    selected_indices = lst_menu.curselection()
    if not selected_indices:
        messagebox.showinfo("", "请先选择默认的模板")
        return
    for i, _ in enumerate(_menu_list):
        _menu_list[i]["checked"] = False

    index = selected_indices[0]
    current_dict: dict = _menu_list.pop(index)
    current_dict["checked"] = True
    _menu_list.insert(0, current_dict)
    write_menu_list(_menu_list)

    lst_menu.delete(0, tk.END)
    for v in _menu_list:
        lst_menu.insert(tk.END, v["value"])

    lst_menu.select_set(0)

    pass


def on_edit():
    lst_menu: Listbox = _widgets["lst_menu"]
    selects = lst_menu.curselection()
    if not selects:
        messagebox.showinfo("", "请先选择要编辑的模板")
        return
    index = selects[0]
    default_content = lst_menu.get(tk.ACTIVE)

    def on_edit_in(new_name):
        _menu_list[index]["value"] = new_name
        write_menu_list(_menu_list)

        lst_menu.delete(index)
        lst_menu.insert(index, new_name)
        lst_menu.select_set(index)

    show_top("修改", default_content, on_edit_in)
    pass


def exist_item(name):
    for v in _menu_list:
        if v["value"] == name:
            return True
    return False


def show_top(tip, default_content, on_add_in):
    popup = Toplevel(root)
    popup.title(tip)
    center_window(popup, 200, 100)
    popup.resizable(False, False)
    lbl_msg = Label(popup, text="", fg="red")
    lbl_msg.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0))

    ety = Entry(popup)
    ety.insert(0, default_content)
    ety.focus_set()
    ety.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10))

    def comfirm(on_add_in):
        new_name = ety.get()
        if not new_name:
            messagebox.showinfo("提示", "不能为空")
            return
        is_has = exist_item(new_name)
        if is_has:
            lbl_msg.config(text="存在同样名字的模板")
            return
        on_add_in(new_name)
        popup.destroy()

    btn1 = Button(popup, text="确定", command=lambda: comfirm(on_add_in))
    btn1.grid(row=2, column=0, padx=(20, 0), pady=(0, 10), ipadx=10)
    btn2 = Button(popup, text="取消", command=popup.destroy)
    btn2.grid(row=2, column=1, padx=(0, 20), pady=(0, 10), ipadx=10)


def on_add():
    def on_add_in(new_name):
        max = -1
        for v in _menu_list:
            current = int(v["key"])
            if current > max:
                max = current
        key = max + 1
        checked = True if key == 0 else False
        menu = {"key": f"{key}", "value": new_name, "checked": checked}
        lst_menu: Listbox = _widgets["lst_menu"]
        lst_menu.insert(tk.END, new_name)
        lst_menu.select_clear(0, tk.END)
        lst_menu.select_set(tk.END)

        current_data = get_template_data()
        load_base(current_data)

        _menu_list.append(menu)
        write_menu_list(_menu_list)

        _base_data[f"{key}"] = current_data
        write_base_data(_base_data)
        scroll_to_bottom()

    show_top("新增", "", on_add_in)
    pass


def scroll_to_bottom():
    lst_menu: Listbox = _widgets["lst_menu"]
    lst_menu.see(tk.END)


def render_listbox(fr):
    sb_lst_menu = Scrollbar(fr)
    _widgets["sb_lst_menu"] = sb_lst_menu
    sb_lst_menu.pack(side=tk.RIGHT, fill=tk.Y)

    lst_menu = Listbox(
        fr, exportselection=False, yscrollcommand=sb_lst_menu.set, name="lst_menu"
    )
    sb_lst_menu.config(command=lst_menu.yview)
    _widgets["lst_menu"] = lst_menu
    lst_menu.bind("<<ListboxSelect>>", left_click)
    lst_menu.bind("<Button-3>", right_click)
    lst_menu.pack(fill="y", expand=True)

    right_click_menu = Menu(fr, tearoff=0)
    right_click_menu.add_command(label="设置为默认", command=on_default)
    right_click_menu.add_command(label="添加", command=on_add)
    right_click_menu.add_command(label="重命名", command=lambda: on_edit())
    right_click_menu.add_command(label="删除", command=on_delete)
    _widgets["right_click_menu"] = right_click_menu


def init_view():
    root.title("送货单")
    center_window(root, 900, 600)
    bold_font = tkFont.Font(size=10, weight="bold")
    # 1.上半部分
    fr2 = Frame(root, bd=1, name="fr2")
    fr2.pack(fill="both", padx=3, pady=3)

    fr21 = Frame(fr2, height=20, bd=1, name="fr21")
    _widgets["fr21"] = fr21
    fr21.pack(fill="x", side="top", padx=3, pady=3)
    lbl_input = Label(fr21, text="输入excel配置", font=bold_font)
    lbl_input.pack(side="left", ipadx=3, ipady=3, padx=3, pady=3)

    fr_input_sep = Frame(
        fr2, height=2, borderwidth=1, relief="groove", name="fr_input_sep"
    )
    _widgets["fr_input_sep"] = fr_input_sep
    fr_input_sep.pack(fill="x", padx=3, pady=(0, 5))

    fr22 = Frame(fr2, bd=1, name="fr22")
    _widgets["fr22"] = fr22
    fr22.pack(fill="both")
    render_file(fr22)

    # 2.下半部分
    fr1 = Frame(root, width=100, bd=1, name="fr1")
    _widgets["fr1"] = fr1
    fr1.pack(fill="both", padx=3, pady=10)

    fr11 = Frame(fr1, height=20, bd=1, name="fr11")
    _widgets["fr11"] = fr11
    fr11.pack(fill="x", side="top", padx=3, pady=3)
    lbl_base = Label(fr11, text="输出excel配置", font=bold_font)
    lbl_base.pack(side="left", ipadx=3, ipady=3, padx=3, pady=3)

    fr_output_sep = Frame(
        fr1, height=2, borderwidth=1, relief="groove", name="fr_output_sep"
    )
    _widgets["fr_output_sep"] = fr_output_sep
    fr_output_sep.pack(fill="x", padx=3, pady=(0, 5))

    # 1.1 左侧listbox
    fr12 = Frame(fr1, relief="groove", width=100, bd=1, name="fr12")
    _widgets["fr12"] = fr12
    fr12.pack(fill="y", side="left")

    fr121 = Frame(fr12, height=20, bd=1, name="fr121")
    _widgets["fr121"] = fr121
    fr121.pack(fill="x")
    lbl_template = Label(fr121, text="模板")
    lbl_template.pack(fill="x", side="left")
    render_image(fr121)

    fr122 = Frame(fr12, height=20, bd=1, name="fr122")
    _widgets["fr122"] = fr122
    fr122.pack(fill="both")
    render_listbox(fr122)

    # 右侧内容
    fr13 = Frame(fr1, relief="groove", height=20, bd=1, name="fr13")
    _widgets["fr13"] = fr13
    fr13.pack(fill="both", padx=3, pady=3)

    fr131 = Frame(fr13, height=20, bd=1, name="fr131")
    _widgets["fr131"] = fr131
    fr131.pack(fill="x")
    render_save_btn(fr131)

    # 设置列宽
    fr132 = Frame(fr13, height=20, bd=1, name="fr132")
    _widgets["fr132"] = fr132
    fr132.pack(fill="x")

    # 设置行高，以及一些固定值
    fr134 = Frame(fr13, height=20, bd=1, name="fr134")
    _widgets["fr134"] = fr134
    fr134.pack(fill="x")

    # 图章
    fr135 = Frame(fr13, height=20, bd=1, name="fr135")
    _widgets["fr135"] = fr135
    fr135.pack(fill="x")


def init_data():
    global _menu_list
    global _base_data

    _menu_list, _menu_pure_list, current_index, current_key = get_menu_list()
    _base_data = get_base_data()
    current_data = copy.deepcopy(get_current_data(_base_data, current_key))
    # 设置listbox
    lst_menu: Listbox = _widgets.get("lst_menu")
    for i in _menu_pure_list:
        lst_menu.insert(tk.END, i)

    # 设置列、行 和 图章
    fr132 = _widgets.get("fr132")
    fr134 = _widgets.get("fr134")
    fr135 = _widgets.get("fr135")
    row_index = 0
    for k, v in current_data.items():
        if k == "input":
            for k1, v1 in v.items():
                key: str = k1
                if key == "column_name":
                    ety_column: Entry = _widgets.get("ety_column")
                    ety_column.insert(0, v1)
                elif key == "sheet_exclude":
                    ety_sheet_exclude: Entry = _widgets.get("ety_sheet_exclude")
                    ety_sheet_exclude.insert(0, v1)
        elif k == "output":
            for k2, v2 in v.items():
                key: str = k2
                if key == "column_width":
                    render_column_width(fr132, v2, v["folder_path"])
                elif key.startswith("row"):
                    render_row(fr134, k2, v2, row_index)
                    row_index += 1
                elif key == "stamp":
                    render_stamp(fr135, v2)
                    pass

    lst_menu.select_set(current_index)
    load_base(current_data)
    pass


def init():
    global root
    root = Tk()
    global _widgets
    _widgets = {}
    init_view()
    init_data()
    root.mainloop()


if __name__ == "__main__":
    init()
