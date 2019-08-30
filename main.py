# -*- coding:utf-8 -*-
import tkinter as tk
import tkinter.filedialog
from tkinter.messagebox import showinfo, askquestion
from tkinter.ttk import Treeview
from utils import *

product_data = None
box_data = {}
is_ignore = None
fill_rate = 0
user_selected_box = []


def show_info(content):
    """
    弹窗提示
    """
    showinfo('提示', content)


def delButton(tree):
    """清空表格控件所有内容"""
    x = tree.get_children()
    for item in x:
        tree.delete(item)


def check_table_format(table_head):
    """
    检查表头是否符合要求
    :param table_head: 表头
    :return:
    """
    msg = ""
    flag = False
    if "订单号" not in table_head:
        msg += "表格中缺少【订单号】字段；"
        flag = True
    if "商品条码" not in table_head:
        msg += "表格中缺少【商品条码】字段；"
        flag = True
    if "数量" not in table_head:
        msg += "表格中缺少【数量】字段；"
        flag = True

    if flag:
        raise Exception(msg)


def calculation(file):
    data = readExcel(file)

    table_head = data['table_head']
    table_body = data["table_body"]

    check_table_format(table_head)

    order_index = table_head.index("订单号")
    item_index = table_head.index("商品条码")
    number_index = table_head.index("数量")
    box_statistics = {}
    orders = {}

    for item in table_body:
        if item[order_index] not in orders:
            orders[item[order_index]] = {"商品": [(item[item_index], int(item[number_index]))]}
        else:
            orders[item[order_index]]['商品'].append((item[item_index], int(item[number_index])))

    for i in orders:
        box_name = ""
        items_max_length = []  # 所有商品最长边的集合
        items_volume = 0  # 商品体积和
        box_volume = 0
        remarks = ""
        item_set = ""
        for c in orders[i]["商品"]:
            # 遍历订单中所有商品
            if product_data.get(c[0]):
                item_parameter = product_data.get(c[0])
            else:

                if is_ignore:
                    item_parameter = product_data.get("默认的计算体积")
                    if not item_parameter:
                        raise Exception("商品资料中不存在[默认的计算体积],请补充记录".format(c[0]))
                    remarks += "{0}无体积资料".format(c[0])
                else:

                    dic = set()
                    for dd in orders:
                        for sp in orders[dd]["商品"]:
                            if not product_data.get(sp[0]):
                                dic.add(sp[0])
                    info = ",".join(dic)
                    raise Exception("运算失败,以下商品无长宽高数据:\n{}".format(info))

            max_length = max(item_parameter["length"], item_parameter["width"], item_parameter["height"])  # 当前物体最长的边
            volume = item_parameter["volume"] * c[1]
            items_max_length.append(max_length)
            items_volume += volume
            item_set += r"{}/{};".format(c[0], c[1])

        flag = False
        for box in box_data:
            # fill_space = box["volume"] * ((float(fill_rate)) / 100)
            fill_space = items_volume * ((100 + int(fill_rate)) / 100)

            if box["box_type"] in user_selected_box:

                if box["volume"] > fill_space:
                    # 箱子的容积 大于 订单总体积* (100+ 填充率)%
                    if max(items_max_length) < max(box['diagonal']):
                        # 不会有超过箱子边长的物品放入
                        box_volume = box["volume"]
                        box_name = box['name']

                        if not box_statistics.get(box['name']):
                            box_statistics[box['name']] = 1
                        else:
                            box_statistics[box['name']] += 1
                        flag = True
                        break

        if not flag:
            box_name = "没有找到合适的箱子(订单体积超大)"
        orders[i]["箱型"] = box_name
        orders[i]["箱型体积"] = box_volume if box_volume else "N/A"
        orders[i]["商品体积"] = items_volume
        orders[i]["装箱率"] = items_volume / box_volume if box_volume else "N/A"
        orders[i]["备注"] = remarks
        orders[i]['商品'] = item_set

    table_title = ["订单号", "商品", "箱型", "商品体积", "箱型体积", "装箱率", "备注"]
    content = []
    for k, v in orders.items():
        tmp = [k]
        for i in table_title[1:]:
            tmp.append(v[i])  # 依次加入其它字段
        content.append(tmp)

    content.insert(0, table_title)

    file_path = writeExcel(content)
    result = {
        "box_statistics": box_statistics,
        "file_path": file_path
    }
    return result


class App:
    def __init__(self, root):
        self.root = root
        self.ignore = tk.IntVar()
        self.ignore.set(1)

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        ww = 420
        wh = 600
        self.x = (sw - ww) / 2
        self.y = (sh - wh) / 2
        root.geometry("%dx%d+%d+%d" % (ww, wh, self.x, self.y))
        root.title('箱型推荐')
        root.resizable(False, False)
        frame = tk.Frame(root)
        frame.grid(row=0, column=0, sticky=tk.E)

        self.la = tk.Label(frame, text='预留空间(百分比)')
        self.la.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.fill_rate = tk.Entry(frame, show=None)
        self.fill_rate.insert(0, 10)
        self.fill_rate.grid(row=0, column=1, padx=5, pady=5)

        self.c1 = tk.Checkbutton(frame, text='缺资料商品不进行提示,以默认体积计算', variable=self.ignore, onvalue=1, offvalue=0)
        self.c1.grid(row=1, column=1, padx=5, pady=5)

        self.btn_openFile = tk.Button(frame, text="选择数据表", fg="blue", command=self.open_file)
        self.btn_openFile.grid(row=3, column=1, padx=5, pady=5)

        description = """计算逻辑:  1. [ 箱体积 ] 大于 [ 订单商品总体积 * (100+预留空间)% ]
               2. [ 订单内所有商品的长宽高 ] 均小于 [ 箱子最长的对角线 ]
               3. 推荐满足以上条件的最小箱型
        """
        self.la = tk.Label(frame, text=description, fg="red", justify='left')
        self.la.grid(row=4, column=0, columnspan=6, rowspan=3, padx=5, pady=5)

        self.box_statistics = Treeview(frame, height=18, show="headings", columns=("箱型", "数量"))

        self.box_statistics.heading("箱型", text="箱型")  # 显示表头
        self.box_statistics.heading("数量", text="数量")

        self.box_statistics.grid(row=7, column=1, padx=5, pady=5)

    def open_file(self):
        tmp_fill_rate = self.fill_rate.get()
        tmp_ignore = self.ignore.get()

        if tmp_fill_rate.isdigit():
            if int(tmp_fill_rate) > 99 or int(tmp_fill_rate) < 0:
                show_info("填充率必须是0~99的数字")
                return
        else:
            show_info("填充率必须是0~99的数字")
            return

        filename = tkinter.filedialog.askopenfilename(title='选择订单明细表', filetypes=[('Excel', '*.xlsx')])

        if filename != '':
            global product_data
            global box_data
            global is_ignore
            global fill_rate

            try:
                product_data = get_setting("商品")
                box_data = get_setting("箱型")
            except Exception as err:
                show_info("{}-{}".format(err, "读取配置时出现异常"))
                return

            # 弹出选择箱型品牌的窗体
            frm = tk.Toplevel()
            frm.title('选择箱型模板')
            frm.geometry("%dx%d+%d+%d" % (200, 300, self.x, self.y))
            frm.resizable(False, False)

            box_type = [i["box_type"] for i in box_data]

            txt = tk.Label(frm, text='请选择箱型模板:')
            txt.pack()

            v = tk.StringVar()
            selected_box = tk.Listbox(frm, selectmode=tk.MULTIPLE,
                                      listvariable=v)  # 这里还有一个selectmode选项，默认是BROWSE（单选,拖动鼠标或方向键可以改变选项），
            selected_box.pack()
            for i in set(box_type):
                selected_box.insert(tk.END, i)

            btn_ok = tk.Button(frm, text="开始计算", fg="blue",
                               command=lambda: self.calculation(filename, tmp_fill_rate, tmp_ignore, frm, eval(v.get()),
                                                                selected_box.curselection()))
            btn_ok.pack()

    def calculation(self, filename, tmp_fill_rate, tmp_ignore, frm, v, select):
        global product_data
        global box_data
        global is_ignore
        global fill_rate
        global user_selected_box

        user_selected_box = []
        for i in select:
            user_selected_box.append(v[i])

        if not user_selected_box:
            show_info("请至少选择一种品牌的箱型")
            return

        is_ignore = tmp_ignore
        fill_rate = tmp_fill_rate
        try:
            result = calculation(filename)
            ask = askquestion('提示', '计算完成,是否打开表格?')
            delButton(self.box_statistics)
            sumV = 0
            d = result["box_statistics"]
            res = sorted(d.items(), key=lambda d: d[1], reverse=True)
            i = 0
            for item in res:
                self.box_statistics.insert('', i, values=(item[0], item[1]))
                sumV += item[1]
                i += 1
            self.box_statistics.insert('', i + 1, values=("合计", sumV))

            if ask == "yes":
                os.startfile(result["file_path"])

        except Exception as err:
            show_info(err)
        finally:
            frm.destroy()


root_windows = tk.Tk()
app = App(root_windows)
# 开始主事件循环
root_windows.mainloop()
