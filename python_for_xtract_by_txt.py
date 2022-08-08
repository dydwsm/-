"""
该py文件是操作txt来实现操作xtract软件.
主要实现功能：
批量在指定截面输入荷载
"""

import pandas as pd
import os
from datetime import datetime
import shutil


def section_loading(lines_list, section_name, bending_moment_direction, loading_name, loading_value):
    """
    处理lines_list,向其中指定的一行前插入指定的内容
    :param lines_list: 要处理的文本，列表形式
    :param section_name: 加载的截面名称
    :param bending_moment_direction: 加载的弯矩方向
    :param loading_name: 荷载工况名称
    :param loading_value: 荷载（轴力）大小
    :return: 返回修改的lines_list
    """
    index_list_1 = []
    loading_str = """
            Begin_Loading
            NAME = {}
            TYPE = Moment Curvature
            # Constant loads applied at first step - negative is read as compression.
            ConstAxial = {}

            # Incrementing load parameters - Positive increments in a positive direction.
            {} = .1130

            Use_Best_Fit = False

            # Include Plastic Hinge length.
            Calc_Moment_Rot = False

            # Analysis Parameters.
            Method = BiSection
            N_Steps_Before_Yield = 10
            N_Steps_After_Yield = 20
            Multiple_On_First_Yield = 2
            BS_Tol = 4.448
            BS_Max_Itter = 40


            End_Loading\n
            """.format(loading_name, loading_value, bending_moment_direction)
    for i in range(len(lines_list)):
        if 'Begin_Section' in lines_list[i] \
                and 'Begin_Builder' in lines_list[i + 1] \
                and section_name in lines_list[i + 2]:
            # print('第 %d 行出现' % i)
            # print(lines_list[i + 2])

            for j in range(len(lines_list[i:])):
                if 'End_Section' in lines_list[i:][j]:
                    # print('第 %d 行就是荷载要插入的位置' % (i + j))
                    index_list_1.append(i + j)  # 显然，这个列表里面第一个值就是要求的
                    # print(index_list_1)
    lines_list.insert(index_list_1[0], loading_str)
    return lines_list


def batch_rename(file_path, new_ext):
    """
    修改指定文件的后缀名
    :param file_path: 原文件路径
    :param new_ext: 新后缀名
    :return: 原目录下新文件 和新文件路径
    """
    file_dir = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)

    # list_file = os.listdir(file_dir)  # 返回指定目录
    ext = os.path.splitext(file_name)  # 返回文件名和后缀，列表形式
    str_time = datetime.now().strftime("%m%d_%H%M_%S")
    file_name_new = ext[0] + str_time + new_ext
    os.rename(
        os.path.join(file_dir, file_name),
        os.path.join(file_dir, file_name_new),
    )
    return file_name_new


def main():
    file_path = "./resources/python_for_xtract.txt"
    file_path2 = "./resources/修改后.txt"
    excel_path = "resources/白沙洲引桥汉阳侧引桥50+75+50m-铁路墩-0804.xlsx+2022_0804_14_42_34/恒载+E2墩底地震响应+04_14_42.xlsx"

    df = pd.read_excel(excel_path, sheet_name='Sheet1', index_col=0)
    print(df)
    print(df.index)
    for i in range(len(df.index)):
        print(i)

    """
    读入文件，写入文件。
    逐行读入成一个列表，然后逐行写入f2中
    """

    section_name_dic = {
        "03#墩": "1.8x3.5-2x36x",
        "04#墩": "2.8x3.5-2x32x",
        "05#墩": "2.8x3.5-2x32x",
    }

    with open(file_path, 'r') as f1, open(file_path2, 'w', encoding='ANSI') as f2:
        lines_list = f1.readlines()
        for i in df.index:
            section_name = section_name_dic[df["单元"][i]]

            if "H(最大)" in df["荷载"][i]:
                bending_moment_direction = "IncMyy"
            elif "Z(最大)" in df["荷载"][i]:
                bending_moment_direction = "IncMxx"

            str_list = [df["单元"][i][:-1], "-", df["荷载"][i][3:5]]
            loading_name = ''.join(str_list)

            loading_value = df['竖向Pz'][i] * (-1)

            # print(section_name)
            # print(bending_moment_direction)
            # print(loading_name)
            # print(loading_value)
            lines_list = section_loading(lines_list, section_name, bending_moment_direction, loading_name,
                                         loading_value)

        for i in lines_list:
            f2.write(i)
    dir_a_str = batch_rename(file_path2, ".xpj")
    print(dir_a_str)

    file_dir_xpj = os.path.dirname(file_path2) + "/" + dir_a_str
    file_dir_xpj_new = os.path.dirname(excel_path) + "/" + dir_a_str
    shutil.copy(file_dir_xpj, file_dir_xpj_new)  # 把xpj文件复制到Excel表格目录下


if __name__ == '__main__':
    main()
