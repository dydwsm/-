# 为了处理midas中数据而学习Python


import pandas as pd
from datetime import datetime
import os
import shutil
from main_for_hanyangce import earthquake_response, earthquakeAnddeadload_response_dundi

file_name = "白沙洲引桥汉阳侧引桥50+75+50m-公路墩-0803.xlsx"
sExcelFile = "./resources/" + file_name
# sExcelFile="./resources/1130-7#墩-2.0m--实心墩.xlsx"
df = pd.read_excel(sExcelFile, sheet_name='Sheet1')
str_time = datetime.now().strftime("%Y_%m%d_%H_%M_%S")
str_contents = "./resources/{}+{}".format(file_name, str_time)
os.makedirs(str_contents)
sExcelFile_2 = str_contents + "/" + file_name
shutil.copy(sExcelFile, sExcelFile_2)


def main():
    ###提取不重复的数据
    # a = df.drop_duplicates(subset=['单元'],keep='first')
    # print(a)
    # 把不重复d元素转换成list:
    b = df['单元'].drop_duplicates().values.tolist()
    print(b)
    print(type(b))
    print(len(b))
    # b1为墩顶，b2为墩底，b3为承台底
    b1 = []
    b2 = []
    for i in range(len(b)):
        if i % 2 == 0:
            b1.append(b[i])
        else:
            b2.append(b[i])
    # b_list为实际墩号的种类，数量
    b_list = []
    for i in range(len(b1)):
        b_list.append([b1[i], b2[i]])
    print(b_list)
    print(len(b_list))
    ######################################################################################
    ######################################################################################
    """
    下面这一段代码的思路是:字典的key作为sheet名称也就是桩号，value作为表格的内容。
    写入到Excel中也是通过字典来实现的
    """
    pier_num = ['03#墩', '04#墩', '05#墩', '03#墩', '9#墩', '10#墩', '11#墩', '12#墩', '13#墩', '14#墩']
    dfs = {}
    filtered_df = []
    for i in range(len(b_list)):
        # print(i)
        condition1 = df["单元"] == b_list[i][0]
        condition2 = df["单元"] == b_list[i][1]

        filtered_df.append(df[(condition1) | (condition2)])
        condition1 = filtered_df[i]['单元'] == b_list[i][0]
        condition2 = filtered_df[i].index % 2 == 0
        # 使数据按照ijj的顺序排列
        filtered_df[i] = filtered_df[i][condition1 & condition2 | ~condition1 & ~condition2]
        dfs[pier_num[i]] = filtered_df[i]
    # print(dfs)
    print(filtered_df[0].head())
    print(len(filtered_df))
    print(type(filtered_df[0]))
    ######################################################################################
    ######################################################################################
    """
    下面要实现的内容：循环处理每一个sheet
    """
    # 输出E1和E2地震响应两张表格
    E1 = 'E1'
    E2 = 'E2'
    earthquake_response(b_list, E1, filtered_df, str_contents, pier_num)
    earthquake_response(b_list, E2, filtered_df, str_contents, pier_num)

    # 输出墩底的内力响应表格
    earthquakeAnddeadload_response_dundi(b_list, E1, filtered_df, str_contents, pier_num)
    earthquakeAnddeadload_response_dundi(b_list, E2, filtered_df, str_contents, pier_num)


if __name__ == '__main__':
    main()
