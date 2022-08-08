# 为了处理midas中数据而学习Python


def earthquake_response(pier_numlist, earth_style, filtered_df, str_contents, pier_num):
    # 该函数为输出地震响应内力的表格，格式为弯矩，剪力，轴力
    filtered_e1h = []
    filtered_e1z = []
    filtered_e1 = []

    dfs = {}
    if str(earth_style) == 'E1':
        earth_loading1 = 'E1H(最大)'
        earth_loading2 = 'E1Z(最大)'
    elif str(earth_style) == 'E2':
        earth_loading1 = 'E2H(最大)'
        earth_loading2 = 'E2Z(最大)'
    for i in range(len(pier_numlist)):
        condition0 = filtered_df[i]["荷载"] == earth_loading1
        condition1 = filtered_df[i]["荷载"] == earth_loading2
        filtered_e1h.append(filtered_df[i][condition0])
        filtered_e1z.append(filtered_df[i][condition1])
        filtered_e1h[i].insert(2, '弯矩-z (kN*m)', filtered_e1h[i].pop('弯矩-z (kN*m)'))
        filtered_e1h[i].insert(3, '剪力-y (kN)', filtered_e1h[i].pop('剪力-y (kN)'))
        filtered_e1h[i].insert(4, '轴向 (kN)', filtered_e1h[i].pop('轴向 (kN)'))
        filtered_e1h[i] = filtered_e1h[i].drop(axis=1, columns=['剪力-z (kN)', '扭矩 (kN*m)', '弯矩-y (kN*m)', '位置'])
        filtered_e1h[i] = filtered_e1h[i].reset_index(drop=True)
        filtered_e1h[i].index = range(filtered_e1h[i].shape[0])
        filtered_e1z[i].insert(0, '弯矩-y (kN*m)', filtered_e1z[i].pop('弯矩-y (kN*m)'))
        filtered_e1z[i].insert(1, '剪力-z (kN)', filtered_e1z[i].pop('剪力-z (kN)'))
        filtered_e1z[i].insert(2, '轴向 (kN)', filtered_e1z[i].pop('轴向 (kN)'))
        filtered_e1z[i].insert(3, '荷载', filtered_e1z[i].pop('荷载'))
        filtered_e1z[i] = filtered_e1z[i].drop(axis=1, columns=['剪力-y (kN)', '扭矩 (kN*m)', '弯矩-z (kN*m)', '位置'])
        filtered_e1z[i] = filtered_e1z[i].reset_index(drop=True)
        filtered_e1.append(pd.concat([filtered_e1h[i], filtered_e1z[i]], axis=1))
        filtered_e1[i] = filtered_e1[i].round(0)
        dfs[pier_num[i]] = filtered_e1[i]
    # str_a = './resources/{}地震响应.xlsx'.format(earth_style)
    # str_a = str_contents + '/{}地震响应.xlsx'.format(earth_style)
    str_a = str_contents + '/{}地震响应'.format(earth_style) + '+{}.xlsx'.format(datetime.now().strftime("%d_%H_%M"))
    writer = pd.ExcelWriter(str_a, engine='xlsxwriter')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()


def earthquakeAnddeadload_response(pier_numlist, earth_style, filtered_df, str_contents, pier_num):
    # 该函数为输出恒载和地震响应内力组合的表格，格式为竖向Pz,顺力Hy,横力Hx,顺弯Mx,横弯My,扭矩Mz

    filtered_e1h = []
    filtered_e1z = []
    filtered_e1 = []

    dfs = {}
    if str(earth_style) == 'E1':
        earth_loading1 = '恒+E1H(最大)'
        earth_loading2 = '恒+E1H(最小)'
        earth_loading3 = '恒-E1H(最大)'
        earth_loading4 = '恒-E1H(最小)'
        earth_loading5 = '恒+E1Z(最大)'
        earth_loading6 = '恒+E1Z(最小)'
        earth_loading7 = '恒-E1Z(最大)'
        earth_loading8 = '恒-E1Z(最小)'
    elif str(earth_style) == 'E2':
        earth_loading1 = '恒+E2H(最大)'
        earth_loading2 = '恒+E2H(最小)'
        earth_loading3 = '恒-E2H(最大)'
        earth_loading4 = '恒-E2H(最小)'
        earth_loading5 = '恒+E2Z(最大)'
        earth_loading6 = '恒+E2Z(最小)'
        earth_loading7 = '恒-E2Z(最大)'
        earth_loading8 = '恒-E2Z(最小)'
    for i in range(len(pier_numlist)):
        condition0 = filtered_df[i]["单元"] == pier_numlist[i][2]
        condition1 = filtered_df[i]["荷载"] == earth_loading1
        condition2 = filtered_df[i]["荷载"] == earth_loading2
        condition3 = filtered_df[i]["荷载"] == earth_loading3
        condition4 = filtered_df[i]["荷载"] == earth_loading4
        condition5 = filtered_df[i]["荷载"] == earth_loading5
        condition6 = filtered_df[i]["荷载"] == earth_loading6
        condition7 = filtered_df[i]["荷载"] == earth_loading7
        condition8 = filtered_df[i]["荷载"] == earth_loading8
        filtered_e1h.append(filtered_df[i][condition0 & (condition1 | condition2 | condition3 | condition4)])
        filtered_e1z.append(filtered_df[i][condition0 & (condition5 | condition6 | condition7 | condition8)])
        # filtered_e1h[i].reindex(columns = ['轴向 (kN)', '剪力-z (kN)', '剪力-y (kN)', '弯矩-y (kN*m)', '弯矩-z (kN*m)', '扭矩 (kN*m)'])
        filtered_e1h[i].insert(1, '竖向Pz', filtered_e1h[i].pop('轴向 (kN)'))
        filtered_e1h[i].insert(2, '顺力Hy', filtered_e1h[i].pop('剪力-z (kN)'))
        filtered_e1h[i].insert(3, '横力Hx', filtered_e1h[i].pop('剪力-y (kN)'))
        filtered_e1h[i].insert(4, '顺弯Mx', filtered_e1h[i].pop('弯矩-y (kN*m)'))
        filtered_e1h[i].insert(5, '横弯My', filtered_e1h[i].pop('弯矩-z (kN*m)'))
        filtered_e1h[i].insert(6, '扭矩Mz', filtered_e1h[i].pop('扭矩 (kN*m)'))
        filtered_e1h[i] = filtered_e1h[i].reset_index(drop=True)
        filtered_e1h[i].iloc[0, 1] = abs(min(filtered_e1h[i]["竖向Pz"], key=abs))
        filtered_e1h[i].iloc[0, 2] = abs(max(filtered_e1h[i]["顺力Hy"], key=abs))
        filtered_e1h[i].iloc[0, 3] = abs(max(filtered_e1h[i]["横力Hx"], key=abs))
        filtered_e1h[i].iloc[0, 4] = abs(max(filtered_e1h[i]["顺弯Mx"], key=abs))
        filtered_e1h[i].iloc[0, 5] = abs(max(filtered_e1h[i]["横弯My"], key=abs))
        filtered_e1h[i].iloc[0, 6] = abs(max(filtered_e1h[i]["扭矩Mz"], key=abs))
        filtered_e1h[i] = filtered_e1h[i].drop(filtered_e1h[i].index[[1, 2, 3]])
        # filtered_e1h[i] = filtered_e1h[i][:1] #只保留第一行,也可以
        filtered_e1h[i].iloc[0, 0] = pier_num[i]

        filtered_e1z[i].insert(1, '竖向Pz', filtered_e1z[i].pop('轴向 (kN)'))
        filtered_e1z[i].insert(2, '顺力Hy', filtered_e1z[i].pop('剪力-z (kN)'))
        filtered_e1z[i].insert(3, '横力Hx', filtered_e1z[i].pop('剪力-y (kN)'))
        filtered_e1z[i].insert(4, '顺弯Mx', filtered_e1z[i].pop('弯矩-y (kN*m)'))
        filtered_e1z[i].insert(5, '横弯My', filtered_e1z[i].pop('弯矩-z (kN*m)'))
        filtered_e1z[i].insert(6, '扭矩Mz', filtered_e1z[i].pop('扭矩 (kN*m)'))
        filtered_e1z[i] = filtered_e1z[i].reset_index(drop=True)
        filtered_e1z[i].iloc[0, 1] = abs(min(filtered_e1z[i]["竖向Pz"], key=abs))
        filtered_e1z[i].iloc[0, 2] = abs(max(filtered_e1z[i]["顺力Hy"], key=abs))
        filtered_e1z[i].iloc[0, 3] = abs(max(filtered_e1z[i]["横力Hx"], key=abs))
        filtered_e1z[i].iloc[0, 4] = abs(max(filtered_e1z[i]["顺弯Mx"], key=abs))
        filtered_e1z[i].iloc[0, 5] = abs(max(filtered_e1z[i]["横弯My"], key=abs))
        filtered_e1z[i].iloc[0, 6] = abs(max(filtered_e1z[i]["扭矩Mz"], key=abs))
        filtered_e1z[i] = filtered_e1z[i].drop(filtered_e1z[i].index[[1, 2, 3]])
        filtered_e1z[i].iloc[0, 0] = pier_num[i]

        filtered_e1.append(pd.concat([filtered_e1h[i], filtered_e1z[i]], axis=0))
        filtered_e1[i] = filtered_e1[i].round(0)

        dfs[pier_num[i]] = filtered_e1[i]
    # str_a = './resources/恒载1+{}地震响应.xlsx'.format(earth_style)
    # str_b = str_contents + '/恒载+{}地震响应.xlsx'.format(earth_style)
    str_b = str_contents + '/恒载+{}地震响应'.format(earth_style) + '+{}.xlsx'.format(datetime.now().strftime("%d_%H_%M"))
    # str_b = './resources/恒载+{}地震响应.xlsx'.format(earth_style)
    # writer = pd.ExcelWriter(str_a, engine='xlsxwriter')
    # for sheet_name in dfs.keys():
    #     dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    # writer.save()

    filtered_e11 = pd.concat(filtered_e1)
    filtered_e11 = filtered_e11.reset_index(drop=True)
    # b_list2 = filtered_e11['单元'].drop_duplicates().values.tolist()
    # for i in b_list2:
    #     condition0 = filtered_e11["单元"] == i
    #     condition1 = filtered_e11["荷载"] == earth_loading1
    #     condition2 = filtered_e11["荷载"] == earth_loading2
    #     condition3 = filtered_e11["荷载"] == earth_loading3
    #     condition4 = filtered_e11["荷载"] == earth_loading4
    filtered_e11.to_excel(str_b)


def earthquakeAnddeadload_response_dundi(pier_numlist, earth_style, filtered_df, str_contents, pier_num):
    # 该函数为输出恒载和地震响应内力组合的表格，墩底，格式为竖向Pz,弯矩M

    filtered_e1h = []
    filtered_e1z = []
    filtered_e1 = []

    dfs = {}
    if str(earth_style) == 'E1':
        earth_loading1 = '恒+E1H(最大)'
        earth_loading2 = '恒+E1H(最小)'
        earth_loading3 = '恒-E1H(最大)'
        earth_loading4 = '恒-E1H(最小)'
        earth_loading5 = '恒+E1Z(最大)'
        earth_loading6 = '恒+E1Z(最小)'
        earth_loading7 = '恒-E1Z(最大)'
        earth_loading8 = '恒-E1Z(最小)'
    elif str(earth_style) == 'E2':
        earth_loading1 = '恒+E2H(最大)'
        earth_loading2 = '恒+E2H(最小)'
        earth_loading3 = '恒-E2H(最大)'
        earth_loading4 = '恒-E2H(最小)'
        earth_loading5 = '恒+E2Z(最大)'
        earth_loading6 = '恒+E2Z(最小)'
        earth_loading7 = '恒-E2Z(最大)'
        earth_loading8 = '恒-E2Z(最小)'
    for i in range(len(pier_numlist)):
        condition0 = filtered_df[i]["单元"] == pier_numlist[i][1]
        condition1 = filtered_df[i]["荷载"] == earth_loading1
        condition2 = filtered_df[i]["荷载"] == earth_loading2
        condition3 = filtered_df[i]["荷载"] == earth_loading3
        condition4 = filtered_df[i]["荷载"] == earth_loading4
        condition5 = filtered_df[i]["荷载"] == earth_loading5
        condition6 = filtered_df[i]["荷载"] == earth_loading6
        condition7 = filtered_df[i]["荷载"] == earth_loading7
        condition8 = filtered_df[i]["荷载"] == earth_loading8
        filtered_e1h.append(filtered_df[i][condition0 & (condition1 | condition2 | condition3 | condition4)])
        filtered_e1z.append(filtered_df[i][condition0 & (condition5 | condition6 | condition7 | condition8)])
        # filtered_e1h[i].reindex(columns = ['轴向 (kN)', '剪力-z (kN)', '剪力-y (kN)', '弯矩-y (kN*m)', '弯矩-z (kN*m)', '扭矩 (kN*m)'])
        filtered_e1h[i].insert(1, '竖向Pz', filtered_e1h[i].pop('轴向 (kN)'))
        filtered_e1h[i].insert(2, '弯矩M', filtered_e1h[i].pop('弯矩-z (kN*m)'))
        filtered_e1h[i] = filtered_e1h[i].reset_index(drop=True)
        filtered_e1h[i].iloc[0, 1] = abs(min(filtered_e1h[i]["竖向Pz"], key=abs))
        filtered_e1h[i].iloc[0, 2] = abs(max(filtered_e1h[i]["弯矩M"], key=abs))
        filtered_e1h[i] = filtered_e1h[i].drop(filtered_e1h[i].index[[1, 2, 3]])
        # filtered_e1h[i] = filtered_e1h[i].drop(filtered_e1h[i].columns['剪力-z (kN)', '剪力-y (kN)', '弯矩-y (kN*m)', '扭矩 (kN*m)'], axis=1)
        filtered_e1h[i] = filtered_e1h[i].drop(filtered_e1h[i].columns[[4, 5, 6, 7, 8]], axis=1)

        filtered_e1h[i].iloc[0, 0] = pier_num[i]

        filtered_e1z[i].insert(1, '竖向Pz', filtered_e1z[i].pop('轴向 (kN)'))
        filtered_e1z[i].insert(2, '弯矩M', filtered_e1z[i].pop('弯矩-y (kN*m)'))
        filtered_e1z[i] = filtered_e1z[i].reset_index(drop=True)
        filtered_e1z[i].iloc[0, 1] = abs(min(filtered_e1z[i]["竖向Pz"], key=abs))
        filtered_e1z[i].iloc[0, 2] = abs(max(filtered_e1z[i]["弯矩M"], key=abs))
        filtered_e1z[i] = filtered_e1z[i].drop(filtered_e1z[i].index[[1, 2, 3]])
        filtered_e1z[i] = filtered_e1z[i].drop(filtered_e1z[i].columns[[4, 5, 6, 7, 8]], axis=1)
        filtered_e1z[i].iloc[0, 0] = pier_num[i]

        filtered_e1.append(pd.concat([filtered_e1h[i], filtered_e1z[i]], axis=0))
        filtered_e1[i] = filtered_e1[i].round(0)

        dfs[pier_num[i]] = filtered_e1[i]
    # str_a = './resources/恒载1+{}墩底地震响应.xlsx'.format(earth_style)
    # str_b = './resources/恒载+{}墩底地震响应.xlsx'.format(earth_style)
    # str_b = str_contents + '/恒载+{}墩底地震响应.xlsx'.format(earth_style)
    str_b = str_contents + '/恒载+{}墩底地震响应'.format(earth_style) + '+{}.xlsx'.format(datetime.now().strftime("%d_%H_%M"))

    # writer = pd.ExcelWriter(str_a, engine='xlsxwriter')
    # for sheet_name in dfs.keys():
    #     dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    # writer.save()

    filtered_e11 = pd.concat(filtered_e1)
    filtered_e11 = filtered_e11.reset_index(drop=True)
    # b_list2 = filtered_e11['单元'].drop_duplicates().values.tolist()
    # for i in b_list2:
    #     condition0 = filtered_e11["单元"] == i
    #     condition1 = filtered_e11["荷载"] == earth_loading1
    #     condition2 = filtered_e11["荷载"] == earth_loading2
    #     condition3 = filtered_e11["荷载"] == earth_loading3
    #     condition4 = filtered_e11["荷载"] == earth_loading4
    filtered_e11.to_excel(str_b)


import pandas as pd
from datetime import datetime
import os
import shutil


def main():
    file_name = "白沙洲引桥汉阳侧引桥50+75+50m-铁路墩-0804.xlsx"
    sExcelFile = "./resources/" + file_name
    # sExcelFile="./resources/1130-7#墩-2.0m--实心墩.xlsx"
    df = pd.read_excel(sExcelFile, sheet_name='Sheet1')
    str_time = datetime.now().strftime("%Y_%m%d_%H_%M_%S")
    str_contents = "./resources/{}+{}".format(file_name, str_time)
    os.makedirs(str_contents)
    sExcelFile_2 = str_contents + "/" + file_name
    shutil.copy(sExcelFile, sExcelFile_2)
    ###提取不重复的数据
    # a = df.drop_duplicates(subset=['单元'],keep='first')
    # print(a)
    # 把不重复d元素转换成list:
    b = df['单元'].drop_duplicates().values.tolist()
    print(b)
    print(type(b))
    print(len(b))
    # b1为墩顶，b2为墩底，b3为承台底
    b1 = [];
    b2 = [];
    b3 = []
    for i in range(len(b)):
        if i % 3 == 0:
            b1.append(b[i])
        elif i % 3 == 1:
            b2.append(b[i])
        else:
            b3.append(b[i])
    # b_list为实际墩号的种类，数量
    b_list = []
    for i in range(len(b1)):
        b_list.append([b1[i], b2[i], b3[i]])
    print(b_list)
    print(len(b_list))
    ######################################################################################
    ######################################################################################
    """
    下面这一段代码的思路是:字典的key作为sheet名称也就是桩号，value作为表格的内容。
    写入到Excel中也是通过字典来实现的
    """
    pier_num = ['03#墩', '04#墩', '05#墩', '00#墩', '01#墩', '10#墩', '11#墩', '12#墩', '13#墩', '14#墩']
    dfs = {}
    filtered_df = []
    for i in range(len(b_list)):
        # print(i)
        condition1 = df["单元"] == b_list[i][0]
        condition2 = df["单元"] == b_list[i][1]
        condition3 = df["单元"] == b_list[i][2]
        filtered_df.append(df[(condition1) | (condition2) | (condition3)])
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

    # 输出承台底的内力响应表格
    earthquakeAnddeadload_response(b_list, E1, filtered_df, str_contents, pier_num)
    earthquakeAnddeadload_response(b_list, E2, filtered_df, str_contents, pier_num)
    # 输出墩底的内力响应表格
    earthquakeAnddeadload_response_dundi(b_list, E1, filtered_df, str_contents, pier_num)
    earthquakeAnddeadload_response_dundi(b_list, E2, filtered_df, str_contents, pier_num)


if __name__ == '__main__':
    main()
