"""
该py文件是操作txt来实现操作xtract软件.
主要实现功能：
批量在指定截面输入荷载
"""


def earthquake_response(pier_numlist, earth_style, filtered_df):
    # 该函数为输出地震响应内力的表格，格式为弯矩，剪力，轴力
    filtered_e1h = []
    filtered_e1z = []
    filtered_e1 = []
    pier_num = ['5#墩', '6#墩', '7#墩', '8#墩', '9#墩', '10#墩', '11#墩', '12#墩', '13#墩', '14#墩']
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


def main():
    file_path = "./resources/修改前.txt"
    file_path2 = "./resources/修改后.txt"

    with open(file_path, 'r', encoding='utf-8') as f1, open(file_path2, 'a+', encoding='utf-8') as f2:
        # lines = f1.readlines()
        print(f1)
        for i in f1:
            f2.write(i)


    pass

    # with open(file_path2, 'w', encoding='utf-8') as f3:
    #     f3.write('chenlaie')


    # print(len(lines))

    # for i in f1:
    #     f2.write(i)

    # print(lines)


if __name__ == '__main__':
    main()
