import pandas as pd

"""
该函数作用，N45～N60，改为N50～N65
只改数字，后面文字不变
"""


def main():
    df = pd.read_excel('中间文件.xlsx')

    df_n_1_a = df["板件编号"].str.extract(r'([N])(\d+)([abcdefg]+)', expand=False)
    df_n_1 = df["板件编号"].str.extract(r'([N])(\d+)', expand=True)

    # df_name = df["板件编号"].str.split('-', expand=True) # 分列的两种方法
    df_name = df["板件编号"].str.partition('-')
    print(df_name)

    # df_n_1.isna()
    # print(df_n_1[1].notna())  # 判断NA值，空值
    df_n_1["板件编号修改后"] = "0"  # 增加一列

    for i in df.index:
        if df_n_1[1].notna()[i]:
            if (int(df_n_1[1][i]) > 44) & (int(df_n_1[1][i]) < 76):
                df_n_1[1][i] = str(int(df_n_1[1][i]) + 5)
            else:
                df_n_1[1][i] = str(df_n_1[1][i])
            df_n_1["板件编号修改后"][i] = str(df_n_1[0][i]) + str(df_n_1[1][i]) + str(df_n_1_a[2][i]) + df_name[2][i]
        else:
            df_n_1["板件编号修改后"][i] = df["板件编号"][i]
    df_n_1["板件编号修改后"] = df_n_1["板件编号修改后"].str.replace('nan', '')
    df_n_1.pop(0)
    df_n_1.pop(1)
    str_a = "修改编号后文件.xlsx"
    df_n_1.to_excel(str_a)


if __name__ == '__main__':
    main()
