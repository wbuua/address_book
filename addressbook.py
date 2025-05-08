import pandas as pd


def com(add):
    print('欢迎使用通讯录\n[---1.查询---]\n[---2.增加---]\n[---3.删除---]\n[---4.修改---]\n[---0.退出---]')
    flag = True
    while flag:
        re = int(input('\n>>>请输入数字：'))
        if re == 1:
            name = str(input('>>>输入人名或输入“all”查询全部通讯录：'))
            if name in add.keys():
                print(f'{name}的号码：{add[name]}')
            elif name == 'all':
                print(add)
            else:
                print('[无此人]')

        elif re == 2:
            name = str(input('>>>输入要增加的名字：'))
            num = int(input('>>>输入号码：'))
            add[name] = num
            print('[已录入]')


        elif re == 3:
            name = str(input('>>>输入要删除的名字：'))
            del add[name]
            print(f'已删除{name}')

        elif re == 4:
            name = str(input('>>>输入要修改的名字：'))
            num = int(input('>>>输入新号码：'))
            add[name] = num
            print(f'已修改{name}')

        elif re == 0:
            print('退出通讯录')
            flag = False

        else:
            print('wrong')
    return add


if __name__ == '__main__':
    add = pd.read_excel('./add.xlsx')
    add = add.set_index('name')['num'].to_dict()
    add = com(add)
    df = pd.DataFrame(list(add.items()), columns=['name', 'num'])
    print(df)
    try:
        df.to_excel('add.xlsx', sheet_name='sheet1', index=False)
        print("通讯录已成功保存到 add.xlsx 文件中")
    except Exception as e:
        print(f"保存通讯录时出错：{e}")
