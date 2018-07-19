import xlrd
import os
import shutil
import re

if __name__ == '__main__':
    print('@COPYRIGHT 供 "关爱抗战老兵公益基金" 使用')
    print('请输入EXCEL文件路径')
    xls_dir = input()
    while not os.path.exists(xls_dir):
        print('请输入正确的EXCEL文件路径')
        print(r'如：C:\Users\liuyi\Desktop\201807指定认养司号员公示名单.xlsx')
        xls_dir = input()
    print('请输入需要整理的源文件夹路径')
    source_dir = input()
    while not os.path.exists(source_dir):
        print('请输入正确的源文件夹路径')
        print(r'如：C:\Users\liuyi\Desktop\source')
        source_dir = input()
    print('请输入保存的目的文件夹路径')
    target_dir = input()
    while not os.path.exists(target_dir):
        print('请输入正确的目的文件夹路径')
        print(r'如：C:\Users\liuyi\Desktop\target')
        target_dir = input()
    print('程序运行中，请稍后')
    # 打开指定excel
    data = xlrd.open_workbook(xls_dir)
    # 获取指定sheet
    table = data.sheets()[0]
    # 获取最大行数
    nrows = table.nrows
    # print(nrows)
    # 捐赠人列表和老兵列表
    donors = []
    laobings = []
    i = 0
    j = 0
    # 循环,生成donors和laobings
    while i < nrows:
        if table.row_values(i)[0] == "中华社会救助基金会·关爱抗战老兵公益基金·抗战老兵助养行动":
            laobing_list = []
            j += 1
            # print(table.row_values(i+1)[0]+'如下：')
            # 司号员加入 donors
            donors.append(table.row_values(i+1)[0])
            i = i+3
            while table.row_values(i)[1]:
                # 老兵名单加入 laobings
                laobing_list.append(table.row_values(i)[1])
                # print(table.row_values(i)[1])
                i += 1
            laobings.append(laobing_list)
        else:
            i = i+1

    i = 0
    j = 0
    k = 0
    error_msgs = []
    for donor in donors:
        # 每个捐赠人创建一个文件夹
        # 规范文件夹名称
        donor_format = re.sub(r'司号员|·指定助养抗战老兵名单|指定助养抗战老兵名单|抗战老兵助养名单|·集结添饭抗战老兵名单|指定认养老兵抗战老兵名单', '', donor)
        donor_format = donor_format.strip()
        path = target_dir + '/'+donor_format
        os.makedirs(path)
        # 将老兵放到相应文件夹
        for laobing in laobings[k]:
            laobing = laobing.strip()
            break_flag = False
            laobing_found = False
            for root, dirs, files in os.walk(source_dir):
                for i in range(len(dirs)):
                    for root1, dirs1, files1 in os.walk(source_dir + '/'+dirs[i]):
                        for j in range(len(files1)):
                            if laobing in files1[j]:
                                file_path = root1+'/'+files1[j]
                                new_file_path = path+'/'+files1[j]
                                shutil.copy(file_path, new_file_path)
                                laobing_found = True
                                break_flag = True
                                break
                        if break_flag:
                            break
                    if break_flag:
                        break
                if break_flag:
                    break
            if not laobing_found:
                error_msgs.append(donor_format+'认养的'+laobing + "找不到!")

        k += 1
    for error_msgs in error_msgs:
        print(error_msgs)
    print('程序执行完毕，按任意键退出')
    exits = input()
