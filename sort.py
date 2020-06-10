import os
from shutil import copy
provinces = ['北京','天津','河北','山西','内蒙古','辽宁','吉林','黑龙江','上海','江苏','浙江','安徽','福建','江西','山东','河南','湖北','湖南','广东','广西','海南','重庆','四川','贵州','云南','西藏','陕西','甘肃','青海','宁夏','新疆']
#provinces = ['北京']
professions = ['核心网','基础网','DC','云','节能','应急','总册附表']
trans = ['接入','IP','STN','传输','数据']
#professions = ['核心网','总纲']

file_path1 = r"C:\Users\x270\Desktop\新建文件夹 (2)"
file_path2 = r"C:\Users\x270\Desktop\新建文件夹 (3)"

count = 0
for province in provinces:
    file_path = os.path.join(file_path1, province)
    try:
        for file in os.listdir(file_path):
            flag = 0
            for profession in professions:
                if file.find(profession) != -1:
                    from_path = os.path.join(file_path, file) #原始文件路径
                    to_path = os.path.join(file_path2, profession) #目标文件路径
                    to_path = os.path.join(to_path, province) #目标文件路径
                    copy(from_path, to_path)
                    flag = 1
                    break
                else:
                    for tran in trans:
                        if file.find(tran) != -1:
                            from_path = os.path.join(file_path, file) #原始文件路径
                            to_path = os.path.join(file_path2, '基础网') #目标文件路径
                            to_path = os.path.join(to_path, province) #目标文件路径
                            copy(from_path, to_path)
                            flag = 1
                            break

            if flag == 0:
                from_path = os.path.join(file_path, file) #原始文件路径
                to_path = os.path.join(file_path2, "其他") #目标文件路径
                to_path = os.path.join(to_path, province) #目标文件路径
                copy(from_path, to_path)
    except:
        print(province)
