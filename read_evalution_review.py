import os  
import xlrd
import xlwt


def file_name(file_dir):   
    files_name = []
    professions = []   
    for root, dirs, files in os.walk(file_dir): #读文件夹中的文件名（根目录，子文件夹目录，文件名）  
        for file in files:  
            files_name.append(os.path.join(root, file)) #合成标准路径名
            professions.append(file[:-5])
    num = len(files_name)              
    return professions, files_name, num #都有哪些省，文件路径，文件总数

def read_profession(profession, file, data, province_list):
    book=xlrd.open_workbook(file)
    table = book.sheet_by_index(0)
    nrows = table.nrows
    ncols = table.ncols
    for row in range(1,nrows):
        province_name = table.row_values(row)[0]
        if province_name == '':
            for i in range(1, nrows-1):
                if (table.row_values(row-i)[0] != ''):
                    province_name = table.row_values(row-i)[0]
                    break
        if province_name not in province_list:
            data[province_name] = {}
            province_list.append(province_name)
        profession_data = []
        for column in range(1,8):
            profession_data.append(table.row_values(row)[column])
        if (profession_data[0] == ''):
            profession_data[0] = 1
        else:
            profession_data[0] = int(profession_data[0])
        profession_data[3] = float(profession_data[3])
        profession_data[3] = float('%.2f'%profession_data[3])
        try:
            data[province_name][profession].append(profession_data)
        except:
            data[province_name][profession] = []
            data[province_name][profession].append(profession_data)



def read(file_dir):
    data = {}
    province_list = []
    professions, files_name, num = file_name(file_dir)
    for index, profession in enumerate(professions):
        read_profession(profession, files_name[index], data, province_list)
        '''
        try:
            read_profession(profession, files_name[index], data, province_list)
        except:
            print(profession)
        '''
    return data


if __name__ == "__main__":
    file_dir = r"C:\Users\x270\Desktop\investment review"
    read(file_dir)