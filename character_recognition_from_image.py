from aip import AipOcr
import re
import os
from openpyxl import load_workbook

# import xlrd
# import xlwt
# from xlutils.copy import copy
# import openpyxl


FPS = r'FPS'
LOST = r'输入/输出丢失'

# 账号网站申请https://cloud.baidu.com/?from=console
APP_ID = '******'
API_KEY = '***********'
SECRET_KEY = '****************'

DOAT_PNG_FILE_DIR = "F:/01_python/Python_20201119_01/dota2_png_file_dir/"
DATA_EXCEL_DIR = "F:/01_python/Python_20201119_01/test_data_output/"

DATA_EXCEL_NAME = "test_data_output.xlsx"
DOTA_SHEET_NAME = "Dota2"
CSGO_SHEET_NAME = "CSGO"
# TITLE_VALUE = ["ID", "DATE", "TIME", "FPS", "PING", "LOST_IN", "LOST_OUT"]


# 识别图片中数字，返回FPS, PING, LOST_IN, LOST_OUT
def get_char_from_dota2_png(png_file_dir):
    fps = ''
    ping = ''
    lost_in = ''
    lost_out = ''
    image = open(png_file_dir, 'rb')
    image = image.read()
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    # res = client.basicGeneral(image) # 通用文字识别，每天 50 000 次免费
    res = client.basicAccurate(image)   # 通用文字高精度识别，每天 800 次免费
    if 'words_result' in res.keys():
        for item in res['words_result']:
            # print(item['words'])
            result = re.search(FPS, item['words'])
            if result:
                # nPos = item['words'].index(FPS)
                # print(item['words'])
                pattern = re.compile(r'\d+')
                res = re.findall(pattern, item['words'])

                res_tmp = res[0]
                if len(res_tmp) > 3:
                    fps = res_tmp[0:3]
                    ping = res_tmp[3:len(res_tmp)]
                else:
                    fps = res[0]
                    ping = res[1]
                # print(fps, ping)
            else:
                result = re.search(LOST, item['words'])
                if result:
                    # print(item['words'])
                    pattern = re.compile(r'\d+')
                    res = re.findall(pattern, item['words'])
                    # print(res)
                    if res:
                        lost = res[0]
                        # print(lost)
                        lost_len = len(lost)
                        lost_pos = int(lost_len/2)
                        # print(lost_pos)
                        lost_in = lost[0:lost_pos]
                        lost_out = lost[lost_pos:lost_len]
                        print(lost_in, lost_out)
    else:
        print(res)
    return fps, ping, lost_in, lost_out


# 修改文件名称，' ' -> '_'
def mod_png_file_name(png_file_dir):
    file_list = os.listdir(png_file_dir)
    for name in file_list:
        index = file_list.index(name)
        # print(file_name, index)

        # 文件名称处理
        file_list[index] = re.sub(r"\s", "_", name)
        # print(name, index)
        os.chdir(png_file_dir)
        os.rename(name, file_list[index])
    # print(file_list[0])
    return file_list


# 文件路径拼接
def merge_dir(file_dir, png_name):
    dir_str = file_dir + png_name
    return dir_str


# 获取年月日，时分秒 screencap 2020-11-20 19-09-39.jpg
def get_date_time(file_name):
    date = file_name[10:20]
    time = file_name[21:29]
    # print(date, time)
    return date, time


#
def xlsx_value_package(index, date_parm, net_param):
    xlsx_value = list()

    xlsx_value.append(index + 1)
    for num in range(0, 2):
        xlsx_value.append(date_parm[num])
    for num in range(0, 4):
        xlsx_value.append(net_param[num])

    return xlsx_value


def write_excel_xlsx(excel_name, sheet_name, xlsx_value):
    wb = load_workbook(excel_name)
    ws = wb[sheet_name]
    ws.append(xlsx_value)
    wb.save(excel_name)


# 函数从这里开始
g_doat_png_file_list = mod_png_file_name(DOAT_PNG_FILE_DIR)
for g_file_name in g_doat_png_file_list:
    print(g_file_name)

    g_png_file_dir = merge_dir(DOAT_PNG_FILE_DIR, g_file_name)
    print(g_png_file_dir)

    g_net_param = get_char_from_dota2_png(g_png_file_dir)
    print(g_net_param[0], g_net_param[1], g_net_param[2], g_net_param[3])

    g_date_parm = get_date_time(g_file_name)
    print(g_date_parm[0], g_date_parm[1])

    g_index = g_doat_png_file_list.index(g_file_name)

    g_xlsx_value = xlsx_value_package(g_index, g_date_parm, g_net_param)

    # 写入execl
    os.chdir(DATA_EXCEL_DIR)
    write_excel_xlsx(DATA_EXCEL_NAME, DOTA_SHEET_NAME, g_xlsx_value)


