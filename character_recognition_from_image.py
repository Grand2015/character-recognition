from aip import AipOcr
import re
import os
from openpyxl import load_workbook
from PIL import Image
import time

# import xlrd
# import xlwt
# from xlutils.copy import copy
# import openpyxl


FPS = r'FPS'
PING = r'PING'
LOST = r'输入'


FPS_CSGO = r'fps'
PING_CSGO = r'ping'
LOSS_CSGO = r'loss'


APP_ID = '23013103'
API_KEY = 'RPifPZIf2C3wkP5fQx29fdIZ'
SECRET_KEY = '8t9d5IlNff0lBkHoWWzdUdHkgoCjDwyx'
'''
APP_ID = '23022152'
API_KEY = 'qPbrsSt446AFl5jH13HpAyQk'
SECRET_KEY = 'y0LEANoKeGdr9n4dCGWR4xseNfIP4Yr2'
'''
DOAT_PNG_FILE_DIR = "F:/01_python/Python_20201119_01/dota2_png_file_dir/"
CSGO_PNG_FILE_DIR = "F:/01_python/Python_20201119_01/csgo_png_file_dir/"
# PNG_FILE_DIR = [DOAT_PNG_FILE_DIR, CSGO_PNG_FILE_DIR]

DATA_EXCEL_DIR = "F:/01_python/Python_20201119_01/test_data_output/"
DATA_EXCEL_NAME = "test_data_output.xlsx"

DOTA_SHEET_NAME = "Dota2"
CSGO_SHEET_NAME = "CSGO"
SHEET_NAME = [DOTA_SHEET_NAME, CSGO_SHEET_NAME]
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
    print(client)
    res = client.basicGeneral(image)  # 通用文字识别，每天 50 000 次免费
    '''
    try:
        res = client.basicGeneral(image)  # 通用文字识别，每天 50 000 次免费
    except Exception as e:
        print("api err！")
        time.sleep(1)
        res = client.basicGeneral(image)
        print(res)
    '''
    # res = client.basicAccurate(image)   # 通用文字高精度识别，每天 800 次免费
    if 'words_result' in res.keys():
        for item in res['words_result']:
            print(item['words'])
            # result = re.search(FPS, item['words'])
            if re.search(FPS, item['words']) or re.search(PING, item['words']):
                # nPos = item['words'].index(FPS)
                # print(item['words'])
                pattern = re.compile(r'\d+')
                res = re.findall(pattern, item['words'])

                if res:
                    if len(res) == 2:
                        fps = res[0]
                        ping = res[1]
                    else:
                        res_tmp = res[0]
                        print(res_tmp)
                        if len(res_tmp) > 4:
                            fps = res_tmp[0:3]
                            ping = res_tmp[3:len(res_tmp)]
                        elif len(res_tmp) > 3:
                            fps = res_tmp[0:2]
                            ping = res_tmp[2:len(res_tmp)]
                        else:
                            print("len of fps & ping is 3！")
                            fps = res_tmp[0:2]
                            ping = res_tmp[2:len(res_tmp)]

                    print(fps, ping)
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


def get_char_from_csgo_png(png_file_dir):
    fps = ''
    ping = ''
    lost_in = ''
    lost_out = ''

    image = open(png_file_dir, 'rb')
    image_bw = image.read()
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    # print(client)

    # res = client.basicGeneral(image_bw)  # 通用文字识别，每天 50 000 次免费
    res = client.basicAccurate(image_bw)   # 通用文字高精度识别，每天 800 次免费
    if 'words_result' in res.keys():
        for item in res['words_result']:
            print(item['words'])
            result = re.search(FPS_CSGO, item['words'])
            if result:
                # nPos = item['words'].index(FPS)
                # print(item['words'])
                pattern = re.compile(r'\d+')
                res = re.findall(pattern, item['words'])
                if res:
                    fps = res[0]
                    result = re.search(PING_CSGO, item['words'])
                    if result:
                        ping = res[len(res)-1]
                print(fps, ping)

            elif re.search(PING_CSGO, item['words']):  # 处理异常情况，识别结果中fps一行，对于那个单数一行
                pattern = re.compile(r'\d+')
                res = re.findall(pattern, item['words'])
                res2 = re.search('.', item['words'])  # 区分是否存在小数点
                if res and 4 == len(res) and res2:
                    fps = res[0]
                    ping = res[len(res)-1]
                if res and 3 == len(res) and res2 is None:
                    fps = res[0]
                    ping = res[len(res)-1]
                elif res:
                    fps = 'null'
                    ping = res[len(res)-1]

            else:
                result = re.search(LOSS_CSGO, item['words'])
                if result:
                    # print(item['words'])
                    pattern = re.compile(r'\d+')
                    res = re.findall(pattern, item['words'])
                    # print(res)
                    if res:
                        lost_in = res[0]
                        print(lost_in)
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
        if net_param[num] and net_param[num] != 'null':
            xlsx_value.append(int(net_param[num]))
        else:
            xlsx_value.append(net_param[num])

    return xlsx_value


def write_excel_xlsx(excel_name, sheet_name, xlsx_value):
    wb = load_workbook(excel_name)
    ws = wb[sheet_name]
    ws.append(xlsx_value)
    wb.save(excel_name)


def cut_image(image_dir, file_dir):
    img = Image.open(image_dir)
    print(img.size)
    if file_dir == DOAT_PNG_FILE_DIR:
        if img.size[0] >= 2560 and img.size[0] >= 1440:
            cropped = img.crop((2300, 0, 2560, 50))  # (left, upper, right, lower) 左上，右下
            # Image._show(cropped)
            cropped.save(image_dir)
    else:
        if img.size[0] >= 2560 and img.size[0] >= 1440:
            cropped = img.crop((1470, 1280, 2150, 1340))  # (left, upper, right, lower) 左上，右下
            # Image._show(cropped)
            cropped.save(image_dir)


# 函数从这里开始
# PNG_FILE_DIR = CSGO_PNG_FILE_DIR
PNG_FILE_DIR = DOAT_PNG_FILE_DIR
g_doat_png_file_list = mod_png_file_name(PNG_FILE_DIR)
for g_file_name in g_doat_png_file_list:
    print(g_file_name)

    # 增加1s延时，防止百度接口不响应普通用户
    time.sleep(1)

    # 图片文件路径和名称拼接
    g_png_file_dir = merge_dir(PNG_FILE_DIR, g_file_name)
    print(g_png_file_dir)

    # CSGO 图片裁剪，提高识别精度
    cut_image(g_png_file_dir, PNG_FILE_DIR)

    # 获取网络参数，Dota和CSGO数据路径不同分开处理
    if PNG_FILE_DIR == DOAT_PNG_FILE_DIR:
        g_net_param = get_char_from_dota2_png(g_png_file_dir)
    else:
        g_net_param = get_char_from_csgo_png(g_png_file_dir)
    print(g_net_param[0], g_net_param[1], g_net_param[2], g_net_param[3])

    # 获取时间参数
    g_date_parm = get_date_time(g_file_name)
    print(g_date_parm[0], g_date_parm[1])

    g_index = g_doat_png_file_list.index(g_file_name)

    # 网络参数、时间参数打包，方便写入excel
    g_xlsx_value = xlsx_value_package(g_index, g_date_parm, g_net_param)
    # 参数写入excel
    os.chdir(DATA_EXCEL_DIR)
    if PNG_FILE_DIR == DOAT_PNG_FILE_DIR:
        write_excel_xlsx(DATA_EXCEL_NAME, DOTA_SHEET_NAME, g_xlsx_value)
    else:
        write_excel_xlsx(DATA_EXCEL_NAME, CSGO_SHEET_NAME, g_xlsx_value)



