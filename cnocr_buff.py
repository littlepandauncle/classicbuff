"""
注意点：
1. 哈卡、龙头、酋长buff分别吃
2. 酋长buff大号在十字路口吃，小号在奥格蹲
3. 像素坐标实在1920*1080分辨率下，以我自己设置确定，
   不同分辨率或者不同设置，自行获取聊天框、角色1、角色2、登陆按钮、掉线提示等几个像素范围和像素点在config文件更改
4. 聊天框设置为不透明，防止误识别
5. 聊天框去掉其他频道信息，防止误识别
6. 差不多1秒监测一次
"""

import os
from cnocr import CnOcr
from PIL import ImageGrab
import time
import win32api
import win32con
import random

import config

task_run_flag = 0

# 1代表哈卡buff；2代表黑龙龙头；3代表奈法龙头；4代表酋长buff
string_dict = {
    '1': config.haka_string,
    '2': config.heilong_string,
    '3': config.naifa_string,
    '4': config.qiuzhang_string
}

delay_dict = {
    '1': config.haka_delay,
    '2': config.longtou_delay,
    '3': config.qiuzhang_delay
}


def get_grab_ocr_result(client, grab_range: tuple):
    # OCR调用频率限制在2次/秒，根据CPU主频不同，延时根据实际需要调整
    res_list = []
    img = ImageGrab.grab(bbox=grab_range)
    img.save('grab.jpg')
    res = client.ocr('grab.jpg')
    if res:
        for str_line in res:
            res_list.append("".join(str_line))
    return res_list


def close_window():
    # win32gui.PostMessage(win_handle, win32con.WM_CLOSE, 0, 0)
    # alt
    win32api.keybd_event(18, 0, 0, 0)
    time.sleep(0.1)
    # F4
    win32api.keybd_event(115, 0, 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(115, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)


def click_left(pos: tuple):
    win32api.SetCursorPos(pos)
    time.sleep(0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)


def double_click_left(pos: tuple):
    win32api.SetCursorPos(pos)
    time.sleep(0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)


def shutoff_pc():
    os.system("shutdown -s")


def back_to_role_list():
    # ESC
    win32api.keybd_event(27, 0, 0, 0)
    win32api.keybd_event(27, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
    # 斜杠
    click_left(config.logout_pos)


def press_1():
    win32api.keybd_event(49, 0, 0, 0)
    win32api.keybd_event(49, 0, win32con.KEYEVENTF_KEYUP, 0)


def press_2():
    win32api.keybd_event(50, 0, 0, 0)
    win32api.keybd_event(50, 0, win32con.KEYEVENTF_KEYUP, 0)


def press_3():
    win32api.keybd_event(51, 0, 0, 0)
    win32api.keybd_event(51, 0, win32con.KEYEVENTF_KEYUP, 0)


def press_space():
    win32api.keybd_event(32, 0, 0, 0)
    win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)


# 1键精华宏；2键上盾；3键上恢复；空格键跳
def random_move():
    random_dict = {
        0: press_2,
        1: press_space
    }
    index = random.randint(0, 1)
    random_dict[index]()


def does_buff_start(choice, pic_text):
    if choice == '1':
        if '莫托尔' in pic_text:
            return True
        else:
            return False
    elif choice == '2':
        if '伦萨克' in pic_text or '萨鲁法尔' in pic_text:
            return True
        else:
            return False
    elif choice == '3':
        if '黑手' in pic_text:
            return True
        else:
            return False
    else:
        return False


def get_buff(win_handle, client, choice):
    offline_check = 0
    start_time = time.time()
    while True:
        if time.time() - start_time >= random.randint(600, 1200):
            print('开始随机动作')
            offline_check = 1
            start_time = time.time()
            random_move()
        if offline_check:
            print('开始掉线监测')
            string_list = get_grab_ocr_result(client, config.offline_range)
            if string_list:
                for item in string_list:
                    if '断开' in item:
                        close_window()
                        return
            offline_check = 0
        chat_list = get_grab_ocr_result(client, config.chat_range)
        if chat_list:
            for item in chat_list:
                if does_buff_start(choice, item):
                    back_to_role_list()
                    time.sleep(1.5)
                    click_left(config.role1_pos)
                    time.sleep(0.5)
                    double_click_left(config.role1_pos)
                    time.sleep(delay_dict[choice])
                    close_window()
                    return


if __name__ == '__main__':
    input_string = input('选择要挂的buff序号：\r\n1 哈卡\r\n2 龙头\r\n3 酋长\r\n请输入序号并回车：')
    print('务必5秒内全屏魔兽窗口！！')
    time.sleep(5)
    # win_handle = win32gui.FindWindow(None, '魔兽世界')
    # win_handle = win32gui.FindWindow('Qt5QWindowIcon', 'OA邮箱')
    ocr = CnOcr()
    get_buff(None, ocr, input_string)
