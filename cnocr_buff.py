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


def time_stamp():
    return time.strftime("%Y-%m-%d %H:%M:%S ", time.localtime())


def get_grab_ocr_result(client, grab_range: list):
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


def close_window_test():
    print(time_stamp(), '关闭游戏')


def click_left(pos: tuple):
    win32api.SetCursorPos(pos)
    time.sleep(0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)


def click_left_test(pos: tuple):
    print(time_stamp(), '选中角色1')


def double_click_left(pos: tuple):
    win32api.SetCursorPos(pos)
    time.sleep(0.5)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)


def double_click_left_test(pos: tuple):
    print(time_stamp(), '进入角色')


def shutoff_pc():
    os.system("shutdown -s")


def back_to_role_list():
    # ESC
    win32api.keybd_event(27, 0, 0, 0)
    win32api.keybd_event(27, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
    # 斜杠
    click_left(config.logout_pos)


def back_to_role_list_test():
    print(time_stamp(), '返回角色界面')


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
        if '莫托尔' in pic_text or '赞达拉之魂' in pic_text:
            return True
        else:
            return False
    elif choice == '2':
        if '伦萨克' in pic_text or '萨鲁法尔' in pic_text or '屠龙者的咆哮' in pic_text:
            return True
        else:
            return False
    elif choice == '3':
        if '黑手' in pic_text or '酋长的祝福' in pic_text:
            return True
        else:
            return False
    else:
        return False


def get_buff(win_handle, client, choice):
    offline_check = 0
    buff_dict = {'1': '哈卡', '2': '龙头', '3': '酋长'}
    start_time = time.time()
    while True:
        if time.time() - start_time >= random.randint(60, 300):
            print(time_stamp(), '开始随机动作')
            offline_check = 1
            start_time = time.time()
            random_move()
        if offline_check:
            print(time_stamp(), '开始掉线监测')
            string_list = get_grab_ocr_result(client, config.offline_range)
            if string_list:
                for item in string_list:
                    if '断开' in item:
                        close_window()
                        print(time_stamp(), '检测到已掉线，中断程序！')
                        return
            offline_check = 0
        chat_list = get_grab_ocr_result(client, config.chat_range)
        if chat_list:
            for item in chat_list:
                print(time_stamp(), '检测到喊话：{}'.format(item))
                if does_buff_start(choice, item):
                    back_to_role_list()
                    # back_to_role_list_test()
                    time.sleep(1.5)
                    click_left(config.role1_pos)
                    # click_left_test(config.role1_pos)
                    time.sleep(0.5)
                    double_click_left(config.role1_pos)
                    # double_click_left_test(config.role1_pos)
                    time.sleep(delay_dict[choice])
                    close_window()
                    # close_window_test()
                    print(time_stamp(), '已成功挂到 {} buff'.format(buff_dict[choice]))
                    return


if __name__ == '__main__':
    print('本工具用于怀旧服蹲世界buff，请仔细阅读下面提示：')
    print('1.目前仅用于1920*1080分辨率，不操作窗口，应该不会有封号风险；')
    print('2.本工具用于单账号蹲buff，且吃buff大号位于角色列表首位，蹲buff小号位于吃buff大号下面；')
    print('3.蹲buff时除DBM禁用其他所有插件，加快载入速度，DBM用于跨位面监控，但是跨位面检测率低，建议单位面监控；')
    print('4.聊天框背景设为完全不透明，字体设置为最大18，不要被动作条遮挡，使用画图软件确定聊天框的 左 上 右 下 像素坐标，确保截图区域底色为聊天框单色;')
    print('5.聊天框频道除NPC大喊其他全部关闭，即“聊天”“通用频道”内全部叉掉，“其他”只保留怪物信息的勾选，其他全叉掉;')
    print('6.快捷键2设置下面宏，且包里有一个对应宏的精华，用于挂机;')
    print('    /dismount')
    print('    /use 次级秘法精华')
    print('    /use 强效秘法精华')
    print(config.chat_range)
    config.chat_range[0] = int(input('请输入聊天框宽的最左像素值：').strip())
    config.chat_range[1] = int(input('请输入聊天框高的最上像素值：').strip())
    config.chat_range[2] = int(input('请输入聊天框宽的最右像素值：').strip())
    config.chat_range[3] = int(input('请输入聊天框高的最下像素值：').strip())
    print(config.chat_range)
    input_string = input('选择要挂的buff序号：\r\n1 哈卡\r\n2 龙头\r\n3 酋长\r\n请输入序号并回车：')
    print('务必5秒内全屏魔兽窗口，并保持魔兽窗口再最前！！')
    time.sleep(5)
    # win_handle = win32gui.FindWindow(None, '魔兽世界')
    # win_handle = win32gui.FindWindow('Qt5QWindowIcon', 'OA邮箱')
    ocr = CnOcr()
    get_buff(None, ocr, input_string)
