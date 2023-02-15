import json
import os
import queue
import time
import pythoncom
import PyHook3

config_json_path = 'when_i_work.json'
config = {}
with open(config_json_path, 'r') as fp:
    config_json = fp.read()
    config = json.loads(config_json)
# 多少个消息写一次文件(大约)
BATCH_COUNT = config.get('batch_count', 10)
# 日志目录
LOG_DIR = config.get('log_dir', 'E:/')
# 日志文件头
LOG_FILE_PREFIX = config.get('log_prefix', 'when_i_work')
# 每条动作最小间隔,小于这个秒数的动作不记+
EVENT_MIN_INTERVAL = min(config.get('event_minimum_interval_seconds', 10), 3600)
# 筛选窗口名,窗口名包含这个列表里的字符串才记
FILTER_BY_WINDOW_NAME = config.get('filter_by_window_name', [])
# 保留几天的日志
KEEP_LOG_DAY = max(config.get('keep_log_day', 2), 2)
KEEP_LOG_SECONDS = KEEP_LOG_DAY * 86400
print(config)

# 消息队列
event_queue = queue.Queue()
# 最后一条记载消息的时间
last_event_time = 0


def write_file(out_str):
    global last_event_time
    timestamp = time.time()
    if timestamp - last_event_time < EVENT_MIN_INTERVAL:
        return
    last_event_time = timestamp
    event_queue.put(f'{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(timestamp))},{out_str}')

    if event_queue.qsize() < BATCH_COUNT:
        return
    file_day = time.strftime("%Y-%m-%d", time.localtime())
    out_file = os.path.join(LOG_DIR, f'{LOG_FILE_PREFIX}_{file_day}.csv')
    # print(event_queue.qsize())
    print(f'{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} 写{out_file}')
    with open(out_file, 'a') as fp:
        while not event_queue.empty():
            e = event_queue.get()
            # print(e)
            fp.write(f'{e}\n')

    dir_or_files = os.listdir(LOG_DIR)
    for dir_file in dir_or_files:
        dir_file_path = os.path.join(LOG_DIR, dir_file)
        if os.path.isfile(dir_file_path):
            if dir_file.startswith(f'{LOG_FILE_PREFIX}_') and dir_file.endswith('.csv'):
                day_str = dir_file[len(LOG_FILE_PREFIX) + 1:-4]
                if len(day_str) != 10:
                    continue
                file_str_time = time.mktime(time.strptime(day_str, "%Y-%m-%d"))
                if timestamp - file_str_time > KEEP_LOG_SECONDS:
                    print(f'{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())} 删除过期日志{dir_file_path}')
                    os.remove(dir_file_path)


def onMouseEvent(event):
    # 监听鼠标事件     
    # print("MessageName:", event.MessageName)
    # print("Message:", event.Message)
    # print("Time:", event.Time)
    # print("Window:", event.Window)
    # print("WindowName:", event.WindowName)
    # print("Position:", event.Position)
    # print("Wheel:", event.Wheel)
    # print("Injected:", event.Injected)
    # print("---")
    is_write = True
    if FILTER_BY_WINDOW_NAME:
        is_write = False
        # print('FILTER_BY_WINDOW_NAME')
        for window_tag in FILTER_BY_WINDOW_NAME:
            # print(window_tag)
            if type(event.WindowName) is str and event.WindowName.find(str(window_tag)) != -1:
                is_write = True
                break
    if is_write:
        out_str = f'{event.MessageName},{event.WindowName},{event.Position},{"Ascii"},{"Key"},{"KeyID"}'
        write_file(out_str)
    # 返回 True 以便将事件传给其它处理程序     
    # 注意，这儿如果返回 False ，则鼠标事件将被全部拦截     
    # 也就是说你的鼠标看起来会僵在那儿，似乎失去响应了     
    return True


def onKeyboardEvent(event):
    # 监听键盘事件     
    # print("MessageName:", event.MessageName)
    # print("Message:", event.Message)
    # print("Time:", event.Time)
    # print("Window:", event.Window)
    # print("WindowName:", event.WindowName)
    # print("Ascii:", event.Ascii, chr(event.Ascii))
    # print("Key:", event.Key)
    # print("KeyID:", event.KeyID)
    # print("ScanCode:", event.ScanCode)
    # print("Extended:", event.Extended)
    # print("Injected:", event.Injected)
    # print("Alt", event.Alt)
    # print("Transition", event.Transition)
    # print("---")
    is_write = True
    if FILTER_BY_WINDOW_NAME:
        is_write = False
        for window_tag in FILTER_BY_WINDOW_NAME:
            if event.WindowName.find(str(window_tag)) != -1:
                is_write = True
                break
    if is_write:
        out_str = f'{event.MessageName},{event.WindowName},{"event.Position"},{event.Ascii, chr(event.Ascii)},{event.Key},{event.KeyID}'
        write_file(out_str)
    # 同鼠标事件监听函数的返回值     
    return True


def main():
    # 创建一个“钩子”管理对象     
    hm = PyHook3.HookManager()
    # 监听所有键盘事件     
    hm.KeyDown = onKeyboardEvent
    # 设置键盘“钩子”     
    hm.HookKeyboard()
    # 监听所有鼠标事件     
    hm.MouseAll = onMouseEvent
    # 设置鼠标“钩子”     
    hm.HookMouse()
    # 进入循环，如不手动关闭，程序将一直处于监听状态
    print('start')
    pythoncom.PumpMessages()
    print('end')


if __name__ == "__main__":
    main()
