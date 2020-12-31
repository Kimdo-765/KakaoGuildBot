import time, win32con, win32api, win32gui, ctypes
import datetime
import ctypes
import pandas as pd

######data preset######
cache_text = ""
CF_TEXT = 1
kernel32 = ctypes.windll.kernel32
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32 = ctypes.windll.user32
user32.GetClipboardData.restype = ctypes.c_void_p

chat_command = "님이 들어왔습니다."
chat_command2 = "님이 나갔습니다."

main_c = '오로라 겨울길드(가입방)'
master = '겨울 운영진'
news = '1. 공지방 (겨울길드)'
flag = '2. 플래그 방 (겨울길드)'
suro = '3. 수로방 (겨울길드)'
suro2 = '3-1. 겨울 1수로'
suro3 = '3-2. 겨울 2수로'

chatroom_dict = {main_c : ["news.txt","news2.txt"], master : [], news : ["suro_alert.txt"], flag : ["flag.txt"], suro : ["suro.txt"], suro2 : ["suro.txt"], suro3 : ["suro.txt"]}

data = {main_c :[],master:[],news:[],flag:[],suro:[],suro2:[],suro3:[]}

for room_name, txt_name in chatroom_dict.items():
    for t in txt_name:
        f = open(t,'r',encoding='UTF8')
        data[room_name].append(f.read())
        f.close()
        
count = 0
news_set = set()
suro_set = set()
flag_set = set()

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con
##################

def get_clipboard_text():
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    user32.OpenClipboard(0)
    try:
        if user32.IsClipboardFormatAvailable(CF_TEXT):
            copy_data = user32.GetClipboardData(CF_TEXT)
            data_locked = kernel32.GlobalLock(copy_data)
            text = ctypes.c_char_p(data_locked)
            text_ = text.value
            kernel32.GlobalUnlock(data_locked)
            user32.CloseClipboard()
            return text_

    except:
        print("[-]Clipboard Copy Error... Retry")
        time.sleep(1)
        return get_clipboard_text()


def kakao_sendtext(chatroom_name, text):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RichEdit50W", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)


def PostKeyEx(hwnd, key, shift, specialkey):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    if IsWindow(hwnd):
        ThreadId = GetWindowThreadProcessId(hwnd, None)
        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)


def copy_chatroom(chatroom_name):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndListControl = win32gui.FindWindowEx(hwndMain, None, "EVA_VH_ListControl_Dblclk", None)
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ctext = get_clipboard_text()

    return ctext	


def SendReturn(hwnd):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def open_chatroom(chatroom_name):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    SendReturn(hwndkakao_edit3)
    time.sleep(0.5)


def chat_last_save(kakao_opentalk_name):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    open_chatroom(kakao_opentalk_name)
    ttext = copy_chatroom(kakao_opentalk_name)
    ttext = ttext.decode('cp949')
    #print(ttext)
    a = ttext.split('\r\n')   
    df = pd.DataFrame(a)

    try:
        return df.index[-2], df.iloc[-2,0]
    except:
        return -1,-1


def copy_routine(kakao_opentalk_name):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    try:
        open_chatroom(kakao_opentalk_name)
        ttext = copy_chatroom(kakao_opentalk_name)
        ttext = ttext.decode('cp949')
        return ttext
    except:
        return copy_routine(kakao_opentalk_name)


def chat_chek_command(kakao_opentalk_name,cls, clst):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    global cache_text

    ttext = copy_routine(kakao_opentalk_name)
    a = ttext.split('\r\n')  
    df = pd.DataFrame(a)

    if clst == df.iloc[-2,0]:
        print("[*]No chat found")
        return df.index[-2], df.iloc[-2,0]
    else:
        print("[*]New chat detected")
        df1 = df.iloc[cls+1:,0]

        found = df1[ df1.str.contains(chat_command, na=False) ]
        found2 = df1[ df1.str.contains(chat_command2, na=False) ]
        if 1<= int(found2.count()) and (kakao_opentalk_name != main_c):
            name_df = df1.str.extract('(.*)님이 나갔습니다.')
            name_df = name_df.dropna()
            captured_name = name_df.iloc[0]
            captured_name = captured_name.values[0]
            print("탈주감지")
            print(captured_name)
            open_chatroom(master)
            msg = "[도현봇]]\n"
            if(kakao_opentalk_name == news):
                msg += "[공지방 탈주 감지]\n"
            elif(kakao_opentalk_name == suro or kakao_opentalk_name == suro2 or kakao_opentalk_name == suro3):
                msg += "[수로방 탈주 감지]\n"
            elif(kakao_opentalk_name == flag):
                msg += "[플래그방 탈주 감지]\n"
            msg +=captured_name
            kakao_sendtext(master, msg)
            
        if 1 <= int(found.count()):
            name_df = df1.str.extract('(.*)님이 들어왔습니다.')
            name_df = name_df.dropna()
            captured_name = name_df.iloc[0]
            captured_name = captured_name.values[0]
            print(captured_name)
            if kakao_opentalk_name == main_c:
                send_all_text(kakao_opentalk_name)
            else :
                global suro_set
                global flag_set
                global news_set
                if kakao_opentalk_name == suro or kakao_opentalk_name == suro2 or kakao_opentalk_name == suro3: #수로방 입장
                    suro_set.update([captured_name])
                    send_all_text(kakao_opentalk_name)
                    print("[+]Send to Suro")
                elif kakao_opentalk_name == flag: #플래그방 입장
                    flag_set.update([captured_name])
                    send_all_text(kakao_opentalk_name)
                    print("[+]Send to Flag")
                elif kakao_opentalk_name == news: #공지방 입장 
                    news_set.update([captured_name])

                intersection = suro_set & flag_set & news_set
                print(suro_set,flag_set,news_set,intersection)
                if intersection: # 3개방 모두 입장했을 시 
                    print("[+]Level up request")
                    open_chatroom(master)
                    for name in intersection:
                        suro_set -= intersection
                        flag_set -= intersection
                        news_set -= intersection
                    name_list = list(intersection)
                    msg = "[도현봇]\n"
                    msg += "[가입 신청]\n"
                    msg += '\n'.join(name_list)
                    kakao_sendtext(master, msg)
                    msg2 ="[도현봇]\n"
                    msg2 += "환영합니다\n필수 오픈톡방 입장이 완료되었습니다.\n"
                    msg2 += "가입방은 나가주시면 됩니다.^^\n"
                    kakao_sendtext(main_c, msg2)
                    print("[*]Request End")
      
            return df.index[-2], df.iloc[-2,0]

        else:
            return df.index[-2], df.iloc[-2,0]

def send_all_text(index):
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    for item in data[index]:
        kakao_sendtext(index, item)
        time.sleep(1)

def main():
    ##################
    #12/28
    #kim do hyun
    #add comment
    ##################
    main_cls,main_clst = chat_last_save(main_c) #리팩토리 귀찮아서 유지
    news_cls,news_clst = chat_last_save(news)
    suro_cls,suro_clst = chat_last_save(suro)
    suro2_cls,suro2_clst = chat_last_save(suro2)
    suro3_cls,suro3_clst = chat_last_save(suro3)
    flag_cls,flag_clst = chat_last_save(flag) # ''
    time_flag = False
    print(main_cls,main_clst,news_cls,news_clst,suro_cls,suro_clst,suro2_cls,suro2_clst,suro3_cls,suro3_clst,flag_cls,flag_clst)
    
    while True:
        try:
            now = time.localtime()
            
            weekIndex = datetime.datetime.today().weekday()
            hourIndex = datetime.datetime.now().hour
            minuteIndex = datetime.datetime.now().minute
            print("[*] Time index %ds day %d : %d" % (weekIndex,hourIndex,minuteIndex))
            if(hourIndex == 20 and minuteIndex == 0) : #수로날 20시 정각에 공지방에 알림 전송
                if(weekIndex == 1) :
                    if(time_flag == False):
                        send_all_text(news)
                        time_flag = True
            else :
                if(time_flag == True):
                    time_flag = False
            print("%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec))
            print("[*]Running....")
            main_cls, main_clst = chat_chek_command(main_c,main_cls, main_clst)
            news_cls, news_clst = chat_chek_command(news,news_cls, news_clst)
            suro_cls, suro_clst = chat_chek_command(suro,suro_cls, suro_clst)
            suro2_cls,suro2_clst = chat_chek_command(suro2,suro2_cls,suro2_clst)
            suro3_cls,suro3_clst = chat_chek_command(suro3,suro3_cls,suro3_clst)
            flag_cls, flag_clst = chat_chek_command(flag,flag_cls, flag_clst)
        except:
            continue

if __name__ == '__main__':
    main()


