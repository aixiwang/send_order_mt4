# -*- coding:utf-8 -*- 
#--------------------------------------------------------------
# CLICK -- automatically send order by win32api, test with mt4
#
# Based on public code from internet
#
# Copyright by Aixi Wang <aixi.wang@hotmail.com>
#--------------------------------------------------------------
# [v0.1 2019.10.19]
# *added order_buy
# TODO: add order_close
#--------------------------------------------------------------
import win32gui, win32con, win32api
import time, math, random
from ctypes import *
import time


#---------------------
# _my_callback
#--------------------- 
def _my_callback(hwnd, extra):
    windows = extra
    temp=[]
    temp.append(hex(hwnd))
    temp.append(win32gui.GetClassName(hwnd))
    temp.append(win32gui.GetWindowText(hwnd))
    windows[hwnd] = temp

#---------------------
# get_child_windows
#---------------------
def get_child_windows(parent):        
    if not parent:         
        return      
    hwndChildList = []     
    win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd),  hwndChildList)          
    return hwndChildList
    
#---------------------
# test_enum_windows
#--------------------- 
def test_enum_windows(): 
    print("Enumerating all windows...")
    h=win32gui.FindWindow(None,'\xba\xec\xce\xe5')
    print(h)
    TestEnumWindows()
    print("All tests done!")
    TestEnumWindows()

#---------------------
# move_and_click
#---------------------
def move_and_click(x,y):
    win32api.SetCursorPos([x,y])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
  
#---------------------
# click_button_by_msg
#---------------------  
def click_button_by_msg(hwnd):
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, 0)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, 0)

#---------------------
# close_window
#---------------------
def close_window(classname,titlename):
    win32gui.PostMessage(win32lib.findWindow(classname, titlename), win32con.WM_CLOSE, 0, 0)
  
#---------------------
# VK_CODE
#---------------------
VK_CODE = {
    'backspace':0x08,
    'tab':0x09,
    'clear':0x0C,
    'enter':0x0D,
    'shift':0x10,
    'ctrl':0x11,
    'alt':0x12,
    'pause':0x13,
    'caps_lock':0x14,
    'esc':0x1B,
    'spacebar':0x20,
    'page_up':0x21,
    'page_down':0x22,
    'end':0x23,
    'home':0x24,
    'left_arrow':0x25,
    'up_arrow':0x26,
    'right_arrow':0x27,
    'down_arrow':0x28,
    'select':0x29,
    'print':0x2A,
    'execute':0x2B,
    'print_screen':0x2C,
    'ins':0x2D,
    'del':0x2E,
    'help':0x2F,
    '0':0x30,
    '1':0x31,
    '2':0x32,
    '3':0x33,
    '4':0x34,
    '5':0x35,
    '6':0x36,
    '7':0x37,
    '8':0x38,
    '9':0x39,
    'a':0x41,
    'b':0x42,
    'c':0x43,
    'd':0x44,
    'e':0x45,
    'f':0x46,
    'g':0x47,
    'h':0x48,
    'i':0x49,
    'j':0x4A,
    'k':0x4B,
    'l':0x4C,
    'm':0x4D,
    'n':0x4E,
    'o':0x4F,
    'p':0x50,
    'q':0x51,
    'r':0x52,
    's':0x53,
    't':0x54,
    'u':0x55,
    'v':0x56,
    'w':0x57,
    'x':0x58,
    'y':0x59,
    'z':0x5A,
    'numpad_0':0x60,
    'numpad_1':0x61,
    'numpad_2':0x62,
    'numpad_3':0x63,
    'numpad_4':0x64,
    'numpad_5':0x65,
    'numpad_6':0x66,
    'numpad_7':0x67,
    'numpad_8':0x68,
    'numpad_9':0x69,
    'multiply_key':0x6A,
    'add_key':0x6B,
    'separator_key':0x6C,
    'subtract_key':0x6D,
    'decimal_key':0x6E,
    'divide_key':0x6F,
    'F1':0x70,
    'F2':0x71,
    'F3':0x72,
    'F4':0x73,
    'F5':0x74,
    'F6':0x75,
    'F7':0x76,
    'F8':0x77,
    'F9':0x78,
    'F10':0x79,
    'F11':0x7A,
    'F12':0x7B,
    'F13':0x7C,
    'F14':0x7D,
    'F15':0x7E,
    'F16':0x7F,
    'F17':0x80,
    'F18':0x81,
    'F19':0x82,
    'F20':0x83,
    'F21':0x84,
    'F22':0x85,
    'F23':0x86,
    'F24':0x87,
    'num_lock':0x90,
    'scroll_lock':0x91,
    'left_shift':0xA0,
    'right_shift ':0xA1,
    'left_control':0xA2,
    'right_control':0xA3,
    'left_menu':0xA4,
    'right_menu':0xA5,
    'browser_back':0xA6,
    'browser_forward':0xA7,
    'browser_refresh':0xA8,
    'browser_stop':0xA9,
    'browser_search':0xAA,
    'browser_favorites':0xAB,
    'browser_start_and_home':0xAC,
    'volume_mute':0xAD,
    'volume_Down':0xAE,
    'volume_up':0xAF,
    'next_track':0xB0,
    'previous_track':0xB1,
    'stop_media':0xB2,
    'play/pause_media':0xB3,
    'start_mail':0xB4,
    'select_media':0xB5,
    'start_application_1':0xB6,
    'start_application_2':0xB7,
    'attn_key':0xF6,
    'crsel_key':0xF7,
    'exsel_key':0xF8,
    'play_key':0xFA,
    'zoom_key':0xFB,
    'clear_key':0xFE,
    '+':0xBB,
    ',':0xBC,
    '-':0xBD,
    '.':0xBE,
    '/':0xBF,
    '`':0xC0,
    ';':0xBA,
    '[':0xDB,
    '\\':0xDC,
    ']':0xDD,
    "'":0xDE,
    '`':0xC0}


class POINT(Structure):
    _fields_ = [("x", c_ulong),("y", c_ulong)]

#------------------
# get_mouse_point   
#------------------    
def get_mouse_point():
    po = POINT()
    windll.user32.GetCursorPos(byref(po))
    return int(po.x), int(po.y)

#------------------
# mouse_click   
#------------------
def mouse_click(x=None,y=None):
    if not x is None and not y is None:
        mouse_move(x,y)
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

#------------------
# mouse_dclick   
#------------------
def mouse_dclick(x=None,y=None):
    if not x is None and not y is None:
        mouse_move(x,y)
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

#------------------
# mouse_move    
#------------------
def mouse_move(x,y):
    windll.user32.SetCursorPos(x, y)

#------------------
# key_input    
#------------------
def key_input(str=''):
    for c in str:
        win32api.keybd_event(VK_CODE[c],0,0,0)
        win32api.keybd_event(VK_CODE[c],0,win32con.KEYEVENTF_KEYUP,0)
        time.sleep(0.01)

#------------------
# key_input2   
#------------------
def key_input2(c):
    win32api.keybd_event(VK_CODE[c],0,0,0)
    win32api.keybd_event(VK_CODE[c],0,win32con.KEYEVENTF_KEYUP,0)
    time.sleep(0.01)


#------------------
# test    
#------------------

def t0():
    pass


def t2():

    mouse_click(800,200)
    for c in 'hello':
        win32api.keybd_event(65,0,0,0)
        win32api.keybd_event(65,0,win32con.KEYEVENTF_KEYUP,0)
    #print get_mouse_point()


def t1():
    #mouse_move(1024,470)aa
    #time.sleep(0.05)
    #mouse_dclick()HELLO
    mouse_dclick(1024,470)
    

def t3():
    mouse_click(1024,470)
    str = 'hello'
    for c in str:
        win32api.keybd_event(VK_CODE[c],0,0,0) 
        win32api.keybd_event(VK_CODE[c],0,win32con.KEYEVENTF_KEYUP,0)
        time.sleep(0.01)


def t4():
    mouse_click(1024,470)
    str = 'hello'
    key_input(str)
  
#---------------------
# order_buy
#---------------------     
def order_buy(vol,up,down):
    try:
        print('buy:',vol,up,down)
        windows = {}
        win32gui.EnumWindows(_my_callback, windows)
        print("Enumerated a total of  windows with %d classes" ,(len(windows)))
        #print('------------------------------')
        #print classes
        #print '-------------------------------'
        for item in windows:
            if (windows[item][1].find('Meta') >= 0): 
                
                #print(windows[item])
                win32gui.SetForegroundWindow(item)
                key_input2('F9')
                time.sleep(1)

                #left, top, right, bottom = win32gui.GetWindowRect(item)
                #print('win pos:',left,top,right,bottom)
                windows2 = {}
                win32gui.EnumWindows(_my_callback, windows2)
                for item2 in windows2:
                    #print(windows2[item2])
                    #if (windows[item2][1].find('Dialog box') >= 0):
                    if windows2[item2][2] == '订单':
                        print('订单------------------')                
                        
                        hwndChildList = get_child_windows(item2)
                        #print('hwndChildList:',hwndChildList)
                        print('---------step 1------------')
                        # step1: fill vol
                        i = 0
                        for subitem in hwndChildList:
                            title = win32gui.GetWindowText(subitem)     
                            clsname = win32gui.GetClassName(subitem)  
                            #print(title,clsname)
                            if clsname == 'Edit' and i == 0:
                                win32gui.SendMessage(subitem,win32con.WM_SETTEXT,None,str(vol))
                                i = 1
                                continue
                                
                            if clsname == 'Edit' and i == 1:
                                win32gui.SendMessage(subitem,win32con.WM_SETTEXT,None,str(down))
                                i = 2
                                continue
                            
                            if clsname == 'Edit' and i == 2:
                                win32gui.SendMessage(subitem,win32con.WM_SETTEXT,None,str(up))
                                i = 3
                                break
                        
                        if i != 3:
                            return -1
                        # step4: click
                        
                        for subitem in hwndChildList:
                            title = win32gui.GetWindowText(subitem)     
                            clsname = win32gui.GetClassName(subitem)  
                            if title == '于市价买':
                                #print(subitem,title,clsname)
                                click_button_by_msg(subitem)
                                
                        time.sleep(3)
                        # step5: confirm execution
                        print('------------ confirm execution ----------')
                        windows2 = {}
                        err_index = 0
                        win32gui.EnumWindows(_my_callback, windows2)
                        for item2 in windows2:
                            if windows2[item2][2] == '订单':            
                                hwndChildList2 = get_child_windows(item2)

                                for subitem2 in hwndChildList2:
                                    title = win32gui.GetWindowText(subitem2)     
                                    clsname = win32gui.GetClassName(subitem2)  
                                    print(title,clsname)
                                    if title == '请检查操作参数再重试.' and clsname == 'Static':
                                        print('found:','请检查操作参数再重试')
                                        err_index += 1
                                    
                                    if title == 'OK' and clsname == 'Button':
                                        print('found ok button, click it')
                                        click_button_by_msg(subitem2)
                                        
                        if err_index > 0:
                            return -2
                        else:
                            return 0
                            
        return -3    
    except Exception as e:
        print('order_buy exception:',str(e))
        return -4
        
#---------------------
# order_close
#---------------------     
def order_close():
    return -1
    
    
#-------------------------
# main
#-------------------------
if __name__ == "__main__":
    retcode = order_buy(0.01,1491.1,1490.2)
    print('retcode:',retcode)
    
    