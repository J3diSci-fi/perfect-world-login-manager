import win32gui
import win32con
import win32process
import time

def get_vk_code_from_key(key):
    vk_codes = {
        '0': 48, #0 key
        '1': 49, #1 key
        '2': 50, #2 key
        '3': 51, #3 key
        '4': 52, #4 key
        '5': 53, #5 key
        '6': 54, #6 key
        '7': 55, #7 key
        '8': 56, #8 key
        '9': 57, #9 key
        'F1': 112,
        'F2': 113,
        'F3': 114,
        'F4': 115,
        'F5': 116,
        'F6': 117,
        'F7': 118,
        'F8': 119,
    }
    
    return vk_codes.get(key,None)

def enviar_tecla_shift_1(hwnd):

    if hwnd:
        # Enviar a combinação de teclas Shift + 1
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_SHIFT, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, ord('1'), 0)
        time.sleep(0.1)  # Aguardar um breve momento (opcional)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, ord('1'), 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_SHIFT, 0)
        print(f"Tecla Shift + 1 enviada para a janela com hwnd {hwnd} com sucesso!")
    else:
        print(f"Janela com hwnd {hwnd} não encontrada.")

def enviar_tecla(hwnd,key):

    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, get_vk_code_from_key(key), 0)
        time.sleep(0.1)  # Aguardar um breve momento (opcional)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, get_vk_code_from_key(key), 0)
        print(f"Tecla {key} enviada para a janela com hwnd {hwnd} com sucesso!")
    else:
        print(f"Janela com hwnd {hwnd} não encontrada.")