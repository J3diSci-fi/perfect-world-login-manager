import win32gui
import win32con
import win32com.client

def ativar(hwnd):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.ShowWindow(hwnd,win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        print(f"Janela com hwnd {hwnd} trazida para o primeiro plano.")
    except Exception as e:
        print(f"Erro ao trazer a janela para o primeiro plano: {str(e)}")

def desativar(hwnd):
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        print(f"Janela com hwnd {hwnd} desfocada.")
    except Exception as e:
        print(f"Erro ao desfocar a janela: {str(e)}")

#ativar(13632986)