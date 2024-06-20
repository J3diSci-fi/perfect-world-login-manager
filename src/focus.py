import win32gui
import win32con
import win32com.client
import pythoncom

def ativar(hwnd):
    try:
        pythoncom.CoInitialize()  # Inicialize a biblioteca COM
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        print(f"Erro ao trazer a janela para o primeiro plano: {str(e)}")
    finally:
        pythoncom.CoUninitialize()  # Desinicialize a biblioteca COM

def desativar(hwnd):
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        print(f"Janela com hwnd {hwnd} desfocada.")
    except Exception as e:
        print(f"Erro ao desfocar a janela: {str(e)}")
