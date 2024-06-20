import os
import subprocess
import time
import json
import psutil
import pygetwindow as gw
import threading
import queue

__shortcut_path = os.getcwd() + r"\shortcuts"
window_info_file = "window_info.json"
last_hwnd = None
task_queue = queue.Queue()
is_processing = False  # Variável para controlar o processamento da fila

def exec_element(shortcut_name,nickname):
    global last_hwnd

    shortcut_name_path = __shortcut_path + f"\\{shortcut_name}.lnk"
    shell_process = subprocess.Popen(shortcut_name_path, shell=True)
    shell_pid = shell_process.pid

    # Aguardar um pouco para garantir que o processo filho seja iniciado
    time.sleep(2)

    # Encontrar o PID do processo relevante
    process_name_variants = ["elementclient.exe", "elementclient_64.exe", "ELEMENTCLIENT.exe"]
    target_pid = None

    for proc in psutil.process_iter(['pid', 'ppid', 'name']):
        if proc.info['ppid'] == shell_pid and proc.info['name'] in process_name_variants:
            target_pid = proc.info['pid']
            break

    if target_pid is None:
        return None  # Não encontrou o processo filho

    # Esperar até que a nova janela seja criada
    hwnd = None
    timeout = 30  # tempo máximo de espera em segundos
    interval = 1  # intervalo de verificação em segundos
    elapsed_time = 0

    while elapsed_time < timeout:
        windows = gw.getWindowsWithTitle("Perfect World Draco")
        if windows:
            current_hwnd = windows[0]._hWnd  # Pegar a primeira janela encontrada
            if current_hwnd != last_hwnd:
                hwnd = current_hwnd
                break
        time.sleep(interval)
        elapsed_time += interval

    if hwnd is None:
        return None  # Não encontrou a nova janela

    # Atualizar o arquivo JSON com as informações do processo
    update_window_info(nickname, target_pid, hwnd)

    # Atualizar o último hwnd
    last_hwnd = hwnd

    print(windows)
    print(hwnd)
    print(target_pid)

    return target_pid, hwnd

def process_queue():
    global is_processing
    while True:
        task = task_queue.get()
        shortcut_name, nickname = task
        if shortcut_name is None:
            break  # Encerrar o thread se encontrar um sinalizador de 
        exec_element(shortcut_name,nickname)
        task_queue.task_done()
    is_processing = False

def add_to_queue(shortcut_name,nickname):
    global is_processing
    task_queue.put((shortcut_name, nickname))
    if not is_processing:
        is_processing = True
        threading.Thread(target=process_queue, daemon=True).start()

def update_window_info(nickname, pid, hwnd):
    if os.path.exists(window_info_file):
        with open(window_info_file, 'r') as file:
            window_info = json.load(file)
    else:
        window_info = {}

    window_info[nickname] = {"pid": pid, "hwnd": hwnd, "status": 'on'}

    with open(window_info_file, 'w') as file:
        json.dump(window_info, file, indent=4)

def close_all_pws():
    process_names = ["elementclient.exe", "elementclient_64.exe", "ELEMENTCLIENT.exe"]

    for process_name in process_names:
        os.system(f"taskkill /f /im {process_name}")

def set_all_status_off():
    try:
        with open('window_info.json', 'r') as file:
            data = json.load(file)
        
        for key in data:
            data[key]['status'] = 'off'

        with open('window_info.json', 'w') as file:
            json.dump(data, file, indent=4)    
        
    except (FileNotFoundError, json.JSONDecodeError):
        print("Arquivo window_info.json não encontrado ou inválido.")

def check_process_status():
    while True:
        try:
            with open(window_info_file, 'r') as file:
                data = json.load(file)

            for key in data:
                pid = data[key]['pid']
                if not psutil.pid_exists(pid):
                    data[key]['status'] = 'off'

            with open(window_info_file, 'w') as file:
                json.dump(data, file, indent=4)

            time.sleep(2)  # Verificar a cada 2 segundos

        except (FileNotFoundError, json.JSONDecodeError):
            print("Arquivo window_info.json não encontrado ou inválido.")
            time.sleep(2)

def reset_json_window():
    try:
        with open('window_info.json', 'w') as file:
            json.dump({}, file, indent=4)    
        
    except (FileNotFoundError, json.JSONDecodeError):
        print("Arquivo window_info.json não encontrado ou inválido.")
    

threading.Thread(target=check_process_status, daemon=True).start()