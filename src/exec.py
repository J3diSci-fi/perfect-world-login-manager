import os
import subprocess
import threading
import psutil
import time
import win32gui
import win32process
import json

current_state = []
flag_treeview = True

def find_launcher():
    # Lê o arquivo JSON para obter o caminho do executável e o diretório
    with open('executable_path.json', 'r') as f:
        dados = json.load(f)
        caminho_executavel = dados["path_executable"]  # Obtém o caminho do executável
        diretorio_atual = dados["path_folder"]  # Obtém o diretório atual do JSON

    # Verifica até três diretórios anteriores ao do projeto
    for i in range(3):  # Verifica até três diretórios anteriores
        diretorio_anterior = os.path.abspath(os.path.join(diretorio_atual, '..', '..' * i))
        caminho_launcher = os.path.join(diretorio_anterior, 'launcher')

        print(diretorio_anterior)
        
        # Verifica se a pasta "launcher" existe
        if os.path.isdir(caminho_launcher):
            # Executa o executável do caminho lido do JSON
            if os.path.isfile(caminho_executavel):
                subprocess.run(caminho_executavel)
                print(f"Executando {caminho_executavel}")
                return
            
    print("Pasta 'launcher' não encontrada nos diretórios anteriores.")

def exec_launcher():
    # Cria uma nova thread para executar a função verificar_e_executar_launcher
    thread = threading.Thread(target=find_launcher)
    thread.start()

def verify_log_execs(flag):

    while flag:
        # Defina os nomes dos processos que deseja procurar
        process_names = ["elementclient_32.exe", "elementclient_64.exe", "elementclient.exe", "ELEMENTCLIENT.EXE"]

        # Cria uma lista para rastrear logins ativos
        active_logins = []  # Alterado de set para lista

        # Itera sobre todos os processos ativos no sistema
        for process in psutil.process_iter(attrs=["pid", "name", "cmdline"]):
            try:
                # Verifica se o nome do processo corresponde a algum dos nomes desejados
                if process.info["name"] in process_names:
                    pid = process.info["pid"]
                    cmdline = process.info["cmdline"]  # Lista de argumentos do processo
                    login = cmdline[4] if len(cmdline) > 4 else None  # Supondo que o login esteja no segundo argumento
                    if login:
                        login = login.split(':', 1)[1].strip()
                        active_logins.append(login)  # Adiciona o login à lista de logins ativos
                        # Verifica se já existe uma entrada com o mesmo login e pid
                        if not any(entry[0] == login for entry in current_state):  # Verifica apenas o login
                            current_state.append([login, pid])  # Adiciona à current_state se não estiver presente
                    print(cmdline)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                # Ignora processos que foram encerrados ou aos quais não temos acesso
                continue

        # Remove logins que não estão mais ativos
        for entry in current_state[:]:  # Usar uma cópia da lista para evitar problemas de modificação
            if entry[0] not in active_logins:  # Verifica apenas o login
                current_state.remove(entry)  # Remove da current_state se não estiver mais ativo

        print(current_state)
        time.sleep(5)
        find_hwnd_window()

def find_hwnd_window():
    global current_state
    """Encontra a janela principal de processos em current_state."""
    for i, entry in enumerate(current_state):  # Itera sobre os logins em current_state
        pid = entry[1]  # Obtém o PID do entry
        def callback(hwnd, pid):
            # Obtém o PID do processo associado a esta janela
            tid, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            # Verifica se o PID da janela corresponde ao PID especificado
            if window_pid == pid:  # Verifica se o PID é igual
                window_title = win32gui.GetWindowText(hwnd)  # Obtém o título da janela
                # Ignora títulos específicos e None
                if window_title and window_title not in ["AXWIN Frame Window", "about:blank", "MSCTFIME UI", "Default IME"]:
                    if hwnd not in current_state[i]:
                        current_state[i].append(hwnd)
                        
        win32gui.EnumWindows(callback, pid)  # Chama EnumWindows para cada PID em current_state

def state_elements(master=None):
    # Cria uma nova thread para executar o loop de verificação
    thread = threading.Thread(target=verify_log_execs,args=(True,))
    thread.daemon = True
    thread.start()

def exec_shortcut(login):
    """Executa o atalho correspondente ao login na pasta 'shortcuts'."""
    # Define o caminho do atalho na pasta 'shortcuts' que está um nível acima de 'src'
    shortcut_path = os.path.join(os.getcwd(), 'shortcuts', f"{login}.lnk")
    
    # Verifica se o atalho existe
    if os.path.isfile(shortcut_path):
        subprocess.Popen(shortcut_path, shell=True)  # Adicionado shell=True
        print(f"Executando atalho: {shortcut_path}")
    else:
        print(f"Atalho não encontrado: {shortcut_path}")

def start_shortcut(login):
    thread = threading.Thread(target=exec_shortcut,args=(login,))
    thread.start()

class TreeviewManager:
    def __init__(self, master):
        self.master = master
        self.flag_treeview = False  # Controle da thread

    def start(self):
        self.flag_treeview = True
        thread = threading.Thread(target=self.manager_treeview)
        thread.daemon = True
        thread.start()

    def stop(self):
        self.flag_treeview = False

    def manager_treeview(self):
        while self.flag_treeview:
            self.master.atualizar_treeview()
            time.sleep(1)