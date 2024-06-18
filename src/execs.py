import os
import subprocess

__shortcut_path = os.getcwd()+r"\shortcuts"

def exec_element(shortcut_name):

    shortcut_name = __shortcut_path +  f"\\{shortcut_name}.lnk"
    subprocess.Popen(shortcut_name,shell=True)

def close_all_pws():
    process_names = ["elementclient.exe", "elementclient_64.exe", "ELEMENTCLIENT.exe"]
    for process_name in process_names:
        os.system(f"taskkill /f /im {process_name}")