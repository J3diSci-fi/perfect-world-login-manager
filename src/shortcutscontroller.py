import win32com.client
import json
import os

__shortcut_path = os.getcwd()+r"\shortcuts"

def criar_atalho(login,senha,nickname,icon_path):

    # Verifica se a pasta shortcuts existe, se não, cria-a
    if not os.path.exists(__shortcut_path):
        os.makedirs(__shortcut_path) 

    with open("./executable_path.json", "r") as json_file:
        data = json.load(json_file)

    caminho_element_executavel = data['path_executable']
    caminho_element = data['path_folder']

    argumentos = [
        f"startbypatcher",
        f"user:{login}",
        f"pwd:{senha}",
        f"role:{nickname}",
    ]

    shell = win32com.client.Dispatch("WScript.Shell")
    
    atalho = shell.CreateShortCut(os.path.join(__shortcut_path, f"{login}.lnk"))
    atalho.TargetPath = caminho_element_executavel
    
    if argumentos:
        atalho.Arguments = " ".join(argumentos)
    
    if caminho_element:
        atalho.WorkingDirectory = caminho_element
    
    if icon_path:
        atalho.IconLocation = f'{icon_path}'

    atalho.save()

    # Retorna o caminho completo para o atalho criado
    caminho_atalho = os.path.join(__shortcut_path, f"{nickname}.lnk")

    return caminho_atalho

def editar_atalho(nome_atalho, login,senha,nickname,icon):
    print(nome_atalho)
    excluir_atalho(nome_atalho)
    criar_atalho(login,senha,nickname,icon)

def excluir_atalho(nome_atalho):

    
    caminho_atalho = os.path.join(__shortcut_path, f"{nome_atalho}.lnk")

    if os.path.exists(caminho_atalho):
        os.remove(caminho_atalho)
        print(f"Atalho '{nome_atalho}' excluído com sucesso!")
    else:
        print(f"Atalho '{nome_atalho}' não encontrado em {__shortcut_path}")

def excluir_todos_atalhos():
    # Verifica se a pasta shortcuts existe
    if not os.path.exists(__shortcut_path):
        print(f"Pasta '{__shortcut_path}' não encontrada.")
        return False

    # Inicializa uma flag para verificar se algum atalho foi excluído
    excluiu_algum = False

    # Lista todos os arquivos no diretório __shortcut_path
    for arquivo in os.listdir(__shortcut_path):
        caminho_arquivo = os.path.join(__shortcut_path, arquivo)
        # Verifica se o arquivo é um atalho (extensão .lnk)
        if os.path.isfile(caminho_arquivo) and caminho_arquivo.endswith(".lnk"):
            os.remove(caminho_arquivo)
            print(f"Atalho '{arquivo}' excluído com sucesso!")
            excluiu_algum = True

    if excluiu_algum:
        print("Todos os atalhos foram excluídos.")
        return True
    else:
        print("Nenhum atalho encontrado para excluir.")
        return False