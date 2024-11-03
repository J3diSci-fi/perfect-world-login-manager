from tkinter import StringVar
from tkinter import ttk
import customtkinter as ctk
from CTkMenuBar import *
from CTkTable import *
from CTkMessagebox import *
from tkinter import filedialog, messagebox
from PIL import Image
import json
import os
from src.shortcutscontroller import criar_atalho,editar_atalho,excluir_atalho,excluir_todos_atalhos,atualizar_atalhos
from src.execs import current_state,close_all_pws,start_shortcut,exec_launcher,state_elements,TreeviewManager
import threading
import time
import keyboard
from src.focus import ativar
from src.actions import enviar_tecla
import pygame
import random

browse_image = ctk.CTkImage(Image.open("./res/search.png"), size=(20, 20))
visible_on = ctk.CTkImage(Image.open("./res/visibility_icon.png"), size=(20, 20))
visible_off = ctk.CTkImage(Image.open("./res/off_visibility_icon.png"), size=(20, 20))
backgrond_image = ctk.CTkImage(Image.open("./res/background.png"),size=(350,350))
seta_cima = ctk.CTkImage(Image.open("./res/seta-para-cima.png"),size=(16,16))
seta_baixo = ctk.CTkImage(Image.open("./res/seta-para-baixo.png"),size=(16,16))
seta_direita = ctk.CTkImage(Image.open("./res/seta-direita.png"),size=(16,16))
seta_esquerda = ctk.CTkImage(Image.open("./res/seta-esquerda.png"),size=(16,16))
confirm = ctk.CTkImage(Image.open("./res/confirm.png"),size=(16,16))
cancel = ctk.CTkImage(Image.open("./res/cancel.png"),size=(16,16))
back = ctk.CTkImage(Image.open("./res/back.png"),size=(20,20))
maximizar = ctk.CTkImage(Image.open("./res/maximizar.png"),size=(24,24))
botao_de_informacao = ctk.CTkImage(Image.open("./res/botao-de-informacao.png"),size=(48,48))

emoji1 = ctk.CTkImage(Image.open("./res/emoji1.png"),size=(96,96))
racas_humanos = ctk.CTkImage(Image.open("./res/racas_humanos.png"),size=(120,120))

class Root(ctk.CTk):

    def __init__(self):
        ctk.set_appearance_mode("light")
        super().__init__()

        self.__windowcfg()
        self.__elements()
        self.check_existing_executable()

        self.mainloop()

    def __windowcfg(self):
        self.title("Selecionar Executável")
        self.resizable(False, False)

        window_width = 460
        window_height = 100
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __elements(self):
        label = ctk.CTkLabel(self, text="Executável:")
        label.grid(row=0, column=0, padx=10, pady=10)

        self.entry = ctk.CTkEntry(self, width=300, state="readonly")
        self.entry.grid(row=0, column=1, padx=10, pady=10)

        self.browse_button = ctk.CTkButton(self, image=browse_image, text="", command=self.browse_file, width=10)
        self.browse_button.grid(row=0, column=2, padx=10, pady=10)

        self.confirm_button = ctk.CTkButton(self, text="Confirmar", command=self.confirm)
        self.confirm_button.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    def browse_file(self):
        file_path = filedialog.askopenfilename(initialdir='./',filetypes=[("Executáveis", "*.exe")])
        if file_path:   
            valid_files = ["ELEMENTCLIENT.exe", "elementclient_64.exe", "elementclient.exe"]
            if any(file_path.endswith(valid_file) for valid_file in valid_files):
                self.entry.configure(state="normal")
                self.entry.delete(0, ctk.END)  # Limpa o campo antes de inserir o novo texto
                self.entry.insert(0, file_path)
                self.entry.configure(state="readonly")
            else:
                CTkMessagebox(title="Info", message="Selecione um executável:\n-->elementclient.exe\n-->ELEMENTCLIENT.EXE\n-->elementclient_64.exe")

    def confirm(self):
        exe_path = self.entry.get()
        folder_path = os.path.dirname(exe_path)

        if exe_path:
            data = {"path_executable": exe_path,
                    "path_folder":folder_path}
            with open("executable_path.json", "w") as json_file:
                json.dump(data, json_file, indent=4)
            
            atualizar_atalhos()
            self.open_Manager()
        else:
            CTkMessagebox(title="Erro", message="Selecione o executável.", icon="cancel")

    def check_existing_executable(self):

        if os.path.exists("executable_path.json"):
            with open("executable_path.json", "r") as json_file:
                data = json.load(json_file)
                if "path_executable" in data and os.path.isfile(data["path_executable"]):
                    self.open_Manager()

    def open_Manager(self):
        self.withdraw()  # Esconde a janela atual
        Manager(self)

class Manager(ctk.CTkToplevel):

    def __init__(self, master):
        super().__init__(master)
        pygame.mixer.init()
        self.master = master

        self.__windowcfg()
        self.__elements()
        self.__frameImage()
        self.__frameAddAccount()
        self.__frameTable()
        self.__framebottom_table()

        self.after(100, self.create_combo_view)  # Chama o método após 100ms
        state_elements()

    def create_combo_view(self):
        self.comboView = ComboRoot(self)  # Atribui a nova instância

    def close_all(self):
        self.destroy()
        self.master.destroy()  # Fecha a janela principal (Root)

    def __windowcfg(self):
        self.title("Login Manager")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close_all)

        window_width = 680
        window_height = 600
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __elements(self):
        menu = CTkMenuBar(master=self)

        # Menu Combo
        menu.add_cascade("Combo",command=self.__comboWindow)

        # Menu Atualizar PW
        atualizar_pw_menu = menu.add_cascade("Atualizar PW")
        atualizar_pw_dropdown = CustomDropdownMenu(widget=atualizar_pw_menu)
        atualizar_pw_dropdown.add_option(option="Abrir Launcher", command=self.open_launcher)

        # Menu Config
        config_menu = menu.add_cascade("Config")
        config_dropdown = CustomDropdownMenu(widget=config_menu)
        config_dropdown.add_option(option="Binds", command=self.open_binds)  # Substitua com a ação desejada
        config_dropdown.add_option(option="Mudar Caminho Executável", command=self.change_executable_path)
        config_dropdown.add_option(option="Resetar Manager", command=self.reset_app)  # Substitua com a ação desejada

    def open_launcher(self):
        exec_launcher()

    def open_binds(self):
        BindRoot()

    def combo_action(self):
        # Action to perform when Combo is selected
        messagebox.showinfo("Combo", "Combo action selected")

    def change_executable_path(self):
        # Action to perform when Mudar Caminho Executável is selected
        self.master.deiconify()  # Show the main window to change executable path
        self.destroy()  # Close the Manager window
    
    def __frameAddAccount(self):
        frame = ctk.CTkFrame(self)
        frame.place(x=10, y=40)
        
        login_label = ctk.CTkLabel(frame, text="Login:")
        login_label.grid(row=0, column=0, padx=10, pady=10)

        self.login_entry = ctk.CTkEntry(frame)
        self.login_entry.grid(row=0, column=1,columnspan=2, sticky="ew", padx=10, pady=10)

        password_label = ctk.CTkLabel(frame, text="Senha:")
        password_label.grid(row=1, column=0, padx=10, pady=10)

        self.password_visible = False
        self.password_entry = ctk.CTkEntry(frame, show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)
        

        self.toggle_button = ctk.CTkButton(frame, image = visible_on, text="", command=self.toggle_password_visibility,width=10)
        self.toggle_button.grid(row=1, column=2, padx=10, pady=10)

        nickname_label = ctk.CTkLabel(frame, text="Personagem:")
        nickname_label.grid(row=2, column=0, padx=10, pady=10)

        self.nickname_entry = ctk.CTkEntry(frame)
        self.nickname_entry.grid(row=2, column=1,columnspan=2, sticky="ew", padx=10, pady=10)

        icon_label = ctk.CTkLabel(frame, text="Ícone:")
        icon_label.grid(row=3, column=0, padx=10, pady=10)

        self.icon_path_label = ctk.CTkLabel(frame, text="")
        self.icon_path_label.grid(row=3, column=1, padx=10, pady=10)

        select_icon_button = ctk.CTkButton(frame, image=browse_image, text="", command=self.select_icon,width=10)
        select_icon_button.grid(row=3, column=2, padx=10, pady=10)

        add_button = ctk.CTkButton(frame, text="Adicionar", command=self.add_account)
        add_button.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    def __frameTable(self):
        frame = ctk.CTkScrollableFrame(self, width=317, height=400)
        frame.place(x=330, y=40)

        # Carregar dados do arquivo JSON
        try:
            with open('data.json', 'r') as f:
                data = json.load(f)['accounts']
        except FileNotFoundError:
            data = []

        # Cabeçalhos das colunas
        headers = ["Login", "Nome do Personagem"]

        # Formatando os dados para a tabela
        formatted_data = [headers]
        for account in data:
            formatted_data.append([account['login'], account['nickname']])

        # Criando a tabela
        self.table = CTkTable(master=frame, values=formatted_data, padx=10, pady=10, command=self.__TableEdit)
        self.table.grid(row=0, column=0)
    
    def __framebottom_table(self):

        frame = ctk.CTkFrame(self, width=317, height=400)
        frame.place(x=330, y=465)
        
        self.close_all = ctk.CTkButton(frame, text= "Fechar todos os PW's",width=307,command=self.__close_pws)
        self.close_all.grid(row=1,column=0,padx=10,pady=10)

        self.trolei = ctk.CTkButton(frame, text= "Não clique aqui !!!",width=307,command=self.__play_sound)
        self.trolei.grid(row=2,column=0,padx=10,pady=10)

    def __play_sound(self):
        # Lista todos os arquivos MP3 na pasta 'res'
        mp3_files = [file for file in os.listdir('./res') if file.endswith('.mp3')]
        
        # Escolhe um arquivo MP3 aleatoriamente
        if mp3_files:
            random_mp3 = random.choice(mp3_files)
            pygame.mixer.music.load(f'./res/{random_mp3}')
            pygame.mixer.music.play()

    def __frameImage(self):

        label_image = ctk.CTkLabel(self,text='',image=backgrond_image)
        label_image.place(x=-10,y=280)

        label_creator = ctk.CTkLabel(self,text='by:hckzn')
        label_creator.place(x=10,y=570)

    def __updateTableAdd(self,login,nickname):
        # Carregar dados do arquivo JSON
        try:
            with open('data.json', 'r') as f:
                data = json.load(f)['accounts']
        except (FileNotFoundError, json.JSONDecodeError):
            data = []

        next_row = len(data)
        new_user = [login,nickname]

        self.table.add_row(new_user,next_row+1)

    def updateTableEdit(self,index,login,nickname):

        self.table.insert(index + 1, 0, login)
        self.table.insert(index + 1, 1, nickname)

    def __TableEdit(self, cell_data):

        try:
            with open('data.json', 'r') as f:
                data = json.load(f)['accounts']
        except (FileNotFoundError, json.JSONDecodeError):
            data = []

        previusRow= self.table.get_selected_row()['row_index']
        if previusRow is not None:
            self.table.deselect_row(previusRow)

        if cell_data['row'] != 0:
            previusRow= self.table.get_selected_row()['row_index']
            currentRow= cell_data['row']

            row_current_data = self.table.get_row(currentRow)
            self.table.select_row(currentRow)

            msg = CTkMessagebox(title="Oque deseja fazer?", message="Oque deseja fazer?",
                        icon="info", option_1="Abrir", option_2="Editar",option_3='Excluir')

            if msg.get() == 'Abrir':
                start_shortcut(row_current_data[0])

            elif msg.get() == 'Excluir':  # Opção "Excluir" selecionada
                # Encontrar e remover a conta correspondente no JSON
                for index, conta in enumerate(data):
                    if conta['login'] == row_current_data[0]:
                        del data[index]
                        break
                
                # Atualizar o arquivo JSON
                try:
                    with open('data.json', 'w') as f:
                        json.dump({'accounts': data}, f, indent=4)
                except IOError:
                    messagebox.showerror("Erro", "Erro ao salvar dados no arquivo JSON.")
                    return
                
                # Atualizar a tabela
                self.table.delete_row(currentRow)

                excluir_atalho(row_current_data[0])
                
            elif msg.get() == "Editar":
                row_index = ''
                login = ''
                password = ''
                nickname = ''
                icon_path = ''

                for index, conta in enumerate(data):
                    if conta['login'] == row_current_data[0]:
                        row_index = index
                        login = conta['login']
                        password = conta['password']
                        nickname = conta['nickname']
                        icon_path = conta['path_icon']

                EditLogin(self,login,password,nickname,icon_path,row_index)
            
            else:  # Opção "Editar" selecionada
                msg.destroy()
                
    def toggle_password_visibility(self):
        if self.password_visible:
            self.password_entry.configure(show="*")
            self.toggle_button.configure(image=visible_on)
        else:
            self.password_entry.configure(show="")
            self.toggle_button.configure(image=visible_off)
        self.password_visible = not self.password_visible
        
    def select_icon(self):
        self.currentPath = None
        icon_path = filedialog.askopenfilename(initialdir= './icons', filetypes=[("Selecione o ícone", "*.ico")])
        if icon_path:
            image = ctk.CTkImage(Image.open(icon_path), size=(32, 32))
            self.icon_path_label.configure(image=image, text='')
            self.currentPath = icon_path      
    
    def add_account(self):
        login = self.login_entry.get()
        password = self.password_entry.get()
        nickname = self.nickname_entry.get()

        if not login or not password or not nickname:
            CTkMessagebox(title="Erro", message="Por favor, preencha todos os campos.", icon="cancel")
            return
        
        try:
            icon_path = self.currentPath
            if not icon_path:
                CTkMessagebox(title="Erro", message="Por favor, selecione um ícone para o executável.", icon="cancel")
                return
        except AttributeError:
            CTkMessagebox(title="Erro", message="Por favor, selecione um ícone para o executável.", icon="cancel")
            return

        # Criando um dicionário com os dados da nova conta
        new_account = {
            'login': login,
            'password': password,
            'nickname': nickname,
            'path_icon': icon_path
        }

        # Lendo os dados atuais do arquivo JSON, se existir
        try:
            with open('data.json', 'r') as f:
                data = json.load(f)
                accounts = data.get('accounts', [])
        except FileNotFoundError:
            accounts = []
        except json.JSONDecodeError:
            accounts = []

        # Verificando se o login já existe
        if any(account['login'] == login for account in accounts):
            CTkMessagebox(title="Info", message="Login já cadastrado!")
            return

        # Adicionando a nova conta à lista existente de contas
        accounts.append(new_account)

        # Atualizando os dados no arquivo JSON
        try:
            with open('data.json', 'w') as f:
                json.dump({'accounts': accounts}, f, indent=4)
        except IOError:
            messagebox.showerror("Erro", "Erro ao salvar dados no arquivo JSON.")
            return

        # Exibindo mensagem de sucesso
        CTkMessagebox(title='Sucesso',message=f"Conta adicionada:\nLogin: {login}\nPersonagem:{nickname}\nÍcone: {icon_path}",
                  icon="check", option_1="Ok")

        self.__updateTableAdd(login,nickname)

        criar_atalho(login,password,nickname,icon_path)

    def __close_pws(self):

        msg = CTkMessagebox(title="Info", message="Está ação irá fechar todas instâncias de PW\n\nDeseja continuar?",
                    icon="info", option_1="Sim", option_2="Não")
        
        if msg.get()=="Sim":
            close_all_pws()

    def reset_app(self):
        
        msg = CTkMessagebox(title="Resetar App", message="Está ação irá apagar todos os dados!",
                  icon="warning", option_1="Sim", option_2="Não")
    
        if msg.get()=="Sim":
            excluir_todos_atalhos()
            for file in ["data.json", "executable_path.json"]:
                file_path = os.getcwd() + f'\\{file}'
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Arquivo '{file}' excluído com sucesso.")
            
            self.master.deiconify()
            self.destroy()
        
    def __comboWindow(self):
        self.withdraw()
        self.comboView.deiconify()

class EditLogin(ctk.CTkToplevel):
    def __init__(self, master, login='None',password='None',nickname='None',icon_path='None',row_index=0):
        super().__init__(master)
        self.master = master
        self.login = login
        self.password = password
        self.nickname = nickname
        self.currentPath = icon_path
        self.row_index = row_index

        self.grab_set()
        self.focus()

        self.__windowcfg()
        self.__frameAddAccount()

    def close_all(self):
        self.destroy()  # Fecha a janela principal (Root)

    def __windowcfg(self):
        self.title(f"Editar {self.nickname}")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close_all)

        window_width = 318
        window_height = 270
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __frameAddAccount(self):
        self.login_var = StringVar(value=self.login)
        self.password_var = StringVar(value=self.password)
        self.nickname_var = StringVar(value=self.nickname)
        
        frame = ctk.CTkFrame(self)
        frame.place(x=10, y=10)

        login_label = ctk.CTkLabel(frame, text="Login:")
        login_label.grid(row=0, column=0, padx=10, pady=10)

        self.login_entry = ctk.CTkEntry(frame,textvariable=self.login_var)
        self.login_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=10, pady=10)

        password_label = ctk.CTkLabel(frame, text="Senha:")
        password_label.grid(row=1, column=0, padx=10, pady=10)

        self.password_visible = False
        self.password_entry = ctk.CTkEntry(frame, show="*",textvariable=self.password_var)
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        self.toggle_button = ctk.CTkButton(frame, image=visible_on, text="", command=self.toggle_password_visibility, width=10)
        self.toggle_button.grid(row=1, column=2, padx=10, pady=10)

        nickname_label = ctk.CTkLabel(frame, text="Nickname:")
        nickname_label.grid(row=2, column=0, padx=10, pady=10)

        self.nickname_entry = ctk.CTkEntry(frame,textvariable=self.nickname_var)
        self.nickname_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=10, pady=10)

        icon_label = ctk.CTkLabel(frame, text="Ícone:")
        icon_label.grid(row=3, column=0, padx=10, pady=10)

        image = ctk.CTkImage(Image.open(f"{self.currentPath}"), size=(32, 32))
        self.icon_path_label = ctk.CTkLabel(frame, image=image,text="")
        self.icon_path_label.grid(row=3, column=1, padx=10, pady=10)

        select_icon_button = ctk.CTkButton(frame, image=browse_image, text="", command=self.select_icon, width=10)
        select_icon_button.grid(row=3, column=2, padx=10, pady=10)

        add_button = ctk.CTkButton(frame, text="Alterar", command=self.edit_account)
        add_button.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    def toggle_password_visibility(self):
        if self.password_visible:
            self.password_entry.configure(show="*")
            self.toggle_button.configure(image=visible_on)
        else:
            self.password_entry.configure(show="")
            self.toggle_button.configure(image=visible_off)
        
        self.password_visible = not self.password_visible

    def select_icon(self):
        icon_path = filedialog.askopenfilename(initialdir= './icons', filetypes=[("Selecione o ícone", "*.ico")])
        if icon_path:
            image = ctk.CTkImage(Image.open(icon_path), size=(32, 32))
            self.icon_path_label.configure(image=image, text='')
            self.currentPath = icon_path  

    def edit_account(self):
        login = self.login_entry.get()
        password = self.password_entry.get()
        nickname = self.nickname_entry.get()
        if not login:
            CTkMessagebox(title="Login", message="Por favor, preencha o campo de Login!!!", icon="cancel")
            return        
        if not password:
            CTkMessagebox(title="Password", message="Por favor, preencha o campo de Password!!!", icon="cancel")
            return
        if not nickname:
            CTkMessagebox(title="Nickname", message="Por favor, preencha o campo de Nickname!!!", icon="cancel")
            return       
        try:
            icon_path = self.currentPath
        except AttributeError:
            CTkMessagebox(title="Ícone", message="Por favor, selecione um ícone.\nMesmo que já esteja inicialemnte mostrando!!!", icon="cancel")
            return

        try:
            with open('data.json', 'r') as f:
                data = json.load(f)
                accounts = data.get('accounts', [])
        except (FileNotFoundError, json.JSONDecodeError):
            accounts = []

        accounts[self.row_index] = {
            'login': login,
            'password': password,
            'nickname': nickname,
            'path_icon': icon_path
        }

        try:
            with open('data.json', 'w') as f:
                json.dump({'accounts': accounts}, f, indent=4)
        except IOError:
            messagebox.showerror("Erro", "Erro ao salvar dados no arquivo JSON.")
            return


        self.master.updateTableEdit(self.row_index,login,nickname)

        editar_atalho(self.login,login,password,nickname,icon_path)

        self.destroy()

class ComboRoot(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.withdraw()
        self.master = master
        self.__windowcfg()
        self.__treeview()
        self.__treeview2()
        self.__buttons()
        self.__frameIntervalo1()
        self.__frameIntervalo2()
        self.__teclaunica()
        
        self.previous_state = []
        
        self.list_second_tree = []

        # Instancia o gerenciador de Treeview
        self.treeview_manager = TreeviewManager(self)
        self.treeview_manager.start()

        self.vieworder = viewOrder(self,self.list_second_tree,self.combobox_intervalo1,self.entry_intervalo1,self.combobox_intervalo2,
                  self.entry_intervalo2,self.activate_checkbox_loop,self.tecla_combobox_loop)
        self.vieworder.withdraw()

    def close_all(self):
        self.master.deiconify()
        self.withdraw()

    def __windowcfg(self):
        self.title("Combar")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close_all)

        self.window_width = 702
        self.window_height = 560
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - self.window_width) // 2
        y = (screen_height - self.window_height) // 2
        self.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))
    
    def __buttons(self):

        # Adicionando os novos botões
        self.btn_left = ctk.CTkButton(self,width=50,image=seta_esquerda, text="", command=self.__leftButton)
        self.btn_left.place(x=320, y=180)  # Posição do botão à esquerda

        self.btn_right = ctk.CTkButton(self,width=50,image=seta_direita, text="", command=self.__rightButton)
        self.btn_right.place(x=320, y=150)  # Posição do botão à direita

        self.btn_up = ctk.CTkButton(self.frame2,width=50,image=seta_cima, text="", command=self.__upButton)
        self.btn_up.place(x=85, y=355)  # Posição do botão à esquerda

        self.btn_down = ctk.CTkButton(self.frame2,width=50,image=seta_baixo, text="", command=self.__downButton)
        self.btn_down.place(x=145, y=355)  # Posição do botão à direita

        self.btn_maximizer = ctk.CTkButton(self,width=10, image=maximizar, text='',command=self.__buttonOrder)
        self.btn_maximizer.place(x=self.window_width-50,y=self.window_height - 43)

        self.btn_back = ctk.CTkButton(self,width=10, image=back, text='',command=self.close_all)
        self.btn_back.place(x=10,y=self.window_height - 38)

    def __treeview(self):
        label_img = ctk.CTkLabel(self,text='',image=racas_humanos)
        label_img.place(x=285,y=15)

        # Frame para o primeiro Treeview
        frame1 = ctk.CTkFrame(self, width=280, height=350)  # Defina o tamanho aqui
        frame1.place(x=10, y=10)  # Ajuste conforme necessário

        # Criação do Treeview para exibir contas com tema escuro
        label = ctk.CTkLabel(frame1, text='CONTAS ABERTA(S)')
        label.place(x=80, y=10)

        self.tree = ttk.Treeview(frame1, columns=("Nickname",), show="headings", style="Dark.Treeview")
        self.tree.heading("Nickname", text="Nickname")
        
        # Centraliza os dados na coluna
        self.tree.column("Nickname", anchor="center")  # Centraliza a coluna "Nickname"
        
        # Adicionando um scrollbar
        self.scrollbar = ttk.Scrollbar(frame1, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        
        self.tree.place(x=15, y=40, width=240, height=300)
        self.scrollbar.place(x=250, y=40, height=300)  # Posiciona o scrollbar ao lado da tree

    def __treeview2(self):
        # Frame para o segundo Treeview
        self.frame2 = ctk.CTkFrame(self, width=280, height=400)  # Defina o tamanho aqui
        self.frame2.place(x=405, y=10)  # Ajuste conforme necessário

        # Criação do Treeview para exibir contas com tema escuro
        label = ctk.CTkLabel(self.frame2, text='CONTAS NO COMBO')
        label.place(x=80, y=10)

        self.tree2 = ttk.Treeview(self.frame2, columns=("Nickname",), show="headings", style="Dark.Treeview")
        self.tree2.heading("Nickname", text="Nickname")
        
        # Centraliza os dados na coluna
        self.tree2.column("Nickname", anchor="center")  # Centraliza a coluna "Nickname"
        
        # Adicionando um scrollbar
        self.scrollbar2 = ttk.Scrollbar(self.frame2, orient="vertical", command=self.tree2.yview)
        self.tree2.configure(yscrollcommand=self.scrollbar2.set)
        
        self.tree2.place(x=15, y=40, width=240, height=300)
        self.scrollbar2.place(x=250, y=40, height=300)  # Posiciona o scrollbar ao lado da tree

    def __upButton(self):
        selected_item = self.tree2.selection()  # Obtém o item selecionado na segunda Treeview
        if selected_item:
            selected_nickname = self.tree2.item(selected_item, 'values')[0]  # Obtém o nickname
            index = self.list_second_tree.index(selected_nickname)  # Encontra o índice do nickname na lista
            if index > 0:  # Verifica se não é o primeiro item
                # Troca os itens na lista
                self.list_second_tree[index], self.list_second_tree[index - 1] = self.list_second_tree[index - 1], self.list_second_tree[index]
                self.atualizar_treeview2()  # Atualiza a segunda Treeview
                self.tree2.selection_set(self.tree2.get_children()[index - 1])  # Mantém a célula selecionada

    def __downButton(self):
        selected_item = self.tree2.selection()  # Obtém o item selecionado na segunda Treeview
        if selected_item:
            selected_nickname = self.tree2.item(selected_item, 'values')[0]  # Obtém o nickname
            index = self.list_second_tree.index(selected_nickname)  # Encontra o índice do nickname na lista
            if index < len(self.list_second_tree) - 1:  # Verifica se não é o último item
                # Troca os itens na lista
                self.list_second_tree[index], self.list_second_tree[index + 1] = self.list_second_tree[index + 1], self.list_second_tree[index]
                self.atualizar_treeview2()  # Atualiza a segunda Treeview
                self.tree2.selection_set(self.tree2.get_children()[index + 1])  # Mantém a célula selecionada

    def __leftButton(self):
        selected_item = self.tree2.selection()  # Obtém o item selecionado na segunda Treeview
        if selected_item:
            selected_nickname = self.tree2.item(selected_item, 'values')[0]  # Obtém o nickname
            self.tree2.delete(selected_item)  # Remove da segunda Treeview
            self.list_second_tree.remove(selected_nickname)  # Remove da lista

    def __rightButton(self):
        selected_item = self.tree.selection()  # Obtém o item selecionado na primeira Treeview
        if selected_item:
            selected_nickname = self.tree.item(selected_item, 'values')[0]  # Obtém o nickname
            if selected_nickname not in self.list_second_tree:  # Verifica se já não está na lista
                self.list_second_tree.append(selected_nickname)  # Adiciona à lista
                self.atualizar_treeview2()  # Atualiza a segunda Treeview

    def atualizar_treeview2(self):
        # Limpa a Treeview2 antes de atualizar
        self.tree2.delete(*self.tree2.get_children())
        # Adiciona os nicknames da lista à Treeview2
        for nickname in self.list_second_tree:
            self.tree2.insert("", "end", values=(nickname,))   

    def atualizar_treeview(self):
        # Obtém o estado atual das contas
        current_nicknames = [entry[0] for entry in current_state]

        # Verifica se o estado atual é diferente do estado anterior
        if current_nicknames != self.previous_state:
            # Limpa o Treeview antes de atualizar
            self.tree.delete(*self.tree.get_children())
            
            # Adiciona os logins ativos do current_state ao Treeview
            for entry in current_state:
                nickname = entry[0]  # Obtém o nickname
                self.tree.insert("", "end", values=(nickname,))  # Insere o nickname no Treeview
            
            # Atualiza o estado anterior
            self.previous_state = current_nicknames

    def __buttonOrder(self):
        self.vieworder.deiconify()
        self.vieworder.key_listener()
        self.withdraw()

    def __validate_numeric_input(self,value_if_allowed):
        if value_if_allowed.isdigit() or value_if_allowed == "":
            return True
        return False

    def __frameIntervalo1(self):
        vcmd = (self.register(self.__validate_numeric_input), '%P')
        frame = ctk.CTkFrame(self, width=200, height=200)
        frame.place(x=10, y=365)  # Ajuste a posição conforme necessário

        label = ctk.CTkLabel(frame, text="Intervalo:")
        label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.combobox_intervalo1 = ctk.CTkComboBox(frame, values=["F1 até F4", "1 ao 5"],width=90,state='readonly')
        self.combobox_intervalo1.set("F1 até F4")
        self.combobox_intervalo1.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        self.entry_intervalo1 = ctk.CTkEntry(frame,validate="key",validatecommand=vcmd,width=130)
        self.entry_intervalo1.insert(0, "1000")  # Define o valor inicial padrão
        self.entry_intervalo1.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        label_ms = ctk.CTkLabel(frame, text="(ms)")
        label_ms.place(x=150,y=57)

    def __frameIntervalo2(self):
        vcmd = (self.register(self.__validate_numeric_input), '%P')
        frame = ctk.CTkFrame(self, width=200, height=200)
        frame.place(x=190, y=365)  # Ajuste a posição conforme necessário

        label = ctk.CTkLabel(frame, text="Intervalo:")
        label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.combobox_intervalo2 = ctk.CTkComboBox(frame, values=["F5 até F8", "6 ao 9"],width=90,state='readonly')
        self.combobox_intervalo2.set("F5 até F8")
        self.combobox_intervalo2.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        self.entry_intervalo2 = ctk.CTkEntry(frame,validate="key",validatecommand=vcmd,width=130)
        self.entry_intervalo2.insert(0, "1000")  # Define o valor inicial padrão
        self.entry_intervalo2.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        label_ms = ctk.CTkLabel(frame, text="(ms)")
        label_ms.place(x=150,y=57)

        separator = ttk.Separator(self, orient='vertical')
        separator.place(x=192, y=370, height=85)  # Ajuste a posição e altura conforme necessário

        frame = ctk.CTkFrame(self, width=200, height=200)
        frame.place(x=100, y=465)  # Ajuste a posição conforme necessário

        # Adiciona um label para imagem ao lado esquerdo
        image_label = ctk.CTkLabel(self, image=botao_de_informacao,text='')  # Substitua 'some_image_variable' pela sua variável de imagem
        image_label.place(x=50,y=472)

        label = ctk.CTkLabel(frame, text="Por padrão é 1000ms\nDependendo do PC. 100ms > \nSó testando para saber 🤣")
        label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    def __teclaunica(self):
        frame = ctk.CTkFrame(self, width=200, height=100)
        frame.place(x=405, y=415)  # Ajuste a posição conforme necessário

        # Checkbox para ativar
        self.activate_checkbox_loop = ctk.CTkCheckBox(frame, text="Ativar Loop",width=19,command=self.toggle_intervals)
        self.activate_checkbox_loop.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Label para "Tecla:"
        tecla_label = ctk.CTkLabel(frame, text="Tecla:")
        tecla_label.grid(row=1, column=0, padx=10, pady=10, sticky="we")

        # Combobox para seleção de teclas
        self.tecla_combobox_loop = ctk.CTkComboBox(frame, width=60, values=["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "1", "2", "3", "4", "5", "6", "7", "8", "9"],state='disabled')
        self.tecla_combobox_loop.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # Label para "Tecla:"
        tecla_label = ctk.CTkLabel(frame, text="Quando ativado repete\napenas essa tecla\nem todas as contas.")
        tecla_label.grid(row=0, column=2, rowspan=2, sticky="we",padx=9)
  
    def toggle_intervals(self):
        # Configurar o estado do tecla_combobox_loop
        state = "readonly" if self.activate_checkbox_loop.get() else "disabled"
        self.tecla_combobox_loop.configure(state=state)
        
        # Desativar ou ativar elementos do frameIntervalo1
        self.combobox_intervalo1.configure(state="disabled" if self.activate_checkbox_loop.get() else "normal")
        self.entry_intervalo1.configure(state="disabled" if self.activate_checkbox_loop.get() else "normal")
        # Desativar ou ativar elementos do frameIntervalo2
        self.combobox_intervalo2.configure(state="disabled" if self.activate_checkbox_loop.get() else "normal")
        self.entry_intervalo2.configure(state="disabled" if self.activate_checkbox_loop.get() else "normal")

class viewOrder(ctk.CTkToplevel):
    def __init__(self, master=None, list_second_tree = [],combobox_interval1=None, time_interval1=None, combobox_interval2=None, time_interval2=None, checkbox_single=None, combobox_single=None):
        super().__init__(master)
        self.list_second_tree = list_second_tree
        self.combobox_interval1 = combobox_interval1
        self.time_interval1 = time_interval1
        self.combobox_interval2 = combobox_interval2
        self.time_interval2 = time_interval2
        self.checkbox_single = checkbox_single
        self.combobox_single = combobox_single
        self.__windowcfg()
        self.__treeview()
        self.__elements()

        # Carregar as teclas do arquivo JSON
        self.para_baixo, self.para_cima = self.load_bind_keys()

        self.key_listener()
        self.previous_state = []

        # Instancia o gerenciador de Treeview
        self.treeview_order = TreeviewManager(self)
        self.treeview_order.start()

        self.thread_flag_singleCombo = True
    
    def key_listener(self):
        self.setup_key_listener()

    def load_bind_keys(self):
        try:
            with open('binds.json', 'r') as json_file:
                settings_data = json.load(json_file)
                para_baixo = settings_data.get("para_baixo", "'")
                para_cima = settings_data.get("para_cima", "+")
        except (FileNotFoundError, json.JSONDecodeError):
            para_baixo = "'"
            para_cima = None
        
        return para_baixo, para_cima

    def atualizar_treeview(self):
        # Verifica se o estado atual de list_second_tree é diferente do estado anterior
        if self.list_second_tree != self.previous_state:

            # Limpa o Treeview antes de atualizar
            self.tree.delete(*self.tree.get_children())
            
            # Adiciona os logins ativos de list_second_tree ao Treeview
            for nickname in self.list_second_tree:
                self.tree.insert("", "end", values=(nickname,))  # Insere o nickname no Treeview

            # Seleciona a primeira célula, se houver elementos
            if self.list_second_tree:
                first_item = self.tree.get_children()[0]
                self.tree.selection_set(first_item)

            # Atualiza o estado anterior com uma cópia de list_second_tree
            self.previous_state = self.list_second_tree.copy()
        else:
            pass

    def close_all(self):
        self.thread_flag_singleCombo = False
        keyboard.unhook_all()
        self.master.deiconify()
        self.withdraw()

    def __windowcfg(self):
        self.protocol("WM_DELETE_WINDOW",lambda: self.close_all())
        self.title("Combar")
        self.resizable(False, False)
        self.attributes('-topmost', True)

        self.window_width = 250
        self.window_height = 500
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - self.window_width) // 2
        y = (screen_height - self.window_height) // 2
        self.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))
    
    def __treeview(self):
        # Criação de um frame para o Treeview
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Criação do Treeview para exibir contas com tema escuro
        self.tree = ttk.Treeview(frame, columns=("Nickname",), show="headings", style="Dark.Treeview")
        self.tree.heading("Nickname", text="Nickname",anchor='center')
        
        # Centraliza os dados na coluna
        self.tree.column("Nickname", anchor="center")  # Centraliza a coluna "Nickname"
        
        self.tree.place(x=10, y=10, width=210, height=395)
        
        # Adiciona um bind para o evento de duplo clique
        self.tree.bind("<Button-1>", self.on_click)

    def on_click(self, event):
        # Obtém a posição do clique
        region = self.tree.identify_region(event.x, event.y)

        # Verifica se o clique foi em uma linha (item) válida
        if region == "cell":  # Verifica se o clique foi em uma célula
            item_id = self.tree.identify_row(event.y)
            if item_id:
                nickname = self.tree.item(item_id, 'values')[0]

                # Procura o hwnd correspondente ao nickname e ativa a janela
                for entry in current_state:
                    if entry[0] == nickname:  # Supondo que entry[0] é o nickname
                        hwnd = entry[2]  # Supondo que entry[2] é o hwnd
                        ativar(hwnd)
                        break
        else:
            print("Clique em uma área não válida.")  # Mensagem opcional para debug

    def __elements(self):
        # ... código existente ...
        
        # Criação de um frame para organizar os botões horizontalmente
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=10,pady=10)

        # Botão para voltar
        back_button = ctk.CTkButton(button_frame, width=20, image=back, text="", command=lambda: self.close_all())
        back_button.pack(side="left", padx=10,pady=10)

        # Botão para combar
        combar_button = ctk.CTkButton(button_frame, text="Combar", command=self.__combar_action)
        combar_button.pack(side="right", padx=10,pady=10)

    def __loopSingleCombo(self):
        # Verifica se há itens na tabela antes de iniciar o loop
        if not self.tree.get_children():
            self.thread_flag_singleCombo = False
            print('Vazio')
            return  # Sai da função se a tabela estiver vazia

        items = self.tree.get_children()  # Captura todos os itens
        item_count = len(items)

        # Variável para controlar se o TAB já foi enviado
        tab_sent = False

        while self.thread_flag_singleCombo:
            for index in range(item_count):
                item = items[index]
                nickname = self.tree.item(item, 'values')[0]

                # Seleciona o item atual visivelmente na Treeview
                self.tree.selection_set(item)
                self.tree.see(item)  # Faz o item visível na árvore
                
                for entry in current_state:
                    if entry[0] == nickname:  # Supondo que entry[0] é o nickname
                        hwnd = entry[2]  # Supondo que entry[2] é o hwnd
                        ativar(hwnd)
                        print(hwnd)
                        time.sleep(0.001)  # Adiciona um delay de 1 segundo
                        tecla = self.combobox_single.get()  # Obtém a tecla selecionada
                        # Envia a tecla 'TAB' apenas na primeira iteração
                        if not tab_sent:
                            enviar_tecla(hwnd, 'TAB')
                            tab_sent = True  # Atualiza o controle para que o TAB não seja enviado novamente
                        enviar_tecla(hwnd, tecla)  # Envia a tecla
                        break
                
                # Aguarda um momento para que o usuário veja a seleção antes de passar para a próxima
                time.sleep(1)  # Tempo para visualizar a seleção

            # Reinicia o índice quando chegar ao final
            print("Reiniciando o loop.")
            time.sleep(1)  # Adiciona um delay antes de reiniciar o loop

        print("Loop encerrado.")

    def __send_keys_to_accounts(self):
        # Verifica se há itens na tabela antes de iniciar a ação
        if not self.tree.get_children():
            print('Vazio, não há contas para enviar teclas.')
            return  # Sai da função se a tabela estiver vazia

        items = self.tree.get_children()  # Captura todos os itens
        item_count = len(items)

        # Verifica se result_list está definida
        if not hasattr(self, 'result_list') or not self.result_list:
            print('result_list está vazia ou não definida.')
            return

        # Obtendo os delays definidos
        delay_interval1 = self.time_interval1.get()  # Obtém o valor do delay do ComboBox
        delay_interval2 = self.time_interval2.get()  # Obtém o valor do delay do ComboBox

        # Dicionário para controlar se o TAB já foi enviado para cada nickname
        tab_sent_status = {self.tree.item(item, 'values')[0]: False for item in items}

        # Itera sobre as teclas na result_list
        for key in self.result_list:
            # Para cada tecla, envia para todos os itens
            for index in range(item_count):
                item = items[index]
                nickname = self.tree.item(item, 'values')[0]

                # Seleciona o item atual visivelmente na Treeview
                self.tree.selection_set(item)
                self.tree.see(item)  # Faz o item visível na árvore

                for entry in current_state:
                    if entry[0] == nickname:  # Supondo que entry[0] é o nickname
                        hwnd = entry[2]  # Supondo que entry[2] é o hwnd
                        ativar(hwnd)
                        print(hwnd)
                        time.sleep(0.001)  # Adiciona um delay para garantir que a janela esteja ativa

                        # Envia o TAB se ainda não foi enviado para este nickname
                        if not tab_sent_status[nickname]:
                            enviar_tecla(hwnd, 'TAB')
                            tab_sent_status[nickname] = True  # Atualiza o controle para que o TAB não seja enviado novamente para esse nickname
                        
                        # Envia a tecla atual para o item
                        enviar_tecla(hwnd, key)  # Envia a tecla
                        break  # Para evitar enviar para o mesmo item várias vezes

                # Aqui você pode decidir qual delay aplicar entre as teclas
                if delay_interval1.isdigit() and delay_interval2.isdigit():
                    delay1 = int(delay_interval1) / 1000
                    delay2 = int(delay_interval2) / 1000
                    time.sleep(delay1 if key in ['F1', 'F2', 'F3', 'F4','1','2','3','4','5'] else delay2)
                else:
                    print('Os intervalos não são válidos. Usando delay padrão de 1 segundo.')
                    time.sleep(1)  # Usar um delay padrão se os valores não forem válidos

            # Aguarda um momento antes de passar para a próxima tecla
            time.sleep(1)  # Tempo para visualizar a seleção

        print("Envio de teclas encerrado.")
        
    def __combar_action(self):
        if self.checkbox_single.get() == 1:
            self.thread_flag_singleCombo = True
            thread = threading.Thread(target=self.__loopSingleCombo)
            thread.daemon = True
            thread.start()
        else:
            # Obtendo os valores dos ComboBox
            interval1 = self.combobox_interval1.get()
            interval2 = self.combobox_interval2.get()

            result_list = []  # Lista para armazenar os valores
            
            # Lógica para o interval1
            if interval1 == "F1 até F4":
                result_list.extend(['F1', 'F2', 'F3', 'F4'])
            elif interval1 == "1 ao 5":
                result_list.extend(['1', '2', '3', '4', '5'])
            
            # Lógica para o interval2
            if interval2 == "F5 até F8":
                result_list.extend(['F5', 'F6', 'F7', 'F8'])
            elif interval2 == "6 ao 9":
                result_list.extend(['6', '7', '8', '9'])
            
            # Armazenando a result_list como um atributo da classe
            self.result_list = result_list
            
            # Imprimindo a lista para testar
            print(self.result_list)

            # Criando uma thread para enviar as teclas
            thread = threading.Thread(target=self.__send_keys_to_accounts)
            thread.daemon = True  # Torna a thread um daemon
            thread.start()  # Inicia a thread

    def setup_key_listener(self):
        keyboard.on_release(self.on_key_up_event)

    def select_next_cell(self):
        selected_item = self.tree.selection()
        if selected_item:
            current_index = self.tree.index(selected_item)
            next_index = (current_index + 1) % len(self.tree.get_children())
            next_item = self.tree.get_children()[next_index]
            self.tree.selection_set(next_item)
            return self.tree.item(next_item, 'values')[0]  # Retorna o nickname

    def select_previous_cell(self):
        selected_item = self.tree.selection()
        if selected_item:
            current_index = self.tree.index(selected_item)
            previous_index = (current_index - 1) % len(self.tree.get_children())
            previous_item = self.tree.get_children()[previous_index]
            self.tree.selection_set(previous_item)
            return self.tree.item(previous_item, 'values')[0]  # Retorna o nickname

    def on_key_up_event(self, event):
        nickname = None
        if event.name.upper() == self.para_baixo:
            nickname = self.select_next_cell()
        elif event.name.upper() == self.para_cima:
            nickname = self.select_previous_cell()

        # Verifica o current_state para encontrar o hwnd correspondente ao nickname
        if nickname is not None:
            for entry in current_state:
                if entry[0] == nickname:  # Supondo que entry[0] é o nickname
                    hwnd = entry[2]  # Supondo que entry[1] é o hwnd
                    ativar(hwnd)
                    break

class BindRoot(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.__windowcfg()
        self.__elements()
        self.load_settings()  # Carrega as configurações ao inicializar

        self.grab_set()
        self.focus()

    def __windowcfg(self):
        self.title("Binds")  # Título alterado para "Binds"
        self.resizable(False, False)

        self.window_width = 280
        self.window_height = 145
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - self.window_width) // 2
        y = (screen_height - self.window_height) // 2
        self.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __elements(self):
        # Label e Combobox para "Tecla Próxima Janela"
        next_window_label = ctk.CTkLabel(self, text="Tecla para Baixo:")
        next_window_label.grid(row=0, column=0, padx=10, pady=10)

        self.next_window_combobox = ctk.CTkComboBox(self, values=["'", "-", "DOWN"], state="readonly")  # Definido como não editável
        self.next_window_combobox.grid(row=0, column=1, padx=10, pady=10)

        # Label e Combobox para "Tecla Janela Anterior"
        previous_window_label = ctk.CTkLabel(self, text="Tecla para Cima:")
        previous_window_label.grid(row=1, column=0, padx=10, pady=10)

        self.previous_window_combobox = ctk.CTkComboBox(self, values=["+", "UP"], state="readonly")  # Definido como não editável
        self.previous_window_combobox.grid(row=1, column=1, padx=10, pady=10)

        # Botão "Salvar"
        save_button = ctk.CTkButton(self, text="Salvar", command=self.save_settings)
        save_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # ... código existente ...

    def load_settings(self):
        # Tenta carregar as configurações do arquivo JSON
        try:
            with open('binds.json', 'r') as json_file:
                settings_data = json.load(json_file)
                # Define os valores das comboboxes com base nos dados carregados
                self.next_window_combobox.set(settings_data.get("para_baixo", "'"))  # Valor padrão se não encontrado
                self.previous_window_combobox.set(settings_data.get("para_cima", "+"))  # Valor padrão se não encontrado
        except (FileNotFoundError, json.JSONDecodeError):
            # Se o arquivo não existir ou estiver corrompido, define valores padrão
            self.next_window_combobox.set("'")
            self.previous_window_combobox.set("+")

    def save_settings(self):
        # Obtém as teclas selecionadas
        next_key = self.next_window_combobox.get()
        previous_key = self.previous_window_combobox.get()

        # Dados a serem salvos no JSON
        settings_data = {
            "para_baixo": next_key,
            "para_cima": previous_key
        }

        # Salva os dados no arquivo JSON
        try:
            with open('binds.json', 'w') as json_file:
                json.dump(settings_data, json_file, indent=4)
            print("Configurações salvas com sucesso.")
        except IOError:
            print("Erro ao salvar as configurações.")