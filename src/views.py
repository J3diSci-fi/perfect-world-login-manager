from tkinter import StringVar
from tkinter import ttk
import customtkinter as ctk
from CTkMenuBar import *
from CTkTable import *
from CTkMessagebox import *
from tkinter import filedialog, messagebox, StringVar
from PIL import Image
import json
import os
from src.shortcutscontroller import criar_atalho,editar_atalho,excluir_atalho,excluir_todos_atalhos
from src.execs import close_all_pws,add_to_queue
import threading
import time
import keyboard
from src.focus import ativar
from src.actions import enviar_tecla_shift_1, enviar_tecla

browse_image = ctk.CTkImage(Image.open("./res/search.png"), size=(20, 20))
visible_on = ctk.CTkImage(Image.open("./res/visibility_icon.png"), size=(20, 20))
visible_off = ctk.CTkImage(Image.open("./res/off_visibility_icon.png"), size=(20, 20))
backgrond_image = ctk.CTkImage(Image.open("./res/background.png"),size=(350,350))
seta_cima = ctk.CTkImage(Image.open("./res/seta-para-cima.png"),size=(16,16))
seta_baixo =ctk.CTkImage(Image.open("./res/seta-para-baixo.png"),size=(16,16))
confirm = ctk.CTkImage(Image.open("./res/confirm.png"),size=(16,16))
cancel = ctk.CTkImage(Image.open("./res/cancel.png"),size=(16,16))
back = ctk.CTkImage(Image.open("./res/back.png"),size=(24,24))

class Root(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.__windowcfg()
        self.__elements()
        self.check_existing_executable()

        self.mainloop()

    def __windowcfg(self):
        self.title("Selecionar Executável")
        self.resizable(False, False)

        window_width = 530
        window_height = 150
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

        label = ctk.CTkLabel(self, text="Título da Janela do PW:")
        label.grid(row=1, column=0, padx=10, pady=10)

        self.entry_window_title = ctk.CTkEntry(self, width=300, state="normal")
        self.entry_window_title.grid(row=1, column=1, padx=10, pady=10)

        self.confirm_button = ctk.CTkButton(self, text="Confirmar", command=self.confirm)
        self.confirm_button.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

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
        window_title = self.entry_window_title.get()

        if exe_path and window_title:
            data = {"path_executable": exe_path,
                    "path_folder":folder_path,
                    "window_title":window_title}
            with open("executable_path.json", "w") as json_file:
                json.dump(data, json_file, indent=4)
            
            self.open_Manager()
        else:
            CTkMessagebox(title="Erro", message="Preencha os campos corretamente.", icon="cancel")

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
        self.master = master

        self.__windowcfg()
        self.__elements()
        self.__frameImage()
        self.__frameAddAccount()
        self.__frameTable()
        self.__framebottom_table()

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
        menu.add_cascade("Combo",command=self.__comboWindow)
        menu.add_cascade("Mudar Caminho Executável",command=self.change_executable_path)
        menu.add_cascade("Resetar App",command=self.reset_app)

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
                add_to_queue(row_current_data[0],row_current_data[1])

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
        ComboRoot(self)

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

        print(icon_path)

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
        frame = ctk.CTkFrame(self)
        frame.place(x=10, y=10)

        login_var = StringVar(value=self.login)
        password_var = StringVar(value=self.password)
        nickname_var = StringVar(value=self.nickname)

        login_label = ctk.CTkLabel(frame, text="Login:")
        login_label.grid(row=0, column=0, padx=10, pady=10)

        self.login_entry = ctk.CTkEntry(frame,textvariable=login_var)
        self.login_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=10, pady=10)

        password_label = ctk.CTkLabel(frame, text="Senha:")
        password_label.grid(row=1, column=0, padx=10, pady=10)

        self.password_visible = False
        self.password_entry = ctk.CTkEntry(frame, show="*" ,textvariable=password_var)
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        self.toggle_button = ctk.CTkButton(frame, image=visible_on, text="", command=self.toggle_password_visibility, width=10)
        self.toggle_button.grid(row=1, column=2, padx=10, pady=10)

        nickname_label = ctk.CTkLabel(frame, text="Nickname:")
        nickname_label.grid(row=2, column=0, padx=10, pady=10)

        self.nickname_entry = ctk.CTkEntry(frame,textvariable=nickname_var)
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
        self.master = master

        self.__windowcfg()
        self.__elements()

        self.grab_set()
        self.focus()

        # Iniciar thread para atualizar a Treeview
        self.flagThread = True
        self.update_thread = threading.Thread(target=self.auto_update_treeview)
        self.update_thread.start()
        
        self.setup_key_listener()

        # Ensure treeview order functions are available
        self.save_treeview_order()
        self.restore_treeview_order([])

    def close_all(self):
        self.flagThread = False
        self.master.deiconify()
        keyboard.unhook_all()
        self.destroy()  # Fecha a janela principal (Root)

    def __windowcfg(self):
        self.title("Combozada")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close_all)

        window_width = 702
        window_height = 420
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __elements(self):
        frame = ctk.CTkFrame(self)
        frame.place(x=265, y=10)

        label_instruct = ctk.CTkLabel(frame, text="Em qual tecla está o ataque auxiliar na barra de skill das contas?")
        label_instruct.grid(row=0, column=0, padx=10, pady=10)

        # Frame 2
        frame2 = ctk.CTkFrame(self)
        frame2.place(x=230, y=70)

        label_atq_auxiliar = ctk.CTkLabel(frame2, text="Tecla Ataque Auxiliar\n(Pegar TG do Líder):")
        label_atq_auxiliar.grid(row=0, column=0, padx=10, pady=10)

        keys = self.get_keyboard_keys()
        self.combo_box_tecla_atq_auxiliar = ctk.CTkComboBox(frame2, state='readonly', values=keys)
        self.combo_box_tecla_atq_auxiliar.grid(row=0, column=1, padx=10, pady=10)
        
        self.button_cancel_atq_auxiliar = ctk.CTkButton(frame2, image=cancel, text="", command=self.cancel_aux_attack,width=10 , state = 'disabled', fg_color='gray')
        self.button_cancel_atq_auxiliar.grid(row=0, column=2, padx=10, pady=10)

        self.button_confirm_atq_auxiliar = ctk.CTkButton(frame2, image=confirm, text="", command=self.confirm_aux_attack,width=10)
        self.button_confirm_atq_auxiliar.grid(row=0, column=3, padx=10, pady=10)

        label_macro = ctk.CTkLabel(frame2, text="Sequência de Teclas:")
        label_macro.grid(row=1, column=0, padx=10, pady=10)

        self.combo_box_tecla_macro = ctk.CTkComboBox(frame2, state='readonly', values=['F1 ao F8','1 ao 9'],width=100)
        self.combo_box_tecla_macro.grid(row=1, column=1, padx=10, pady=10)

        # Variável associada à entrada
        self.macro_ms_var = StringVar()

        # Adicionar observador à variável
        self.macro_ms_var.trace_add("write", self.validate_numeric_input)

        self.button_cancel_macro = ctk.CTkButton(frame2, image=cancel, text="", command=self.cancel_macro,width=10, state = 'disabled', fg_color='gray')
        self.button_cancel_macro.grid(row=1, column=2, padx=10, pady=10)

        self.button_confirm_macro = ctk.CTkButton(frame2, image=confirm, text="", command=self.confirm_macro,width=10)
        self.button_confirm_macro.grid(row=1, column=3, padx=10, pady=10)

        label_ms = ctk.CTkLabel(frame2,text='(ms) Recomendado: 100-200 \n ou até menos')
        label_ms.grid(row=2,column=0, padx=10, pady=10)

        self.input_macro_ms = ctk.CTkEntry(frame2,width=100,textvariable=self.macro_ms_var)
        self.input_macro_ms.grid(row=2,column=1, padx=10, pady=10)

        self.button_cancel_macro_ms = ctk.CTkButton(frame2,image=cancel,text='', width=10, command=self.cancel_macro_ms,  state = 'disabled', fg_color='gray')
        self.button_cancel_macro_ms.grid(row=2,column=2, padx=10, pady=10)

        self.button_confirm_macro_ms = ctk.CTkButton(frame2,image=confirm,text='',width=10, command=self.confirm_macro_ms)
        self.button_confirm_macro_ms.grid(row=2,column=3, padx=10, pady=10)

        # Frame 3
        frame3 = ctk.CTkFrame(self)
        frame3.place(x=255, y=233)

        label_hotkey_tg = ctk.CTkLabel(frame3, text="Hotkey pegar TG líder:")
        label_hotkey_tg.grid(row=0, column=0, padx=10, pady=10)

        self.combo_box_hotkey_tg = ctk.CTkComboBox(frame3, state='readonly', values=keys)
        self.combo_box_hotkey_tg.grid(row=0, column=1, padx=10, pady=10)

        self.button_cancel_hotkey_tg = ctk.CTkButton(frame3, image=cancel, text="", command=self.cancel_hotkey_tg,width=10, state = 'disabled', fg_color='gray')
        self.button_cancel_hotkey_tg.grid(row=0, column=2, padx=10, pady=10)

        self.button_confirm_hotkey_tg = ctk.CTkButton(frame3, image=confirm, text="", command=self.confirm_hotkey_tg,width=10)
        self.button_confirm_hotkey_tg.grid(row=0, column=3, padx=10, pady=10)

        label_hotkey_combar = ctk.CTkLabel(frame3, text="Hotkey para Combar:")
        label_hotkey_combar.grid(row=1, column=0, padx=10, pady=10)

        self.combo_box_hotkey_combar = ctk.CTkComboBox(frame3, state='readonly', values=keys)
        self.combo_box_hotkey_combar.grid(row=1, column=1, padx=10, pady=10)

        self.button_cancel_hotkey_combar = ctk.CTkButton(frame3, image=cancel, text="", command=self.cancel_hotkey_combar,width=10, state = 'disabled', fg_color='gray')
        self.button_cancel_hotkey_combar.grid(row=1, column=2, padx=10, pady=10)

        self.button_confirm_hotkey_combar = ctk.CTkButton(frame3, image=confirm, text="", command=self.confirm_hotkey_combar,width=10)
        self.button_confirm_hotkey_combar.grid(row=1, column=3, padx=10, pady=10)

        label_obs = ctk.CTkLabel(frame3, text="Obs: Não utilize a mesma tecla para ambas as Hotkeys(bug).")
        label_obs.grid(row=2, column=0, columnspan=4, padx=10, pady=10)

        # Frame 4 - Treeview e Botões
        frame4 = ctk.CTkFrame(self)
        frame4.place(x=10, y=10)

        label5 = ctk.CTkLabel(frame4, text="(Irá combar na ordem)\nO Líder da PT deve ser o primeiro!")
        label5.pack(side='top', padx=10, pady=5)

        style = ttk.Style()
        style.theme_use("clam")  # Using a modern theme as a base
        style.configure("Treeview",
                        background="#2e2e2e",
                        foreground="white",
                        rowheight=25,
                        fieldbackground="#2e2e2e")
        style.map('Treeview', background=[('selected', '#4a4a4a')])

        style.configure("Treeview.Heading", background="#4a4a4a", foreground="white", font=("Arial", 10, "bold"))
        style.configure("Treeview.Cell", anchor="center")

        self.tree = ttk.Treeview(frame4, columns=('Personagens Logados',), show='headings', style="Treeview",height=15)
        self.tree.heading('Personagens Logados', text='Personagens Logados', anchor='center')
        self.tree.column('Personagens Logados', anchor='center')
        self.tree.pack(side='top')

        # Adicionar dados do arquivo JSON
        self.update_treeview_data()

        button_up = ctk.CTkButton(frame4, image=seta_cima, text="", command=self.move_up, width=10)
        button_up.pack(side='left', padx=10, pady=5, anchor='e', expand=True)

        button_down = ctk.CTkButton(frame4, image=seta_baixo, text="", command=self.move_down, width=10)
        button_down.pack(side='right', padx=10, pady=5, anchor='w', expand=True)

        button_back_to_manager = ctk.CTkButton(self, image=back, text="", command=self.close_all, width=10)
        button_back_to_manager.place(x=655,y=383)

    def get_treeview_data(self):
        """Helper function to get current treeview data as a list of tuples"""
        return [self.tree.item(item)['values'][0] for item in self.tree.get_children()]
    
    def validate_numeric_input(self, *args):
        value = self.macro_ms_var.get()
        if not value.isdigit():
            self.macro_ms_var.set(''.join(filter(str.isdigit, value)))

    def update_treeview_data(self):
        current_data = self.get_treeview_data()
        new_data = []

        try:
            with open('window_info.json', 'r') as file:
                data = json.load(file)
                for key, value in data.items():
                    if value.get('status') == 'on':
                        new_data.append(key)
        except (FileNotFoundError, json.JSONDecodeError):
            new_data = ['']

        if current_data != new_data:
            selected_item = self.tree.selection()
            selected_value = self.tree.item(selected_item)['values'][0] if selected_item else None
            
            # Save the current order of the treeview
            current_order = self.save_treeview_order()

            self.tree.delete(*self.tree.get_children())
            for item in new_data:
                self.tree.insert('', 'end', values=(item,))
            
            # Restore the treeview order
            self.restore_treeview_order(current_order)

            # Restore the selected item
            if selected_value:
                for item in self.tree.get_children():
                    if self.tree.item(item)['values'][0] == selected_value:
                        self.tree.selection_set(item)
                        break

    def auto_update_treeview(self):
        while self.flagThread:
            self.update_treeview_data()
            time.sleep(3)  # Atualiza a cada 3 segundos

    def save_treeview_order(self):
        """Saves the current order of items in the Treeview."""
        return [self.tree.item(item)['values'][0] for item in self.tree.get_children()]

    def restore_treeview_order(self, order):
        """Restores the order of items in the Treeview based on the provided order."""
        current_items = {self.tree.item(item)['values'][0]: item for item in self.tree.get_children()}
        for value in order:
            if value in current_items:
                self.tree.move(current_items[value], '', 'end')

    def get_keyboard_keys(self):
        return [str(i) for i in range(1, 10)] + [f"F{i}" for i in range(1, 9)]

    def cancel_aux_attack(self):
        self.button_cancel_atq_auxiliar.configure(state='disabled')
        self.button_cancel_atq_auxiliar.configure(fg_color='gray')
        self.button_confirm_atq_auxiliar.configure(state='normal')
        self.button_confirm_atq_auxiliar.configure(fg_color='#1F6AA5')
        self.combo_box_tecla_atq_auxiliar.configure(state='normal')
    
    def confirm_aux_attack(self):
        if self.combo_box_tecla_atq_auxiliar.get() == '':
            CTkMessagebox(title="Info", message="Selecione a tecla do comando auxiliar na barra de skill.")
            return
        
        self.button_cancel_atq_auxiliar.configure(state='normal')
        self.button_cancel_atq_auxiliar.configure(fg_color='#1F6AA5')
        self.button_confirm_atq_auxiliar.configure(state='disabled')
        self.button_confirm_atq_auxiliar.configure(fg_color='gray')
        self.combo_box_tecla_atq_auxiliar.configure(state='disabled')
    
    def cancel_macro(self):
        self.button_cancel_macro.configure(state='disabled')
        self.button_cancel_macro.configure(fg_color='gray')
        self.button_confirm_macro.configure(state='normal')
        self.button_confirm_macro.configure(fg_color='#1F6AA5')
        self.combo_box_tecla_macro.configure(state='normal')

    def confirm_macro(self):
        if self.combo_box_tecla_macro.get() == '':
            CTkMessagebox(title="Info", message="Selecione a tecla do combo na barra de skill.")
            return
        
        self.button_cancel_macro.configure(state='normal')
        self.button_cancel_macro.configure(fg_color='#1F6AA5')
        self.button_confirm_macro.configure(state='disabled')
        self.button_confirm_macro.configure(fg_color='gray')
        self.combo_box_tecla_macro.configure(state='disabled')

    def cancel_macro_ms(self):
        self.button_cancel_macro_ms.configure(state='disabled')
        self.button_cancel_macro_ms.configure(fg_color='gray')
        self.button_confirm_macro_ms.configure(state='normal')
        self.button_confirm_macro_ms.configure(fg_color='#1F6AA5')
        self.input_macro_ms.configure(state='normal')

    def confirm_macro_ms(self):
        if self.input_macro_ms.get() == '':
            CTkMessagebox(title="Info", message="Preencha o campo ms.")
            return
        self.button_cancel_macro_ms.configure(state='normal')
        self.button_cancel_macro_ms.configure(fg_color='#1F6AA5')
        self.button_confirm_macro_ms.configure(state='disabled')
        self.button_confirm_macro_ms.configure(fg_color='gray')
        self.input_macro_ms.configure(state='disabled')
    
    def cancel_hotkey_tg(self):
        self.button_cancel_hotkey_tg.configure(state='disabled')
        self.button_cancel_hotkey_tg.configure(fg_color='gray')
        self.button_confirm_hotkey_tg.configure(state='normal')
        self.button_confirm_hotkey_tg.configure(fg_color='#1F6AA5')
        self.combo_box_hotkey_tg.configure(state='normal')

    def confirm_hotkey_tg(self):
        if self.combo_box_hotkey_tg.get() == '':
            CTkMessagebox(title="Info", message="Selecione a hotkey para pegar TG líder com as contas.")
            return
        
        self.button_cancel_hotkey_tg.configure(state='normal')
        self.button_cancel_hotkey_tg.configure(fg_color='#1F6AA5')
        self.button_confirm_hotkey_tg.configure(state='disabled')
        self.button_confirm_hotkey_tg.configure(fg_color='gray')
        self.combo_box_hotkey_tg.configure(state='disabled')
    
    def cancel_hotkey_combar(self):
        self.button_cancel_hotkey_combar.configure(state='disabled')
        self.button_cancel_hotkey_combar.configure(fg_color='gray')
        self.button_confirm_hotkey_combar.configure(state='normal')
        self.button_confirm_hotkey_combar.configure(fg_color='#1F6AA5')
        self.combo_box_hotkey_combar.configure(state='normal')

    def confirm_hotkey_combar(self):
        if self.combo_box_hotkey_combar.get() == '':
            CTkMessagebox(title="Info", message="Selecione a hotkey para combar com as contas.")
            return
        
        self.button_cancel_hotkey_combar.configure(state='normal')
        self.button_cancel_hotkey_combar.configure(fg_color='#1F6AA5')
        self.button_confirm_hotkey_combar.configure(state='disabled')
        self.button_confirm_hotkey_combar.configure(fg_color='gray')
        self.combo_box_hotkey_combar.configure(state='disabled')

    def move_up(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            if index > 0:
                self.tree.move(selected_item, self.tree.parent(selected_item), index - 1)

    def move_down(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            if index < len(self.tree.get_children()) - 1:
                self.tree.move(selected_item, self.tree.parent(selected_item), index + 1)

    def expand_all(self):
        for item in self.tree.get_children():
            self.tree.item(item, open=True)

    def setup_key_listener(self):
        keyboard.on_release(self.on_key_up_event)

    def on_key_up_event(self, event):
        key_atq_auxiliar = self.combo_box_tecla_atq_auxiliar.get()
        keys_action = self.combo_box_tecla_macro.get()
        ms = self.input_macro_ms.get()
        
        hotkey_tg_lider = self.combo_box_hotkey_tg.get()
        hotkey_combar = self.combo_box_hotkey_combar.get()
        
        if event.name == hotkey_tg_lider:
            print(f"Tecla {hotkey_tg_lider} solta")
            if (self.button_confirm_atq_auxiliar.cget('state') == 'normal') or \
           (self.button_confirm_hotkey_combar.cget('state') == 'normal') or \
           (self.button_confirm_hotkey_tg.cget('state') == 'normal') or \
           (self.button_confirm_macro.cget('state') == 'normal') or \
           (self.button_confirm_macro_ms.cget('state') == 'normal'):


                CTkMessagebox(title="Info", message="Confirme todos os campos.")
                return
            else:
                dataTable = self.print_treeview_values()

                try:
                    with open('window_info.json', 'r') as file:
                        data = json.load(file)
                except (FileNotFoundError, json.JSONDecodeError):
                    print('Error')

                for index, key in enumerate(dataTable):
                    if index != 0:
                        hwnd = data[key]['hwnd']
                        ativar(hwnd)
                        time.sleep(int(ms)/1000)
                        enviar_tecla_shift_1(hwnd)
                        enviar_tecla(hwnd,key_atq_auxiliar)  
        
        if event.name == hotkey_combar:
            print(f"Tecla {hotkey_combar} solta")
            if (self.button_confirm_atq_auxiliar.cget('state') == 'normal') or \
           (self.button_confirm_hotkey_combar.cget('state') == 'normal') or \
           (self.button_confirm_hotkey_tg.cget('state') == 'normal') or \
           (self.button_confirm_macro.cget('state') == 'normal') or \
           (self.button_confirm_macro_ms.cget('state') == 'normal'):
            
                CTkMessagebox(title="Info", message="Confirme todos os campos.")
                return
            else:
                if keys_action == "F1 ao F8":
                    key_list = [f"F{i}" for i in range(1, 9)]
                    dataTable = self.print_treeview_values()

                    try:
                        with open('window_info.json', 'r') as file:
                            data = json.load(file)
                    except (FileNotFoundError, json.JSONDecodeError):
                        print('Error')
                              
                    for tecla in key_list:
                        for index, key in enumerate(dataTable):
                            print(tecla)
                            hwnd = data[key]['hwnd']
                            ativar(hwnd)
                            time.sleep(int(ms)/1000)
                            enviar_tecla(hwnd,tecla)  

    def print_treeview_values(self):
        data = self.get_treeview_data()
        return data