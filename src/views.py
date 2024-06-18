import customtkinter as ctk
from CTkMenuBar import *
from CTkTable import *
from CTkMessagebox import *
from tkinter import filedialog, messagebox, StringVar
from PIL import Image
import json
import os
from src.shortcutscontroller import criar_atalho,editar_atalho,excluir_atalho,excluir_todos_atalhos
from src.execs import exec_element,close_all_pws

browse_image = ctk.CTkImage(Image.open("./res/search.png"), size=(20, 20))
visible_on = ctk.CTkImage(Image.open("./res/visibility_icon.png"), size=(20, 20))
visible_off = ctk.CTkImage(Image.open("./res/off_visibility_icon.png"), size=(20, 20))
backgrond_image = ctk.CTkImage(Image.open("./res/background.png"),size=(350,350))

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

        window_width = 463
        window_height = 100
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.after(200, lambda: self.wm_iconbitmap('./res/icon.ico'))
        self.after(200, lambda: self.iconbitmap('./res/icon.ico'))

    def __elements(self):
        self.label = ctk.CTkLabel(self, text="Executável:")
        self.label.grid(row=0, column=0, padx=10, pady=10)

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
            
            CTkMessagebox(title="Sucesso",message="O caminho do executável foi salvo com sucesso!",
                  icon="check", option_1="Ok")

            self.open_Manager()
        else:
            CTkMessagebox(title="Erro", message="Nenhum caminho de executável selecionado.", icon="cancel")

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
                exec_element(row_current_data[0])

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
            print('x')
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

        CTkMessagebox(title="EM CONSTRUÇÃO!!", message="EM CONSTRUÇÃO!!", icon="cancel")

class EditLogin(ctk.CTkToplevel):
    def __init__(self, master, login='None',password='None',nickname='None',icon_path='None',row_index=0):
        super().__init__(master)
        self.master = master
        self.login = login
        self.password = password
        self.nickname = nickname
        self.icon_path = icon_path
        self.row_index = row_index

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

        image = ctk.CTkImage(Image.open(self.icon_path), size=(32, 32))
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
        self.currentPath = None
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

        accounts[self.row_index - 1] = {
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

        CTkMessagebox(title='Conta editada',message=f"Conta editada !!",
                  icon="check", option_1="Ok")

        self.master.updateTableEdit(self.row_index,login,nickname)

        editar_atalho(self.login,login,password,nickname,icon_path)

        self.destroy()