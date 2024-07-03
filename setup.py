import sys
from cx_Freeze import setup, Executable

# Opções de build do cx_Freeze
build_exe_options = {
    "packages": [],
    "includes": [],
    "excludes": [],
    "include_files": [('./res/icon.ico', 'icon.ico')]  # Inclui o ícone do aplicativo
}

# Configuração do executável
exe = Executable(
    script='main.py',  # Substitua 'seu_script.py' pelo nome do seu script Python
    base='Win32GUI' if sys.platform == 'win32' else None,  # Use 'Win32GUI' para ocultar o console no Windows
    icon='./res/icon.ico'  # Caminho para o ícone
)

# Configuração do setup
setup(
    name='Login Manager',
    version='1.0',
    description='Login Manager',
    options={"build_exe": build_exe_options},
    executables=[exe]
)