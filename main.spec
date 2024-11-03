# -*- mode: python -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # O seu script principal
    pathex=['.'],  # Caminho onde o script está localizado
    binaries=[],
    datas=[
        ('./res/icon.ico', '.'),  # Incluindo o ícone no executável
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)  # Use a.zipped_data em vez de a.zipped

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Login Manager',  # Nome do programa
    debug=False,  # Mantenha False em produção
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Não exibir console
    icon='./res/icon.ico',  # O caminho do ícone a ser usado
    onefile=True,  # Adiciona a opção onefile
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Login Manager',  # Nome do programa
)
