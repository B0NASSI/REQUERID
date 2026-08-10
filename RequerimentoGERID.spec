# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

icone_datas = [
    ('GERID LOGO.ico', '.'),
    ('Logo-RS-completa-colorida.ico', '.'),
]

# ---------------------------------------------------------------------------
# REQUERID.exe — programa principal
# ---------------------------------------------------------------------------
app_datas = list(icone_datas) + [
    ('modelos/template.docx', 'modelos'),
    ('modelos/template_exigencia.docx', 'modelos'),
    ('dados/SOLICITAÇÕES GERID.xlsx', 'dados'),
    ('NOTAS DE ATUALIZAÇÃO', 'NOTAS DE ATUALIZAÇÃO'),
]
app_binaries = []
app_hiddenimports = []
for pacote in ('ttkbootstrap', 'tkinterdnd2'):
    tmp_ret = collect_all(pacote)
    app_datas += tmp_ret[0]; app_binaries += tmp_ret[1]; app_hiddenimports += tmp_ret[2]

a_app = Analysis(
    ['main.py'],
    pathex=[SPECPATH],
    binaries=app_binaries,
    datas=app_datas,
    hiddenimports=app_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy', 'pywinauto', 'adodbapi', 'isapi', 'pythonwin',
        'setuptools', 'pip', 'unittest', 'email', 'http', 'xmlrpc',
        'ftplib', 'multiprocessing',
    ],
    noarchive=False,
    optimize=0,
)
pyz_app = PYZ(a_app.pure)

exe_app = EXE(
    pyz_app,
    a_app.scripts,
    [],
    exclude_binaries=True,
    name='REQUERID',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['GERID LOGO.ico'],
)

coll_app = COLLECT(
    exe_app,
    a_app.binaries,
    a_app.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='REQUERID',
)

# ---------------------------------------------------------------------------
# REQUERID Launcher.exe — checa atualizações no GitHub antes de abrir o app
# ---------------------------------------------------------------------------
launcher_datas = list(icone_datas)
launcher_binaries = []
launcher_hiddenimports = []
for pacote in ('ttkbootstrap', 'PIL', 'requests'):
    tmp_ret = collect_all(pacote)
    launcher_datas += tmp_ret[0]; launcher_binaries += tmp_ret[1]; launcher_hiddenimports += tmp_ret[2]

a_launcher = Analysis(
    ['launcher.py'],
    pathex=[SPECPATH],
    binaries=launcher_binaries,
    datas=launcher_datas,
    hiddenimports=launcher_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy', 'pywinauto', 'adodbapi', 'isapi', 'pythonwin',
        'setuptools', 'pip', 'unittest', 'xmlrpc',
        'ftplib', 'multiprocessing',
    ],
    noarchive=False,
    optimize=0,
)
pyz_launcher = PYZ(a_launcher.pure)

exe_launcher = EXE(
    pyz_launcher,
    a_launcher.scripts,
    [],
    exclude_binaries=True,
    name='REQUERID Launcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['GERID LOGO.ico'],
    contents_directory='_internal_launcher',
)

coll_launcher = COLLECT(
    exe_launcher,
    a_launcher.binaries,
    a_launcher.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='REQUERID Launcher',
)
