# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['newDesign.py'],
             pathex=['C:\\Python38\\Lib\\site-packages\\', 'C:\\Users\\Evolutivelabs\\Desktop\\VBA們\\newDesign'],
             binaries=[],
             datas=[('google_api_python_client-1.9.3.dist-info','.'),('googleapiclient','.')],
             hiddenimports=['google-api-python-client'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='newDesign',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True , icon='main.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='newDesign')
