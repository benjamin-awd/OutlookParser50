# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['outlook.py', 'excel.py', 'setup_db.py'],
             pathex=['C:\\Users\\benja\\OneDrive\\Desktop\\Email'],
             binaries=[],
             datas=[('emails.db', '.')],
             hiddenimports=['win32timezone'],
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
          name='outlook',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='outlook')
