# -*- mode: python -*-

block_cipher = None


a = Analysis(['mailmerge.py'],
             pathex=['C:\\Python36-32\\Python Scripts\\Mail Merge App'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='mailmerge',
          debug=False,
          strip=False,
          upx=True,
          console=True , icon='email-icon.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='mailmerge')
