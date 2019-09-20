# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['payments.py'],
             pathex=['D:\\Git\\Payments_control\\src'],
             binaries=[],
             datas=[( 'D:\\Git\\Payments_control\\src\\resources\\*.ico', 'resources' ),
                    ( 'D:\\Git\\Payments_control\\src\\resources\\*.pdf', 'resources' ),
                    ( 'D:\\Git\\Payments_control\\src\\resources\\*.png', 'resources' ),
                    ( 'D:\\Git\\Payments_control\\src\\resources\\*.txt', 'resources' )],
             hiddenimports=['babel.numbers'],
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
          name='payments',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='payments')
