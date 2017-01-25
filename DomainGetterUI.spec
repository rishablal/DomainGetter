# -*- mode: python -*-

block_cipher = None


a = Analysis(['DomainGetterUI.py'],
             pathex=['/Users/rishablal/dev/DomainGetter'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=['hooks'],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='DomainGetterUI',
          debug=True,
          strip=False,
          upx=False,
          console=False )
app = BUNDLE(exe,
             name='DomainGetterUI.app',
             icon=None,
             bundle_identifier=None)
