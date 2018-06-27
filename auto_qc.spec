# -*- mode: python -*-

block_cipher = None


a = Analysis(['auto_qc.py'],
             pathex=['C:\\Users\\erothfels\\Documents\\qc_dev\\'],
             binaries=[],
             datas=[('C:\\Users\\erothfels\\Documents\\qc_dev\\gen_arcgis_file.py', '.'),
                 ('C:\\Users\\erothfels\\Documents\\qc_dev\\template.xlsx', '.'),
                 ('C:\\Users\\erothfels\\Documents\\qc_dev\\drops_template.xlsx', '.')],
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
          name='auto_qc',
          icon='C:\\Users\\erothfels\\Documents\\qc_dev\\qc.ico',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='auto_qc')
