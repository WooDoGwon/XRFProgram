# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import re

project_dir = Path.cwd()
main_path = project_dir / 'main.py'
main_text = main_path.read_text(encoding='utf-8')
version_match = re.search(r"APP_VERSION\s*=\s*['\"]([^'\"]+)['\"]", main_text)
APP_VERSION = version_match.group(1) if version_match else '1.0.0'
APP_EXE_NAME = f'XRF_Report_Auto_Input_v{APP_VERSION}'
logo_path = project_dir / 'Logo.png'
app_icon_path = project_dir / 'AppIcon.ico'
app_icon_png_path = project_dir / 'AppIcon.png'
asset_paths = (logo_path, app_icon_path, app_icon_png_path)
data_files = [(str(asset_path), '.') for asset_path in asset_paths if asset_path.exists()]
app_icon_arg = str(app_icon_path) if app_icon_path.exists() else None

version_parts = [part for part in APP_VERSION.split('.') if part.isdigit()]
while len(version_parts) < 4:
    version_parts.append('0')
version_parts = version_parts[:4]
version_tuple = ', '.join(version_parts)
version_text = '.'.join(version_parts)
version_info_path = project_dir / 'pyinstaller_version_info.txt'
version_info_path.write_text(
    f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({version_tuple}),
    prodvers=({version_tuple}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        '040904B0',
        [
        StringStruct('CompanyName', 'XRF Tools'),
        StringStruct('FileDescription', 'XRF Report Auto Input'),
        StringStruct('FileVersion', '{version_text}'),
        StringStruct('InternalName', '{APP_EXE_NAME}'),
        StringStruct('OriginalFilename', '{APP_EXE_NAME}.exe'),
        StringStruct('ProductName', 'XRF Report Auto Input'),
        StringStruct('ProductVersion', '{version_text}')
        ])
      ]),
    VarFileInfo([VarStruct('Translation', [1042, 1200])])
  ]
)
""",
    encoding='utf-8',
)

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=data_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=str(version_info_path),
    icon=app_icon_arg,
)
