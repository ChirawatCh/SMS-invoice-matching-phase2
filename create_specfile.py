import os

# Define the directory containing the Python scripts
scripts_directory = 'scripts'

# List all Python files in the directory with their full paths
python_files = [os.path.join(scripts_directory, f) for f in os.listdir(scripts_directory) if f.endswith('.py')]

# Create the spec file content with explicit paths for each Python file
spec_content = f"""
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    {python_files},
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GeneratedExecutable',  # Replace with your desired name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""

# Write the content to a .spec file
with open('generated_EXE.spec', 'w') as spec_file:
    spec_file.write(spec_content)
