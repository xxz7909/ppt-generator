# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — CCTV-17 收视日报 PPT 生成器
用法:  pyinstaller ppt_generator.spec
"""

import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# 项目根目录
_ROOT = os.path.abspath(SPECPATH)

# 收集有动态导入的包的全量数据/二进制/隐式导入
_datas, _binaries, _hiddenimports = [], [], []
for pkg in ('openpyxl', 'pandas', 'pptx', 'lxml', 'PIL', 'et_xmlfile'):
    d, b, h = collect_all(pkg)
    _datas    += d
    _binaries += b
    _hiddenimports += h

a = Analysis(
    ['GUI.py'],
    pathex=[_ROOT],
    binaries=_binaries,
    datas=_datas + [
        # 运行时需要的 json 配置
        (os.path.join(_ROOT, 'demo_layout_config.json'), '.'),
        # demo 模板 pptx — 打包进去供首次运行时自动检测
        (os.path.join(_ROOT, 'output_ppt', 'CCTV-17频道收视日报 0210 demo（简版）.pptx'), 'output_ppt'),
    ],
    hiddenimports=_hiddenimports + [
        # 本地模块
        'generate_report_config_driven',
        'generate_report',
        'data_reader',
        'ppt_config',
        'slide_utils',
        # 补充常见动态导入
        'lxml.etree',
        'lxml._elementpath',
        'lxml.objectify',
        'openpyxl.cell._writer',
        'openpyxl.workbook.child',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 不需要的模块 — 减小体积
        'matplotlib', 'scipy', 'sklearn',
        'selenium', 'pyautogui', 'bypy',
        'xlwings', 'xlrd', 'xlwt',
        'tkinterdnd2',
        'pytest', 'unittest',
    ],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CCTV17收视日报生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 窗口程序，无控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 如需图标可在此指定 .ico 路径
)
