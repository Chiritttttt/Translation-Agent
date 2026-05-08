# -*- mode: python ; coding: utf-8 -*-
"""
TranslationAgent - PyInstaller ONEFILE spec
============================================
Builds a single standalone executable.

Usage:
    pyinstaller TranslationAgent-onefile.spec --clean
"""

import os
import sys
import platform

# ── Project root (directory containing this .spec file) ──
SPEC_DIR = os.path.dirname(os.path.abspath(SPECPATH))
os.chdir(SPEC_DIR)

# ── Path to app icon ──
ICON_PATH = os.path.join(SPEC_DIR, 'app', 'image.png')

# ── Determine icon argument based on platform ──
if platform.system() == 'Windows':
    ICON_ARG = ICON_PATH if ICON_PATH.endswith('.ico') else 'NONE'
elif platform.system() == 'Darwin':
    ICON_ARG = ICON_PATH if os.path.exists(ICON_PATH) else 'NONE'
else:
    ICON_ARG = 'NONE'

# ── Collect data files ──
datas = []

# App icon → bundled beside the exe
if os.path.exists(ICON_PATH):
    datas.append((ICON_PATH, 'app'))

# tiktoken BPE data (critical — otherwise tiktoken crashes at runtime)
try:
    from PyInstaller.utils.hooks import collect_data_files
    datas += collect_data_files('tiktoken')
except Exception:
    pass

# simplemma language data
try:
    from PyInstaller.utils.hooks import collect_data_files
    datas += collect_data_files('simplemma')
except Exception:
    pass

# ── Hidden imports ──
hiddenimports = [
    # ── Local modules ──
    'glossary',
    'subtitle_handler',
    'file_handler',

    # ── AI / LLM ──
    'openai',
    'tiktoken',
    'tiktoken_ext',
    'tiktoken_ext.openai_public',
    'tiktoken_core',

    # ── Text processing ──
    'langchain_text_splitters',
    'simplemma',

    # ── HTTP / HTML ──
    'requests',
    'urllib3',
    'charset_normalizer',
    'certifi',
    'bs4',
    'bs4.element',
    'html.parser',

    # ── Document I/O ──
    'fitz',
    'PyMuPDF',
    'docx',
    'docx.opc',
    'docx.opc.constants',
    'docx.opc.oxml',
    'docx.oxml',
    'docx.oxml.ns',
    'docx.oxml.coreprops',
    'docx.oxml.text.paragraph',
    'docx.shared',
    'docx.enum',
    'docx.enum.text',
    'docx.enum.table',
    'pptx',
    'pptx.enum',
    'pptx.enum.shapes',
    'pptx.oxml',
    'openpyxl',
    'openpyxl.styles',
    'reportlab',
    'reportlab.pdfbase',
    'reportlab.pdfbase.ttfonts',
    'reportlab.pdfbase.pdfmetrics',
    'reportlab.pdfbase.cidfonts',
    'reportlab.lib.pagesizes',
    'reportlab.lib.styles',
    'reportlab.lib.units',
    'reportlab.lib.colors',
    'reportlab.lib.enums',
    'reportlab.platypus',
    'reportlab.platypus.flowables',

    # ── Image / OCR ──
    'PIL',
    'PIL.Image',
    'PIL.ImageFont',
    'PIL.ImageDraw',
    'PIL.ImageFilter',
    'PIL.ImageOps',
    'pdf2image',

    # ── Subtitle ──
    'pysrt',

    # ── Environment ──
    'dotenv',

    # ── Utilities ──
    'joblib',
    'icecream',
    'asttokens',
    'pygments',
    'execute',
    'colorama',

    # ── PyQt6 sub-modules ──
    'PyQt6',
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
    'PyQt6.sip',
]

# ── Platform-specific imports ──
if platform.system() == 'Windows':
    hiddenimports += [
        'pywin32',
        'win32com',
        'win32com.client',
        'win32api',
        'win32file',
    ]

# ── Collect submodules for key packages ──
try:
    from PyInstaller.utils.hooks import collect_submodules
    hiddenimports += collect_submodules('openai')
    hiddenimports += collect_submodules('fitz')
except Exception:
    pass

# ═══════════════════════════════════════════════════════════
# Analysis
# ═══════════════════════════════════════════════════════════
a = Analysis(
    ['gui.py'],
    pathex=[SPEC_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'black',
        'flake8',
        'ruff',
        'mypy',
        'gradio',
    ],
    noarchive=False,
    optimize=0,
)

# ═══════════════════════════════════════════════════════════
# PYZ (Python bytecode archive)
# ═══════════════════════════════════════════════════════════
pyz = PYZ(a.pure)

# ═══════════════════════════════════════════════════════════
# EXE (onefile — everything packed into a single executable)
# ═══════════════════════════════════════════════════════════
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TranslationAgent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'vcruntime140_1.dll',
        'python3*.dll',
        'PyQt6*.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_ARG if ICON_ARG != 'NONE' else None,
)
