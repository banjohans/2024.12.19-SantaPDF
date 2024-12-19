from setuptools import setup

APP = ['santapdf.py']  # Replace with the name of your Python file
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'iconfile': 'icon.icns',
    'packages': [
        'pandas', 'PyPDF2', 
        'jaraco.collections', 'jaraco.text', 'jaraco.context', 'jaraco.functools',
        'more-itertools', 'autocommand',
    ],
    'excludes': [
        'PyInstaller',
        'PySide2', 'PyQt5',
    ],
}


setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
