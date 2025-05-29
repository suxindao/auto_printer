from setuptools import setup

APP = ['batch_printer_mac.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': [],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
