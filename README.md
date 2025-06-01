### 编译程序，打印指定目录下新增文件(pdf或者excel文件)
    pyinstaller -F auto_printer.py
    
### 编译程序，打印指定目录下文件(pdf或者excel文件)
    pyinstaller -F batch_printer.py
### 从左到右向日葵
    657 098 865
    688 015 466
    604 451 356

### 构建 mac 的 app程序

    1. 安装依赖：
    
    pip install py2app
    
    2. 在项目目录下创建 setup.py：
    
    from setuptools import setup
    
    APP = ['your_script.py']
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
    
    3. 构建应用：
    
    python setup.py py2app
    
    构建完成后，.app 文件会在 dist/ 目录中生成。

