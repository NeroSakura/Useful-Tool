封装EXE方法：
  1.安装 PyInstaller
  pip install pyinstaller

  2.准备工作：确保你的代码文件（例如 main.py）和所有依赖项都已经正确安装
  
  3.打包代码
  在**命令行**中，切换到包含你的 Python 脚本的目录（例：cd C:\Users\zhou_\PycharmProjects\PythonProject1），然后运行以下命令来打包代码：
  pyinstaller --onefile main.py      （pyinstaller --onefile test2.py）
