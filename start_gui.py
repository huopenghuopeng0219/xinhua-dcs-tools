"""启动当前项目的 PySide6 窗口，复用旧项目已安装的依赖。"""

import sys
from pathlib import Path

legacy_packages = Path("F:/360MoveData/Users/霍鹏/Desktop/mimo prj/自动生成接线表的代码/Lib/site-packages")
if legacy_packages.is_dir():
    sys.path.append(str(legacy_packages))

from 接线总表生成器_窗口 import main

main()
