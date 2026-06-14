# 接线表和转组态生成器 V1.1

## 功能说明

1. **生成接线总表**：从卡件排布文件（.xlsm/.xlsx/.xls）自动生成接线总表（.xlsx）
2. **生成转组态**：从接线总表生成转组态文件（.xls）

## 重要说明

- **转组态文件必须是 .xls 格式**，因为公司的"设计辅助工具"使用 ODBC Excel 驱动读取，该驱动只支持 .xls 格式
- 接线总表使用 .xlsx 格式（openpyxl 生成）
- 转组态使用 .xls 格式（xlwt 生成）

## 使用方法

### 命令行版本

```bash
# 生成接线总表
py generate_wiring_table_new.py <卡件排布文件路径>

# 生成转组态
py generate_wiring_table_new.py <接线总表路径> 转组态
```

### 窗口版本

```bash
# 运行窗口程序
py 接线总表生成器_窗口.py

# 或运行打包后的exe
dist\接线表和转组态生成器.exe
```

## 输出文件

| 功能 | 输入 | 输出 |
|------|------|------|
| 生成接线总表 | 卡件排布文件（.xlsm/.xlsx/.xls） | 接线总表.xlsx |
| 生成转组态 | 接线总表.xlsx | 转组态.xls |

## 转组态文件结构

### IO表（A-Z，26列）
- A-P：从接线总表映射数据
- Q-Z：留空（开关报警、机柜、测点组、供电、DO类型、端子1-3、KKS、系统）

### DEVICE表（A-V，22列）
- 仅表头，数据为空

### ESMA表（A-M，13列）
- 仅表头，数据为空

## 依赖库

- openpyxl：读写 xlsx 文件
- xlrd：读取 xls 文件
- xlwt：写入 xls 文件
- PySide6：窗口界面（可选）

## 安装依赖

```bash
pip install openpyxl xlrd xlwt PySide6
```

## 打包成exe

```bash
pyinstaller --onefile --windowed --name 接线表和转组态生成器 接线总表生成器_窗口.py
```

## 版本历史

- V1.0：基础接线总表生成功能
- V1.1：添加转组态生成功能，支持 xls/xlsx/xlsm 源文件
