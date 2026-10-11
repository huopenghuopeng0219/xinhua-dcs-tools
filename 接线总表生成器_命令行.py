# -*- coding: utf-8 -*-
"""
接线总表生成器 V1.0
命令行版本 - 双击运行，按提示操作
"""
import os
import sys

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def main():
    print("=" * 50)
    print("       接线总表生成器 V1.0")
    print("=" * 50)
    print()
    
    # 获取文件路径
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        print("请输入Excel文件路径（或把文件拖入此窗口）：")
        print("支持格式：.xlsm .xlsx .xls")
        print()
        file_path = input("文件路径: ").strip().strip('"')
    
    if not file_path or not os.path.exists(file_path):
        print(f"错误：文件不存在 - {file_path}")
        input("按回车键退出...")
        return
    
    # 选择输出格式
    print()
    print("请选择输出格式：")
    print("  1. xlsx (Excel 2007+)")
    print("  2. xls (WPS兼容)")
    print()
    choice = input("请输入选择 (1或2，默认1): ").strip()
    ext = "xls" if choice == "2" else "xlsx"
    
    print()
    print(f"源文件: {os.path.basename(file_path)}")
    print(f"输出格式: {ext}")
    print()
    print("正在生成接线总表，请稍候...")
    
    try:
        from generate_wiring_table_new import WiringTableGenerator
        
        generator = WiringTableGenerator(file_path)
        output_dir = os.path.dirname(file_path)
        output_path = os.path.join(output_dir, f"接线总表.{ext}")
        
        count = generator.generate(output_path)
        
        print()
        print("=" * 50)
        print("  生成完成！")
        print("=" * 50)
        print(f"  文件位置: {output_path}")
        print(f"  数据行数: {count}")
        print("=" * 50)
        
    except Exception as e:
        print()
        print(f"错误：生成失败 - {str(e)}")
    
    print()
    input("按回车键退出...")


if __name__ == "__main__":
    main()
