# -*- coding: utf-8 -*-
"""
接线总表生成器 V1.0
图形界面版本 - 点击选择Excel文件，一键生成接线总表
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys

# 添加当前目录到路径，以便导入生成器模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_wiring_table_new import WiringTableGenerator
from material_difference import generate_difference, output_path_for


class WiringTableApp:
    """接线总表生成器图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("接线表和物料差异表生成器")
        self.root.geometry("540x370")
        self.root.resizable(False, False)
        
        self.file_path = None
        self.output_format = tk.StringVar(value="xlsx")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """创建界面组件"""
        # 标题
        title_label = tk.Label(self.root, text="接线表和物料差异表生成器", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=10)
        
        # 文件选择区域
        file_frame = tk.LabelFrame(self.root, text="选择源文件", padx=10, pady=10)
        file_frame.pack(padx=20, pady=5, fill="x")
        
        self.file_label = tk.Label(file_frame, text="未选择文件", anchor="w", fg="gray")
        self.file_label.pack(side="left", fill="x", expand=True)
        
        select_btn = tk.Button(file_frame, text="选择文件", command=self._select_file, width=10)
        select_btn.pack(side="right")
        
        # 输出格式选择
        format_frame = tk.LabelFrame(self.root, text="输出格式", padx=10, pady=10)
        format_frame.pack(padx=20, pady=5, fill="x")
        
        tk.Radiobutton(format_frame, text="xlsx (Excel 2007+)", variable=self.output_format, value="xlsx").pack(side="left", padx=10)
        tk.Radiobutton(format_frame, text="xls (WPS兼容)", variable=self.output_format, value="xls").pack(side="left", padx=10)
        
        # 生成按钮
        self.generate_btn = tk.Button(self.root, text="生成接线总表", command=self._generate, 
                                       font=("微软雅黑", 12, "bold"), height=2, state="disabled")
        self.generate_btn.pack(pady=10)

        self.diff_btn = tk.Button(self.root, text="生成差异表", command=self._generate_difference,
                                  font=("微软雅黑", 12, "bold"), height=2, state="disabled")
        self.diff_btn.pack(pady=3)
        
        # 状态栏
        self.status_label = tk.Label(self.root, text="请先选择Excel源文件", fg="blue")
        self.status_label.pack(pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(padx=20, pady=5, fill="x")
    
    def _select_file(self):
        """选择源Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择带卡件排布的Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsm *.xlsx *.xls"),
                ("带宏的Excel文件", "*.xlsm"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename, fg="black")
            self.generate_btn.config(state="normal")
            self.diff_btn.config(state="normal")
            self.status_label.config(text=f"已选择: {filename}", fg="green")
    
    def _generate(self):
        """生成接线总表"""
        if not self.file_path:
            messagebox.showwarning("提示", "请先选择源文件！")
            return
        
        # 禁用按钮，显示进度
        self.generate_btn.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="正在生成接线总表，请稍候...", fg="orange")
        self.root.update()
        
        try:
            # 创建生成器
            generator = WiringTableGenerator(self.file_path)
            
            # 确定输出路径和格式
            output_dir = os.path.dirname(os.path.abspath(__file__))
            ext = self.output_format.get()
            output_path = os.path.join(output_dir, f"接线总表.{ext}")
            
            # 生成
            count = generator.generate(output_path)
            
            # 成功
            self.progress.stop()
            self.status_label.config(text=f"生成完成！共 {count} 行数据", fg="green")
            messagebox.showinfo("成功", 
                f"接线总表已生成！\n\n"
                f"文件位置: {output_path}\n"
                f"数据行数: {count}\n"
                f"输出格式: {ext}")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text=f"生成失败: {str(e)}", fg="red")
            messagebox.showerror("错误", f"生成失败:\n{str(e)}")
        
        finally:
            self.generate_btn.config(state="normal")

    def _generate_difference(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先选择机柜设计文件！")
            return
        if not self.file_path.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showwarning("提示", "差异表需要 .xlsx 或 .xlsm 机柜设计文件")
            return
        target = output_path_for(self.file_path)
        if target.exists() and not messagebox.askyesno("覆盖文件", f"差异表已存在，是否覆盖？\n{target}"):
            return
        self.generate_btn.config(state="disabled")
        self.diff_btn.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="正在统计 -MX 工作表的整件物料，请稍候...", fg="orange")
        self.root.update()
        try:
            path, count = generate_difference(self.file_path)
            self.status_label.config(text=f"生成完成！共 {count} 种代号", fg="green")
            messagebox.showinfo("成功", f"差异表已生成！\n\n文件位置：{path}")
        except Exception as error:
            self.status_label.config(text=f"生成失败：{error}", fg="red")
            messagebox.showerror("错误", f"生成失败：\n{error}")
        finally:
            self.progress.stop()
            self.generate_btn.config(state="normal")
            self.diff_btn.config(state="normal")


def main():
    root = tk.Tk()
    app = WiringTableApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
