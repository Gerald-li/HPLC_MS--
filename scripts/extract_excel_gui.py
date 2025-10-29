# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import pandas as pd
import os
import sys
from pathlib import Path

def get_range_data_from_pandas(df, cell_range):
    """
    Extract data from a range using pandas DataFrame
    """
    num_rows, num_cols = df.shape
    
    if ':' not in cell_range:
        # Single cell: parse column letters and row
        for i in range(1, len(cell_range)):
            if cell_range[i:].isdigit():
                col_str = cell_range[:i]
                row = int(cell_range[i:])
                break
        else:
            raise ValueError(f"Invalid cell format: {cell_range}")
        col_idx = column_index_from_string(col_str) - 1  # 0-based
        
        if row > num_rows:
            print(f"Warning: Row {row} in {cell_range} exceeds sheet rows ({num_rows})")
            return None
        if col_idx >= num_cols:
            print(f"Warning: Column {col_str} ({col_idx+1}) in {cell_range} exceeds sheet columns ({num_cols})")
            return None
        return df.iloc[row-1, col_idx]
    else:
        # Range like F1:P1 or A1:C3 or AA1:AB5
        start_cell, end_cell = cell_range.split(':')
        
        # Parse start
        for i in range(1, len(start_cell)):
            if start_cell[i:].isdigit():
                start_col_str = start_cell[:i]
                start_row = int(start_cell[i:])
                break
        else:
            raise ValueError(f"Invalid start cell: {start_cell}")
        
        # Parse end
        for i in range(1, len(end_cell)):
            if end_cell[i:].isdigit():
                end_col_str = end_cell[:i]
                end_row = int(end_cell[i:])
                break
        else:
            raise ValueError(f"Invalid end cell: {end_cell}")
        
        start_col_idx = column_index_from_string(start_col_str) - 1
        end_col_idx = column_index_from_string(end_col_str) - 1
        
        if start_row > num_rows or end_row > num_rows:
            print(f"Warning: Rows in {cell_range} exceed sheet rows ({num_rows})")
            return []
        
        if end_col_idx >= num_cols:
            print(f"Warning: Columns in {cell_range} exceed sheet columns ({num_cols}); returning available data")
        
        values = []
        for row_idx in range(start_row-1, min(end_row, num_rows)):
            for col_idx in range(start_col_idx, min(end_col_idx + 1, num_cols)):
                values.append(df.iloc[row_idx, col_idx])
            # If range extends beyond cols, append None for missing
            if end_col_idx >= num_cols:
                values.extend([None] * (end_col_idx + 1 - num_cols))
        return values

def parse_position_args(position_args):
    """
    Parse position arguments in format "cell_range-column_header"
    Returns list of tuples: (cell_range, column_header)
    """
    parsed_positions = []
    
    for arg in position_args:
        # Split by last dash to separate cell_range and column_header
        if '-' in arg:
            # Find the last dash that separates cell range from column header
            parts = arg.split('-')
            
            # Try to find the split point where the left part is a valid cell reference
            for i in range(len(parts)-1, 0, -1):
                left_part = '-'.join(parts[:i])
                right_part = '-'.join(parts[i:])
                
                # Check if left_part is a valid cell reference (contains letters and numbers)
                if any(char.isdigit() for char in left_part) and any(char.isalpha() for char in left_part):
                    parsed_positions.append((left_part, right_part))
                    break
            else:
                # If no valid split found, use the whole string as both cell_range and column_header
                parsed_positions.append((arg, arg))
        else:
            # No dash found, use the same string for both
            parsed_positions.append((arg, arg))
    
    return parsed_positions

def extract_excel_info_single_file(input_excel_path, out_excel_path, position_args):
    """
    Extract data from a single Excel file
    """
    results = []
    
    # Parse position arguments
    parsed_positions = parse_position_args(position_args)
    
    # Read all sheets with pandas
    try:
        excel_file = pd.ExcelFile(input_excel_path, engine='openpyxl', engine_kwargs={'read_only': True, 'data_only': True})
    except Exception as e:
        raise Exception(f"Failed to open Excel file: {e}")
    
    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(input_excel_path, sheet_name=sheet_name, header=None, engine='openpyxl', engine_kwargs={'read_only': True, 'data_only': True})
        except Exception as e:
            print(f"Failed to read sheet {sheet_name}: {e}")
            continue
        
        rowdict = {'sheet': sheet_name, 'source_file': os.path.basename(input_excel_path)}
        
        for cell_range, column_header in parsed_positions:
            # Extract values from range/position using pandas
            try:
                val = get_range_data_from_pandas(df, cell_range)
                rowdict[column_header] = val
            except Exception as e:
                print(f"Error extracting {cell_range} from {sheet_name}: {e}")
                rowdict[column_header] = None
        
        results.append(rowdict)
    
    # Prepare DataFrame and save
    if results:
        df_output = pd.DataFrame(results)
        
        # Reorder columns to have 'sheet' and 'source_file' first, then the specified column headers
        final_cols = ['sheet', 'source_file'] + [column_header for _, column_header in parsed_positions]
        
        # Only include columns that actually exist in the DataFrame
        final_cols = [col for col in final_cols if col in df_output.columns]
        
        df_output = df_output[final_cols]
        df_output.to_excel(out_excel_path, index=False)
        return len(results)
    else:
        raise Exception("No data extracted from any sheet")

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据提取工具")
        self.root.geometry("800x700")
        
        # 存储输入文件
        self.input_files = []
        
        # 默认参数
        self.default_params = [
            "F4-样品名称", "N7-类型", "F6-进样体积", "A15-保留时间", 
            "N8-样品含量", "N6-位置", "E15-峰宽[min]", "G15-峰面积", 
            "H15-峰高", "K15-峰面积%", "C15-类型", "G16-总和"
        ]
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 输入文件选择
        ttk.Label(main_frame, text="输入Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_listbox = tk.Listbox(main_frame, height=5)
        self.file_listbox.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        file_button_frame = ttk.Frame(main_frame)
        file_button_frame.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(file_button_frame, text="移除选中", command=self.remove_selected_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_files).pack(side=tk.LEFT, padx=2)
        
        # 输出设置
        ttk.Label(main_frame, text="输出设置:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(output_frame, text="输出目录:").pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(output_frame, textvariable=self.output_dir_var, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.browse_output_dir).pack(side=tk.LEFT)
        
        ttk.Label(main_frame, text="输出文件名:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_name_var = tk.StringVar(value="提取结果.xlsx")
        ttk.Entry(main_frame, textvariable=self.output_name_var).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # 输出模式
        self.output_mode = tk.StringVar(value="single")
        output_mode_frame = ttk.Frame(main_frame)
        output_mode_frame.grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Radiobutton(output_mode_frame, text="合并所有文件到一个Excel", variable=self.output_mode, value="single").pack(side=tk.LEFT)
        ttk.Radiobutton(output_mode_frame, text="每个输入文件单独输出", variable=self.output_mode, value="multiple").pack(side=tk.LEFT)
        
        # 参数设置
        ttk.Label(main_frame, text="提取参数 (格式: 单元格-列名):").grid(row=5, column=0, sticky=tk.W, pady=5)
        
        param_frame = ttk.Frame(main_frame)
        param_frame.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 参数文本框
        self.param_text = scrolledtext.ScrolledText(param_frame, height=8, width=60)
        self.param_text.pack(fill=tk.BOTH, expand=True)
        
        # 初始化默认参数
        self.load_default_params()
        
        # 参数按钮框架
        param_button_frame = ttk.Frame(main_frame)
        param_button_frame.grid(row=6, column=1, sticky=tk.W, pady=5)
        
        ttk.Button(param_button_frame, text="恢复默认参数", command=self.load_default_params).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_button_frame, text="添加参数", command=self.add_parameter).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_button_frame, text="清空参数", command=self.clear_parameters).pack(side=tk.LEFT, padx=2)
        
        # 日志输出
        ttk.Label(main_frame, text="运行日志:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.log_text = scrolledtext.ScrolledText(main_frame, height=10, width=80)
        self.log_text.grid(row=7, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 运行按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=1, pady=10)
        
        ttk.Button(button_frame, text="开始提取", command=self.start_extraction).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
        # 配置权重使界面可伸缩
        main_frame.rowconfigure(7, weight=1)
    
    def load_default_params(self):
        """加载默认参数"""
        self.param_text.delete(1.0, tk.END)
        for param in self.default_params:
            self.param_text.insert(tk.END, param + "\n")
    
    def add_parameter(self):
        """添加新参数"""
        new_param = tk.simpledialog.askstring("添加参数", "请输入参数 (格式: 单元格-列名):")
        if new_param:
            self.param_text.insert(tk.END, new_param + "\n")
    
    def clear_parameters(self):
        """清空所有参数"""
        self.param_text.delete(1.0, tk.END)
    
    def add_files(self):
        """添加Excel文件"""
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.input_files:
                self.input_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.input_files.pop(index)
            self.file_listbox.delete(index)
    
    def clear_files(self):
        """清空文件列表"""
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)
    
    def browse_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir_var.set(directory)
    
    def log(self, message):
        """添加日志信息"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def get_parameters(self):
        """从文本框获取参数列表"""
        text = self.param_text.get(1.0, tk.END).strip()
        if not text:
            return []
        return [line.strip() for line in text.split('\n') if line.strip()]
    
    def start_extraction(self):
        """开始提取过程"""
        if not self.input_files:
            messagebox.showerror("错误", "请至少选择一个Excel文件")
            return
        
        parameters = self.get_parameters()
        if not parameters:
            messagebox.showerror("错误", "请至少设置一个提取参数")
            return
        
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录")
            return
        
        output_name = self.output_name_var.get()
        if not output_name:
            messagebox.showerror("错误", "请输入输出文件名")
            return
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        self.log("开始提取数据...")
        
        try:
            if self.output_mode.get() == "single":
                # 合并所有文件到一个Excel
                self.extract_to_single_file(parameters, output_dir, output_name)
            else:
                # 每个文件单独输出
                self.extract_to_multiple_files(parameters, output_dir, output_name)
            
            messagebox.showinfo("完成", "数据提取完成！")
            
        except Exception as e:
            error_msg = f"提取过程中出现错误: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def extract_to_single_file(self, parameters, output_dir, output_name):
        """提取所有文件数据到一个Excel文件"""
        all_results = []
        output_path = os.path.join(output_dir, output_name)
        
        for input_file in self.input_files:
            self.log(f"处理文件: {os.path.basename(input_file)}")
            
            try:
                # 读取Excel文件
                excel_file = pd.ExcelFile(input_file, engine='openpyxl', engine_kwargs={'read_only': True, 'data_only': True})
                parsed_positions = parse_position_args(parameters)
                
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(input_file, sheet_name=sheet_name, header=None, 
                                         engine='openpyxl', engine_kwargs={'read_only': True, 'data_only': True})
                    except Exception as e:
                        self.log(f"  警告: 无法读取工作表 {sheet_name}: {e}")
                        continue
                    
                    rowdict = {'sheet': sheet_name, 'source_file': os.path.basename(input_file)}
                    
                    for cell_range, column_header in parsed_positions:
                        try:
                            val = get_range_data_from_pandas(df, cell_range)
                            rowdict[column_header] = val
                        except Exception as e:
                            self.log(f"  警告: 无法提取 {cell_range}: {e}")
                            rowdict[column_header] = None
                    
                    all_results.append(rowdict)
                    self.log(f"  已提取工作表: {sheet_name}")
                    
            except Exception as e:
                self.log(f"  错误处理文件 {os.path.basename(input_file)}: {e}")
                continue
        
        if all_results:
            df_output = pd.DataFrame(all_results)
            final_cols = ['sheet', 'source_file'] + [column_header for _, column_header in parsed_positions]
            final_cols = [col for col in final_cols if col in df_output.columns]
            df_output = df_output[final_cols]
            df_output.to_excel(output_path, index=False)
            self.log(f"数据已保存到: {output_path}")
            self.log(f"总共提取了 {len(all_results)} 个工作表的数据")
        else:
            raise Exception("没有提取到任何数据")
    
    def extract_to_multiple_files(self, parameters, output_dir, output_name):
        """为每个输入文件创建单独的输出文件"""
        base_name = os.path.splitext(output_name)[0]
        extension = os.path.splitext(output_name)[1] or '.xlsx'
        
        for input_file in self.input_files:
            self.log(f"处理文件: {os.path.basename(input_file)}")
            
            input_base = os.path.splitext(os.path.basename(input_file))[0]
            output_file = f"{base_name}_{input_base}{extension}"
            output_path = os.path.join(output_dir, output_file)
            
            try:
                count = extract_excel_info_single_file(input_file, output_path, parameters)
                self.log(f"  成功提取 {count} 个工作表，保存到: {output_file}")
            except Exception as e:
                self.log(f"  错误: {e}")

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelExtractorApp(root)
    root.mainloop()