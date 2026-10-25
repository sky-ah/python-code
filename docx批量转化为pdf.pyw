import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess
import sys
import pythoncom
import win32com.client
import time

class DocxToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF转换器 (WPS版)")
        self.root.geometry("600x450")
        
        # 设置样式
        self.root.configure(bg='#f0f0f0')
        
        # 创建主框架
        main_frame = tk.Frame(root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = tk.Label(main_frame, text="Word文档批量转PDF (支持WPS)", 
                               font=('Arial', 16, 'bold'), bg='#f0f0f0', fg='#2E86C1')
        title_label.pack(pady=(0, 20))
        
        # 源文件夹选择
        source_frame = tk.Frame(main_frame, bg='#f0f0f0')
        source_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(source_frame, text="源文件夹:", bg='#f0f0f0', 
                font=('Arial', 10)).pack(side=tk.LEFT)
        
        self.source_path = tk.StringVar()
        source_entry = tk.Entry(source_frame, textvariable=self.source_path, 
                               width=40, font=('Arial', 10))
        source_entry.pack(side=tk.LEFT, padx=(10, 10), fill=tk.X, expand=True)
        
        tk.Button(source_frame, text="浏览", command=self.select_source_folder,
                 bg='#4CAF50', fg='white', font=('Arial', 10)).pack(side=tk.RIGHT)
        
        # 目标文件夹选择
        target_frame = tk.Frame(main_frame, bg='#f0f0f0')
        target_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(target_frame, text="目标文件夹:", bg='#f0f0f0', 
                font=('Arial', 10)).pack(side=tk.LEFT)
        
        self.target_path = tk.StringVar()
        target_entry = tk.Entry(target_frame, textvariable=self.target_path, 
                               width=40, font=('Arial', 10))
        target_entry.pack(side=tk.LEFT, padx=(10, 10), fill=tk.X, expand=True)
        
        tk.Button(target_frame, text="浏览", command=self.select_target_folder,
                 bg='#4CAF50', fg='white', font=('Arial', 10)).pack(side=tk.RIGHT)
        
        # 转换按钮
        convert_frame = tk.Frame(main_frame, bg='#f0f0f0')
        convert_frame.pack(pady=20)
        
        self.convert_btn = tk.Button(convert_frame, text="开始转换", 
                                     command=self.start_conversion,
                                     bg='#2196F3', fg='white', font=('Arial', 12, 'bold'),
                                     width=20, height=2)
        self.convert_btn.pack()
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=400)
        self.progress.pack(pady=10)
        
        # 状态信息
        self.status_label = tk.Label(main_frame, text="就绪", bg='#f0f0f0', 
                                     font=('Arial', 10), fg='#666')
        self.status_label.pack()
        
        # 日志区域
        log_frame = tk.Frame(main_frame, bg='#f0f0f0')
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 创建日志文本框和滚动条的框架
        log_text_frame = tk.Frame(log_frame, bg='#f0f0f0')
        log_text_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(log_text_frame, text="转换日志:", bg='#f0f0f0', 
                font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        
        # 创建文本框和滚动条
        text_frame = tk.Frame(log_text_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(text_frame, height=8, font=('Arial', 9))
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        # 初始化
        self.doc_files = []
        self.conversion_thread = None
        self.stop_conversion = False
        
        # 启动时检查依赖
        self.root.after(100, self.check_dependencies)
        
    def check_dependencies(self):
        """检查必要的依赖"""
        try:
            import win32com
            import pythoncom
            self.log("✓ 依赖库检查通过")
        except ImportError as e:
            self.log("✗ 缺少必要的依赖库，尝试安装...")
            self.install_dependencies()
            
    def install_dependencies(self):
        """安装必要的依赖"""
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            self.log("✓ 依赖库安装完成，请重启程序")
            messagebox.showinfo("安装完成", "依赖库安装完成，请重启程序")
        except Exception as e:
            self.log(f"✗ 依赖库安装失败: {str(e)}")
            
    def select_source_folder(self):
        folder = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if folder:
            self.source_path.set(folder)
            self.log(f"已选择源文件夹: {folder}")
            # 显示找到的文件数量
            files = self.get_word_files(folder)
            self.log(f"找到 {len(files)} 个Word文档")
            
    def select_target_folder(self):
        folder = filedialog.askdirectory(title="选择PDF保存文件夹")
        if folder:
            self.target_path.set(folder)
            self.log(f"已选择目标文件夹: {folder}")
            
    def log(self, message):
        """在日志区域添加消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def get_word_files(self, folder):
        """获取文件夹中的所有Word文档（支持.doc和.docx）"""
        word_files = []
        try:
            for file in os.listdir(folder):
                if file.lower().endswith(('.docx', '.doc')):
                    word_files.append(file)
        except Exception as e:
            self.log(f"读取文件夹失败: {str(e)}")
        return word_files
        
    def convert_with_wps(self, doc_path, pdf_path):
        """使用WPS转换单个文件"""
        try:
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 尝试创建WPS应用程序对象
            try:
                wps = win32com.client.Dispatch("wps.Application")
            except:
                try:
                    wps = win32com.client.Dispatch("KWPS.Application")
                except:
                    return False, "无法启动WPS，请确保WPS已安装"
            
            # 设置WPS不可见
            wps.Visible = False
            
            # 打开文档
            doc = wps.Documents.Open(doc_path)
            
            # 另存为PDF (17表示PDF格式)
            doc.SaveAs(pdf_path, 17)
            
            # 关闭文档
            doc.Close()
            
            # 退出WPS
            wps.Quit()
            
            # 释放COM资源
            pythoncom.CoUninitialize()
            
            # 等待文件写入完成
            time.sleep(0.5)
            
            return True, "转换成功"
            
        except Exception as e:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False, str(e)
            
    def start_conversion(self):
        """开始转换过程"""
        source_folder = self.source_path.get()
        target_folder = self.target_path.get()
        
        if not source_folder or not target_folder:
            messagebox.showwarning("警告", "请选择源文件夹和目标文件夹！")
            return
            
        if not os.path.exists(source_folder):
            messagebox.showerror("错误", "源文件夹不存在！")
            return
            
        # 获取Word文档列表
        self.doc_files = self.get_word_files(source_folder)
        
        if not self.doc_files:
            messagebox.showinfo("提示", "源文件夹中没有找到Word文档（.doc/.docx）")
            return
            
        # 创建目标文件夹（如果不存在）
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)
            self.log(f"创建目标文件夹: {target_folder}")
            
        # 禁用转换按钮
        self.convert_btn.config(state='disabled')
        self.stop_conversion = False
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        # 在新线程中执行转换
        self.conversion_thread = threading.Thread(target=self.do_conversion)
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
        
    def do_conversion(self):
        """执行转换（在后台线程中运行）"""
        source_folder = self.source_path.get()
        target_folder = self.target_path.get()
        
        total_files = len(self.doc_files)
        successful = 0
        failed = 0
        failed_files = []
        
        self.log(f"找到 {total_files} 个Word文档，开始转换...")
        self.log("-" * 50)
        
        for i, doc_file in enumerate(self.doc_files):
            if self.stop_conversion:
                self.log("转换已停止")
                break
                
            # 更新进度条
            progress_value = ((i) / total_files) * 100
            self.root.after(0, self.update_progress, progress_value)
            self.root.after(0, self.update_status, f"正在转换: {doc_file}")
            
            # 构建完整路径
            doc_path = os.path.join(source_folder, doc_file)
            pdf_name = os.path.splitext(doc_file)[0] + '.pdf'
            pdf_path = os.path.join(target_folder, pdf_name)
            
            self.log(f"[{i+1}/{total_files}] 转换: {doc_file}")
            
            # 执行转换
            success, message = self.convert_with_wps(doc_path, pdf_path)
            
            if success:
                successful += 1
                self.log(f"  ✓ 成功 -> {pdf_name}")
            else:
                failed += 1
                failed_files.append(doc_file)
                self.log(f"  ✗ 失败: {message}")
                
        # 完成
        self.root.after(0, self.update_progress, 100)
        self.root.after(0, self.update_status, f"转换完成！成功: {successful}, 失败: {failed}")
        self.root.after(0, self.enable_convert_button)
        
        self.log("-" * 50)
        self.log(f"转换完成！总计: {total_files}, 成功: {successful}, 失败: {failed}")
        
        if failed_files:
            self.log("\n转换失败的文件:")
            for f in failed_files:
                self.log(f"  - {f}")
                
        self.root.after(0, self.show_completion_message, successful, failed)
        
    def update_progress(self, value):
        self.progress['value'] = value
        
    def update_status(self, message):
        self.status_label.config(text=message)
        
    def enable_convert_button(self):
        self.convert_btn.config(state='normal')
        
    def show_completion_message(self, successful, failed):
        messagebox.showinfo("完成", f"转换完成！\n成功: {successful} 个文件\n失败: {failed} 个文件")

def main():
    # 创建主窗口
    root = tk.Tk()
    app = DocxToPdfConverter(root)
    
    # 设置窗口居中
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    main()