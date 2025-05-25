import os
import sys
import json
import base64
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, font
from PIL import Image, ImageTk
import pandas as pd
import threading
from datetime import datetime
import requests
import time
import io

class ImageRecognitionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片批量识别工具")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f0f0")
        
        # 设置应用样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("TLabel", font=("Arial", 10))
        
        # 创建主框架，使用grid布局以支持窗口大小调整
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        # 配置根窗口的行列权重，使其可调整大小
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        
        # 创建顶部控制区域
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.main_frame.columnconfigure(0, weight=1)
        
        # 文件夹选择
        self.folder_label = ttk.Label(self.control_frame, text="图片文件夹:")
        self.folder_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(self.control_frame, textvariable=self.folder_path)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.control_frame.columnconfigure(1, weight=1)
        
        self.browse_button = ttk.Button(self.control_frame, text="浏览...", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        # API密钥输入
        self.api_label = ttk.Label(self.control_frame, text="豆包API密钥:")
        self.api_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.api_key = tk.StringVar()
        self.api_key.set("")  # 设置默认API密钥
        self.api_entry = ttk.Entry(self.control_frame, textvariable=self.api_key)
        self.api_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # 模型ID输入 - 默认设置为用户指定的模型
        self.model_label = ttk.Label(self.control_frame, text="模型ID:")
        self.model_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        self.model_id = tk.StringVar()
        self.model_id.set("doubao-1.5-vision-pro")  # 基础模型ID
        self.model_entry = ttk.Entry(self.control_frame, textvariable=self.model_id)
        self.model_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        # 模型版本选择
        self.version_label = ttk.Label(self.control_frame, text="模型版本:")
        self.version_label.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        
        self.version_var = tk.StringVar()
        self.version_combobox = ttk.Combobox(self.control_frame, textvariable=self.version_var, width=15)
        self.version_combobox['values'] = ["", "250328", "250428", "250115"]
        self.version_combobox.current(1)  # 默认选择250328版本
        self.version_combobox.grid(row=2, column=3, padx=5, pady=5)
        
        # 添加自定义提示词输入框 - 默认设置为用户指定的提示词
        self.prompt_label = ttk.Label(self.control_frame, text="自定义提示词:")
        self.prompt_label.grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        
        self.prompt_text = tk.Text(self.control_frame, height=3, width=60)
        self.prompt_text.grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        # 设置默认提示词为用户指定的内容
        self.prompt_text.insert(tk.END, "你是一名TEMU跨境运营，根据上传的图片，结合图片风格以及文字，给出中文标题描述，标题以图案结尾，并给出英文翻译。")
        
        # 处理按钮
        self.button_frame = ttk.Frame(self.control_frame)
        self.button_frame.grid(row=4, column=0, columnspan=4, pady=5, sticky="e")
        
        self.test_api_button = ttk.Button(self.button_frame, text="测试API连接", command=self.test_api_connection)
        self.test_api_button.pack(side=tk.LEFT, padx=5)
        
        self.process_button = ttk.Button(self.button_frame, text="开始处理", command=self.start_processing)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        # 创建中间区域，包含图片预览和结果显示
        self.middle_frame = ttk.Frame(self.main_frame)
        self.middle_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        self.main_frame.rowconfigure(1, weight=1)
        
        # 创建图片预览区域
        self.preview_frame = ttk.LabelFrame(self.middle_frame, text="图片预览")
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 使用PanedWindow来分割图片列表和图片显示区域
        self.preview_paned = ttk.PanedWindow(self.preview_frame, orient=tk.HORIZONTAL)
        self.preview_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建图片列表
        self.image_list_frame = ttk.Frame(self.preview_paned)
        self.preview_paned.add(self.image_list_frame, weight=1)
        
        self.image_listbox = tk.Listbox(self.image_list_frame)
        self.image_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.image_listbox.bind('<<ListboxSelect>>', self.show_selected_image)
        
        self.listbox_scrollbar = ttk.Scrollbar(self.image_list_frame, orient="vertical", command=self.image_listbox.yview)
        self.listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.image_listbox.config(yscrollcommand=self.listbox_scrollbar.set)
        
        # 创建图片显示区域
        self.image_display_frame = ttk.Frame(self.preview_paned)
        self.preview_paned.add(self.image_display_frame, weight=3)
        
        self.image_label = ttk.Label(self.image_display_frame)
        self.image_label.pack(fill=tk.BOTH, expand=True)
        
        # 创建结果显示区域
        self.result_frame = ttk.LabelFrame(self.main_frame, text="识别结果")
        self.result_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        self.main_frame.rowconfigure(2, weight=1)
        
        # 创建表格显示 - 简化字段，只显示文件名、中文标题和英文标题
        self.tree_frame = ttk.Frame(self.result_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 简化表格列，只保留必要字段
        self.tree_columns = ("文件名", "中文标题", "英文标题")
        self.result_tree = ttk.Treeview(self.tree_frame, columns=self.tree_columns, show="headings")
        
        # 设置列宽和标题
        for col in self.tree_columns:
            self.result_tree.heading(col, text=col)
            if col == "文件名":
                self.result_tree.column(col, width=200)
            else:
                self.result_tree.column(col, width=300)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.result_tree.yview)
        self.tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.config(yscrollcommand=self.tree_scrollbar.set)
        
        # 创建底部按钮区域 - 始终显示导出按钮
        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.grid(row=3, column=0, sticky="ew", pady=5)
        
        # 添加导出路径选择框
        self.export_path_label = ttk.Label(self.bottom_frame, text="导出路径:")
        self.export_path_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.export_path = tk.StringVar()
        self.export_path.set(os.path.expanduser("~/Desktop"))  # 默认导出到桌面
        self.export_path_entry = ttk.Entry(self.bottom_frame, textvariable=self.export_path)
        self.export_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.bottom_frame.columnconfigure(1, weight=1)
        
        self.export_path_button = ttk.Button(self.bottom_frame, text="选择...", command=self.browse_export_path)
        self.export_path_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 添加多个导出按钮，提供更多选择，并使其始终可见
        self.export_button_frame = ttk.Frame(self.bottom_frame)
        self.export_button_frame.grid(row=0, column=3, padx=5, pady=5)
        
        self.export_excel_button = ttk.Button(self.export_button_frame, text="导出Excel", command=self.export_excel)
        self.export_excel_button.pack(side=tk.LEFT, padx=5)
        
        self.export_csv_button = ttk.Button(self.export_button_frame, text="导出CSV", command=self.export_csv)
        self.export_csv_button.pack(side=tk.LEFT, padx=5)
        
        self.export_json_button = ttk.Button(self.export_button_frame, text="导出JSON", command=self.export_json)
        self.export_json_button.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, sticky="ew", pady=5)
        
        # 状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, anchor=tk.W)
        self.status_label.grid(row=5, column=0, sticky="ew")
        
        # 存储图片和结果
        self.image_files = []
        self.image_data = []
        self.current_photo = None
        
        # 调试模式
        self.debug_mode = True
        
        # API端点 - 修正为官方文档中的正确端点
        self.api_endpoint = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
        
        # 绑定窗口大小调整事件
        self.root.bind("<Configure>", self.on_window_resize)
    
    def on_window_resize(self, event):
        """处理窗口大小调整事件"""
        # 只处理根窗口的大小调整
        if event.widget == self.root:
            # 更新图片显示
            if hasattr(self, 'current_image_path') and self.current_image_path:
                self.show_selected_image(None)
    
    def browse_export_path(self):
        """选择导出路径"""
        path_selected = filedialog.askdirectory()
        if path_selected:
            self.export_path.set(path_selected)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.load_images_from_folder(folder_selected)
    
    def load_images_from_folder(self, folder_path):
        self.image_files = []
        self.image_listbox.delete(0, tk.END)
        
        try:
            for filename in os.listdir(folder_path):
                if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif')):
                    self.image_files.append(os.path.join(folder_path, filename))
                    self.image_listbox.insert(tk.END, filename)
            
            self.status_var.set(f"已加载 {len(self.image_files)} 张图片")
            
            # 如果有图片，显示第一张
            if self.image_files:
                self.image_listbox.selection_set(0)
                self.show_selected_image(None)
        except Exception as e:
            messagebox.showerror("错误", f"加载图片时出错: {str(e)}")
    
    def show_selected_image(self, event):
        try:
            # 获取选中的索引
            selected_indices = self.image_listbox.curselection()
            if selected_indices:
                index = selected_indices[0]
                image_path = self.image_files[index]
                self.current_image_path = image_path
                
                # 打开图片
                img = Image.open(image_path)
                
                # 获取图片显示区域的大小
                display_width = self.image_display_frame.winfo_width() - 20
                display_height = self.image_display_frame.winfo_height() - 20
                
                # 确保有最小显示尺寸
                display_width = max(display_width, 300)
                display_height = max(display_height, 300)
                
                # 调整图片大小以适应显示区域，保持纵横比
                img_width, img_height = img.size
                ratio = min(display_width / img_width, display_height / img_height)
                new_width = int(img_width * ratio)
                new_height = int(img_height * ratio)
                
                img = img.resize((new_width, new_height), Image.LANCZOS)
                
                # 转换为PhotoImage对象
                self.current_photo = ImageTk.PhotoImage(img)
                
                # 更新图片显示
                self.image_label.config(image=self.current_photo)
                
                # 显示图片信息
                width, height = img.size
                format_type = img.format
                self.status_var.set(f"图片: {os.path.basename(image_path)} | 尺寸: {width}x{height} | 格式: {format_type}")
        except Exception as e:
            messagebox.showerror("错误", f"显示图片时出错: {str(e)}")
    
    def get_full_model_id(self):
        """获取完整的模型ID，包括版本号（如果有）"""
        base_model_id = self.model_id.get().strip()
        version = self.version_var.get().strip()
        
        if version:
            # 检查基础模型ID是否已经包含版本号
            if version in base_model_id:
                return base_model_id
            else:
                return f"{base_model_id}-{version}"
        else:
            return base_model_id
    
    def test_api_connection(self):
        """测试API连接是否正常"""
        if not self.api_key.get().strip():
            messagebox.showwarning("警告", "请输入豆包API密钥")
            return
        
        # 获取完整模型ID
        model_id = self.get_full_model_id()
        
        # 显示测试中状态
        self.status_var.set("正在测试API连接...")
        self.root.update_idletasks()
        
        try:
            # 构建API请求头
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key.get().strip()}"
            }
            
            # 构建简单的请求数据，只包含文本
            payload = {
                "model": model_id,
                "messages": [
                    {
                        "role": "user",
                        "content": "测试连接"
                    }
                ]
            }
            
            # 发送请求到API端点
            response = requests.post(
                self.api_endpoint,
                headers=headers,
                json=payload,
                timeout=10  # 设置10秒超时
            )
            
            # 检查响应
            if response.status_code == 200:
                messagebox.showinfo("成功", f"API连接测试成功！\n模型ID: {model_id}\n响应状态码: {response.status_code}")
                self.status_var.set("API连接测试成功")
                
                # 打印响应内容用于调试
                if self.debug_mode:
                    print(f"API测试响应内容: {response.text[:500]}...")
            else:
                error_message = f"API连接测试失败: 状态码 {response.status_code}"
                try:
                    error_json = response.json()
                    if "error" in error_json:
                        error_message += f"\n错误信息: {error_json['error']['message']}"
                    else:
                        error_message += f"\n响应内容: {response.text}"
                except:
                    error_message += f"\n响应内容: {response.text}"
                
                messagebox.showerror("错误", error_message)
                self.status_var.set("API连接测试失败")
                
                # 提供模型ID建议
                if "model not found" in response.text.lower() or "not found" in response.text.lower():
                    self.suggest_model_id()
        except Exception as e:
            messagebox.showerror("错误", f"API连接测试出错: {str(e)}")
            self.status_var.set(f"API连接测试出错: {str(e)}")
    
    def suggest_model_id(self):
        """提供模型ID建议"""
        suggestion = "模型ID可能不正确或未开通，请尝试以下操作：\n\n"
        suggestion += "1. 登录火山引擎控制台，确认您已开通豆包API服务\n"
        suggestion += "2. 在控制台中查看已开通的模型列表和具体模型ID\n"
        suggestion += "3. 尝试以下常用模型ID：\n"
        suggestion += "   - doubao-1.5-vision-pro (不带版本号)\n"
        suggestion += "   - doubao-1.5-vision-pro-250328\n"
        suggestion += "   - doubao-1.5-vision-pro-250428\n"
        suggestion += "   - doubao-1.5-thinking-vision-pro-250428\n\n"
        suggestion += "注意：模型ID必须完全匹配，包括版本号（如有）"
        
        messagebox.showinfo("模型ID建议", suggestion)
    
    def start_processing(self):
        if not self.image_files:
            messagebox.showwarning("警告", "请先选择包含图片的文件夹")
            return
        
        if not self.api_key.get().strip():
            messagebox.showwarning("警告", "请输入豆包API密钥")
            return
        
        # 获取提示词
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt:
            prompt = "你是一名TEMU跨境运营，根据上传的图片，结合图片风格以及文字，给出中文标题描述，标题以图案结尾，并给出英文翻译。"
            self.prompt_text.delete("1.0", tk.END)
            self.prompt_text.insert(tk.END, prompt)
        
        # 清空之前的结果
        self.result_tree.delete(*self.result_tree.get_children())
        self.image_data = []
        
        # 启动处理线程
        threading.Thread(target=lambda: self.process_images(prompt), daemon=True).start()
    
    def process_images(self, prompt):
        try:
            total_images = len(self.image_files)
            self.progress_var.set(0)
            
            for i, image_path in enumerate(self.image_files):
                # 更新状态
                filename = os.path.basename(image_path)
                self.status_var.set(f"正在处理: {filename} ({i+1}/{total_images})")
                
                # 分析图片
                image_info = self.analyze_image(image_path, prompt)
                self.image_data.append(image_info)
                
                # 更新表格
                self.update_result_tree(image_info)
                
                # 更新进度条
                progress = (i + 1) / total_images * 100
                self.progress_var.set(progress)
                
                # 让UI有机会更新
                self.root.update_idletasks()
                
                # 添加短暂延迟，避免API限流
                time.sleep(0.5)
            
            self.status_var.set(f"处理完成，共 {total_images} 张图片")
            messagebox.showinfo("完成", f"已成功处理 {total_images} 张图片")
            
        except Exception as e:
            self.status_var.set(f"处理出错: {str(e)}")
            messagebox.showerror("错误", f"处理图片时出错: {str(e)}")
    
    def analyze_image(self, image_path, prompt):
        try:
            # 获取文件名（含后缀）
            filename = os.path.basename(image_path)
            
            # 调用豆包API进行图片识别，传入自定义提示词
            api_response = self.call_doubao_api(image_path, prompt)
            
            # 解析API响应，提取中文标题和英文翻译
            chinese_title, english_title = self.parse_api_response(api_response, filename)
            
            # 返回简化的图片信息，只包含文件名、中文标题和英文标题
            return {
                "文件名": filename,
                "中文标题": chinese_title,
                "英文标题": english_title,
                "API响应": api_response  # 保存原始响应用于调试
            }
        except Exception as e:
            print(f"处理图片 {image_path} 时出错: {str(e)}")
            return {
                "文件名": os.path.basename(image_path),
                "中文标题": "处理失败",
                "英文标题": "Processing Failed",
                "错误": str(e)
            }
    
    def prepare_image_base64(self, image_path):
        """
        准备图片的base64编码，确保格式正确
        修复Invalid base64 image_url错误
        """
        try:
            # 直接读取图片文件的二进制数据
            with open(image_path, "rb") as f:
                image_data = f.read()
            
            # 直接对原始二进制数据进行base64编码
            base64_data = base64.b64encode(image_data).decode('utf-8')
            
            # 确保没有换行符
            base64_data = base64_data.replace('\n', '').replace('\r', '')
            
            return base64_data
        except Exception as e:
            print(f"准备图片base64编码时出错: {str(e)}")
            raise e
    
    def call_doubao_api(self, image_path, prompt):
        """
        调用火山引擎豆包API进行图片识别，使用自定义提示词
        严格按照官方文档格式构建请求
        """
        try:
            # 获取API密钥和完整模型ID
            api_key = self.api_key.get().strip()
            model_id = self.get_full_model_id()
            
            # 准备图片base64编码，确保格式正确
            image_base64 = self.prepare_image_base64(image_path)
            
            # 构建API请求头
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
            
            # 构建请求数据，严格按照官方文档格式
            payload = {
                "model": model_id,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{image_base64}"
                                }
                            }
                        ]
                    }
                ],
                "temperature": 0.7,
                "top_p": 0.9,
                "stream": False
            }
            
            # 打印请求信息用于调试（不包含图片base64数据）
            if self.debug_mode:
                debug_payload = payload.copy()
                debug_payload["messages"][0]["content"][1]["image_url"]["url"] = "data:image/jpeg;base64,[BASE64_DATA]"
                print(f"API请求URL: {self.api_endpoint}")
                print(f"请求头: {headers}")
                print(f"请求体: {json.dumps(debug_payload, ensure_ascii=False, indent=2)}")
            
            # 发送请求到API端点
            response = requests.post(
                self.api_endpoint,
                headers=headers,
                json=payload,
                timeout=30  # 设置30秒超时
            )
            
            # 打印响应状态和内容用于调试
            if self.debug_mode:
                print(f"API响应状态码: {response.status_code}")
                print(f"API响应内容: {response.text[:500]}...")  # 只打印前500个字符
            
            # 解析响应
            if response.status_code == 200:
                result = response.json()
                # 根据API响应结构提取内容
                if "choices" in result and len(result["choices"]) > 0:
                    if "message" in result["choices"][0] and "content" in result["choices"][0]["message"]:
                        description = result["choices"][0]["message"]["content"]
                        return description
                    else:
                        return "API响应格式异常: 无法找到content字段"
                else:
                    return "API响应格式异常: 无法找到choices字段"
            else:
                # 增强错误处理，提供更详细的错误信息
                error_message = f"API调用失败: 状态码 {response.status_code}"
                try:
                    error_json = response.json()
                    if "error" in error_json:
                        error_message += f" - {error_json['error']['message']}"
                    else:
                        error_message += f" - {response.text}"
                except:
                    error_message += f" - {response.text}"
                
                print(error_message)
                
                # 如果是模型ID问题，提供更具体的建议
                if response.status_code == 400 and ("model not found" in response.text.lower() or "not found" in response.text.lower()):
                    error_message += "\n\n可能的原因：模型ID不正确或未开通。请在火山引擎控制台确认已开通该模型，并检查模型ID是否正确。"
                    # 在UI线程中显示模型ID建议
                    self.root.after(0, self.suggest_model_id)
                
                return f"API调用失败: {error_message}"
                
        except Exception as e:
            error_message = f"API调用出错: {str(e)}"
            print(error_message)
            return error_message
    
    def parse_api_response(self, api_response, filename):
        """
        从API响应中解析出中文标题和英文翻译
        增强解析逻辑，提高成功率
        """
        try:
            # 如果API调用失败，直接返回错误信息
            if api_response.startswith("API调用失败") or api_response.startswith("API调用出错") or api_response.startswith("API响应格式异常"):
                return f"{filename} - 处理失败", "Processing Failed"
            
            # 打印完整的API响应用于调试
            if self.debug_mode:
                print(f"解析API响应: {api_response}")
            
            # 尝试从API响应中提取中文标题和英文翻译
            lines = api_response.split('\n')
            
            chinese_title = ""
            english_title = ""
            
            # 查找包含"中文标题"或"标题"的行
            for line in lines:
                if "中文标题" in line or "标题" in line and "：" in line:
                    parts = line.split("：", 1)
                    if len(parts) > 1:
                        chinese_title = parts[1].strip()
                        break
            
            # 查找包含"英文翻译"或"英文标题"的行
            for line in lines:
                if "英文翻译" in line or "英文标题" in line and "：" in line:
                    parts = line.split("：", 1)
                    if len(parts) > 1:
                        english_title = parts[1].strip()
                        break
            
            # 如果没有找到明确的标题格式，尝试智能提取
            if not chinese_title and not english_title:
                # 尝试查找冒号分隔的内容
                for line in lines:
                    if ":" in line:
                        parts = line.split(":", 1)
                        key = parts[0].strip().lower()
                        value = parts[1].strip()
                        
                        if "中文" in key or "标题" in key:
                            chinese_title = value
                        elif "英文" in key or "translation" in key:
                            english_title = value
            
            # 如果仍然没有找到，尝试分离中英文内容
            if not chinese_title and not english_title:
                chinese_lines = []
                english_lines = []
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # 检查是否是中文内容（简单判断）
                    if any('\u4e00' <= char <= '\u9fff' for char in line):
                        chinese_lines.append(line)
                    # 检查是否是英文内容（简单判断）
                    elif all(ord(char) < 128 for char in line if char.strip()) and len(line.strip()) > 5:
                        english_lines.append(line)
                
                if chinese_lines:
                    chinese_title = chinese_lines[0]
                if english_lines:
                    english_title = english_lines[0]
            
            # 如果仍然没有找到，使用整个响应作为中文标题
            if not chinese_title:
                chinese_title = api_response[:50] + "..." if len(api_response) > 50 else api_response
            
            # 如果没有找到英文翻译，使用中文标题
            if not english_title:
                english_title = "Translation not provided"
            
            # 确保标题以"图案"结尾（如果不是错误信息）
            if not chinese_title.startswith("处理失败") and not chinese_title.endswith("图案"):
                chinese_title += "图案"
            
            # 如果标题太长，截断它
            if len(chinese_title) > 100:
                chinese_title = chinese_title[:97] + "..."
            if len(english_title) > 100:
                english_title = english_title[:97] + "..."
            
            return chinese_title, english_title
            
        except Exception as e:
            print(f"解析API响应出错: {str(e)}")
            return f"{filename} - 解析失败", "Parsing Failed"
    
    def update_result_tree(self, image_info):
        """
        更新结果表格
        """
        # 在UI线程中更新表格
        self.root.after(0, lambda: self._update_tree(image_info))
    
    def _update_tree(self, image_info):
        values = [image_info.get(col, "") for col in self.tree_columns]
        self.result_tree.insert("", tk.END, values=values)
    
    def export_excel(self):
        """
        导出结果到Excel文件
        """
        self._export_data("xlsx", "Excel 文件", "*.xlsx")
    
    def export_csv(self):
        """
        导出结果到CSV文件
        """
        self._export_data("csv", "CSV 文件", "*.csv")
    
    def export_json(self):
        """
        导出结果到JSON文件
        """
        if not self.image_data:
            messagebox.showwarning("警告", "没有可导出的数据")
            return
        
        try:
            # 使用用户选择的导出路径
            export_dir = self.export_path.get()
            if not os.path.exists(export_dir):
                os.makedirs(export_dir)
            
            # 生成文件名
            file_name = f"图片识别结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            file_path = os.path.join(export_dir, file_name)
            
            # 导出到JSON
            with open(file_path, 'w', encoding='utf-8') as f:
                # 只导出简化字段
                simplified_data = []
                for item in self.image_data:
                    simplified_data.append({
                        "文件名": item.get("文件名", ""),
                        "中文标题": item.get("中文标题", ""),
                        "英文标题": item.get("英文标题", "")
                    })
                json.dump(simplified_data, f, ensure_ascii=False, indent=2)
            
            self.status_var.set(f"已成功导出到 {file_path}")
            messagebox.showinfo("成功", f"已成功导出到 {file_path}")
            
            # 打开文件所在文件夹
            self._open_file_location(file_path)
            
        except Exception as e:
            self.status_var.set(f"导出出错: {str(e)}")
            messagebox.showerror("错误", f"导出JSON时出错: {str(e)}")
    
    def _export_data(self, format_type, file_type_desc, file_extension):
        """
        通用导出函数
        """
        if not self.image_data:
            messagebox.showwarning("警告", "没有可导出的数据")
            return
        
        try:
            # 使用用户选择的导出路径
            export_dir = self.export_path.get()
            if not os.path.exists(export_dir):
                os.makedirs(export_dir)
            
            # 生成文件名
            file_name = f"图片识别结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{format_type}"
            file_path = os.path.join(export_dir, file_name)
            
            # 创建DataFrame - 只包含简化字段
            simplified_data = []
            for item in self.image_data:
                simplified_data.append({
                    "文件名": item.get("文件名", ""),
                    "中文标题": item.get("中文标题", ""),
                    "英文标题": item.get("英文标题", "")
                })
            
            df = pd.DataFrame(simplified_data)
            
            # 根据格式导出
            if format_type == "xlsx":
                df.to_excel(file_path, index=False)
            elif format_type == "csv":
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            
            self.status_var.set(f"已成功导出到 {file_path}")
            messagebox.showinfo("成功", f"已成功导出到 {file_path}")
            
            # 打开文件所在文件夹
            self._open_file_location(file_path)
            
        except Exception as e:
            self.status_var.set(f"导出出错: {str(e)}")
            messagebox.showerror("错误", f"导出{file_type_desc}时出错: {str(e)}")
    
    def _open_file_location(self, file_path):
        """
        打开文件所在位置
        """
        try:
            if sys.platform == 'win32':
                os.startfile(os.path.dirname(file_path))
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{os.path.dirname(file_path)}"')
            else:  # Linux
                os.system(f'xdg-open "{os.path.dirname(file_path)}"')
        except Exception as e:
            print(f"打开文件位置失败: {str(e)}")

def main():
    root = tk.Tk()
    app = ImageRecognitionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
