import tkinter as tk
from tkinter import filedialog, ttk, messagebox, colorchooser, simpledialog
import pandas as pd
import qrcode
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import threading
from datetime import datetime
import json

class LabelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("高级标签生成器")
        self.root.geometry("1300x900")
        self.root.configure(bg="#f5f5f5")
        self.root.minsize(1000, 700)
        
        # 加载配置
        self.load_config()
        
        # 数据存储
        self.df = None
        self.label_width = self.config.get('label_width', 300)
        self.label_height = self.config.get('label_height', 400)
        self.qr_size = self.config.get('qr_size', 150)
        self.field_display_types = {}
        self.field_prefixes = {}
        self.field_suffixes = {}
        self.field_font_sizes = {}
        self.field_colors = {}
        self.field_order = []
        self.output_dir = self.config.get('output_dir', os.getcwd())
        self.bg_color = self.config.get('bg_color', '#FFFFFF')
        self.text_color = self.config.get('text_color', '#000000')
        self.qr_color = self.config.get('qr_color', '#000000')
        self.preview_row = 0
        self.total_rows = 0
        self.custom_fields = {}  # 存储自定义字段内容
        
        # 创建菜单
        self.create_menu()
        
        # GUI布局
        self.setup_ui()
        
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入 Excel", command=self.import_excel)
        file_menu.add_command(label="导出配置", command=self.export_config)
        file_menu.add_command(label="导入配置", command=self.import_config)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="设置输出目录", command=self.set_output_dir)
        menubar.add_cascade(label="编辑", menu=edit_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def load_config(self):
        self.config = {}
        try:
            if os.path.exists("label_config.json"):
                with open("label_config.json", "r") as f:
                    self.config = json.load(f)
        except:
            pass
    
    def save_config(self):
        config = {
            'label_width': self.label_width,
            'label_height': self.label_height,
            'qr_size': self.qr_size,
            'output_dir': self.output_dir,
            'bg_color': self.bg_color,
            'text_color': self.text_color,
            'qr_color': self.qr_color,
            'field_order': self.field_order,
            'field_display_types': {k: v.get() for k, v in self.field_display_types.items()},
            'field_prefixes': {k: v.get() for k, v in self.field_prefixes.items()},
            'field_suffixes': {k: v.get() for k, v in self.field_suffixes.items()},
            'field_font_sizes': {k: v.get() for k, v in self.field_font_sizes.items()},
            'field_colors': {k: v for k, v in self.field_colors.items()},
            'custom_fields': self.custom_fields
        }
        
        try:
            with open("label_config.json", "w") as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            messagebox.showerror("保存配置失败", f"错误: {str(e)}")
    
    def setup_ui(self):
        # 设置主题
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=6, font=("Arial", 10), background="#4CAF50", foreground="white")
        style.configure("TLabel", font=("Arial", 10), background="#f5f5f5")
        style.configure("TEntry", font=("Arial", 10))
        style.configure("TCombobox", font=("Arial", 10))
        style.configure("Header.TFrame", background="#e0e0e0")
        style.configure("Section.TLabel", font=("Arial", 11, "bold"), background="#e0e0e0")
        
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 左侧配置区域
        config_frame = ttk.Frame(main_frame, width=700)
        config_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # 顶部框架：Excel导入
        top_frame = ttk.Frame(config_frame, padding=10, relief="groove", style="Header.TFrame")
        top_frame.pack(fill="x", pady=5)
        
        ttk.Label(top_frame, text="Excel 文件:", style="Section.TLabel").pack(side="left")
        self.file_label = ttk.Label(top_frame, text="未选择文件", font=("Arial", 9))
        self.file_label.pack(side="left", padx=10)
        
        ttk.Button(top_frame, text="导入 Excel", command=self.import_excel).pack(side="right", padx=5)
        
        # 字段配置区域
        field_config_frame = ttk.Frame(config_frame, padding=10, relief="groove")
        field_config_frame.pack(fill="both", expand=True, pady=5)
        
        ttk.Label(field_config_frame, text="字段配置", style="Section.TLabel").pack(anchor="w", pady=(0, 10))
        
        # 字段列表与操作按钮
        list_frame = ttk.Frame(field_config_frame)
        list_frame.pack(fill="x", pady=5)
        
        ttk.Label(list_frame, text="字段顺序:").pack(side="left", padx=(0, 10))
        
        # 添加字段按钮
        add_field_frame = ttk.Frame(list_frame)
        add_field_frame.pack(side="right", padx=5)
        ttk.Button(add_field_frame, text="+ 添加字段", command=self.add_custom_field, width=10).pack(side="right", padx=2)
        
        # 字段列表框
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(side="left", fill="x", expand=True)
        
        self.content_listbox = tk.Listbox(listbox_frame, height=5, width=30, font=("Arial", 10), selectmode=tk.SINGLE)
        self.content_listbox.pack(side="left", fill="x", expand=True, padx=5)
        
        # 列表框滚动条
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.content_listbox.yview)
        listbox_scrollbar.pack(side="right", fill="y")
        self.content_listbox.config(yscrollcommand=listbox_scrollbar.set)
        
        # 字段操作按钮
        sort_frame = ttk.Frame(list_frame)
        sort_frame.pack(side="left", padx=5)
        ttk.Button(sort_frame, text="↑ 上移", command=self.move_up, width=8).pack(pady=2, fill="x")
        ttk.Button(sort_frame, text="↓ 下移", command=self.move_down, width=8).pack(pady=2, fill="x")
        ttk.Button(sort_frame, text="置顶", command=self.move_top, width=8).pack(pady=2, fill="x")
        ttk.Button(sort_frame, text="置底", command=self.move_bottom, width=8).pack(pady=2, fill="x")
        ttk.Button(sort_frame, text="删除", command=self.remove_field, width=8).pack(pady=2, fill="x")
        
        # 字段配置滚动区域 - 修复滚动条问题
        field_scroll_container = ttk.Frame(field_config_frame)
        field_scroll_container.pack(fill="both", expand=True)
        
        self.field_canvas = tk.Canvas(field_scroll_container, bg="#f5f5f5", highlightthickness=0)
        self.field_scrollbar = ttk.Scrollbar(field_scroll_container, orient="vertical", command=self.field_canvas.yview)
        self.field_scrollable_frame = ttk.Frame(self.field_canvas)
        
        self.field_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.field_canvas.configure(scrollregion=self.field_canvas.bbox("all"))
        )
        
        self.field_canvas.create_window((0, 0), window=self.field_scrollable_frame, anchor="nw")
        self.field_canvas.configure(yscrollcommand=self.field_scrollbar.set)
        
        self.field_canvas.pack(side="left", fill="both", expand=True)
        self.field_scrollbar.pack(side="right", fill="y")
        
        # 标签配置区域
        label_config_frame = ttk.Frame(config_frame, padding=10, relief="groove")
        label_config_frame.pack(fill="x", pady=5)
        
        ttk.Label(label_config_frame, text="标签设置", style="Section.TLabel").pack(anchor="w", pady=(0, 10))
        
        # 尺寸配置
        size_frame = ttk.Frame(label_config_frame)
        size_frame.pack(fill="x", pady=5)
        
        ttk.Label(size_frame, text="标签宽度:").grid(row=0, column=0, padx=5, sticky="e")
        self.width_entry = ttk.Entry(size_frame, width=8)
        self.width_entry.insert(0, str(self.label_width))
        self.width_entry.grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(size_frame, text="标签高度:").grid(row=0, column=2, padx=5, sticky="e")
        self.height_entry = ttk.Entry(size_frame, width=8)
        self.height_entry.insert(0, str(self.label_height))
        self.height_entry.grid(row=0, column=3, padx=5, sticky="w")
        
        ttk.Label(size_frame, text="二维码尺寸:").grid(row=0, column=4, padx=5, sticky="e")
        self.qr_size_entry = ttk.Entry(size_frame, width=8)
        self.qr_size_entry.insert(0, str(self.qr_size))
        self.qr_size_entry.grid(row=0, column=5, padx=5, sticky="w")
        
        ttk.Button(size_frame, text="更新尺寸", command=self.update_size).grid(row=0, column=6, padx=10)
        
        # 颜色配置
        color_frame = ttk.Frame(label_config_frame)
        color_frame.pack(fill="x", pady=5)
        
        ttk.Label(color_frame, text="背景颜色:").grid(row=0, column=0, padx=5, sticky="e")
        self.bg_color_btn = ttk.Button(color_frame, text="选择", width=8, 
                                     command=lambda: self.choose_color("bg_color"))
        self.bg_color_btn.grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(color_frame, text="文字颜色:").grid(row=0, column=2, padx=5, sticky="e")
        self.text_color_btn = ttk.Button(color_frame, text="选择", width=8, 
                                       command=lambda: self.choose_color("text_color"))
        self.text_color_btn.grid(row=0, column=3, padx=5, sticky="w")
        
        ttk.Label(color_frame, text="二维码颜色:").grid(row=0, column=4, padx=5, sticky="e")
        self.qr_color_btn = ttk.Button(color_frame, text="选择", width=8, 
                                     command=lambda: self.choose_color("qr_color"))
        self.qr_color_btn.grid(row=0, column=5, padx=5, sticky="w")
        
        # 输出目录
        output_frame = ttk.Frame(label_config_frame)
        output_frame.pack(fill="x", pady=5)
        
        ttk.Label(output_frame, text="输出目录:").pack(side="left")
        self.output_dir_label = ttk.Label(output_frame, text=self.output_dir, font=("Arial", 9), width=50)
        self.output_dir_label.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(output_frame, text="浏览", command=self.set_output_dir).pack(side="left", padx=5)
        
        # 预览行选择
        preview_frame = ttk.Frame(label_config_frame)
        preview_frame.pack(fill="x", pady=5)
        
        ttk.Label(preview_frame, text="预览行:").pack(side="left")
        self.preview_spin = ttk.Spinbox(preview_frame, from_=1, to=100, width=8)
        self.preview_spin.pack(side="left", padx=5)
        self.preview_spin.set(1)
        self.preview_spin.bind("<KeyRelease>", self.update_preview)
        self.preview_spin.bind("<<Increment>>", self.update_preview)
        self.preview_spin.bind("<<Decrement>>", self.update_preview)
        
        ttk.Label(preview_frame, text="总行数:").pack(side="left", padx=(20, 5))
        self.total_rows_label = ttk.Label(preview_frame, text="0")
        self.total_rows_label.pack(side="left")
        
        # 进度条
        self.progress = ttk.Progressbar(label_config_frame, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(fill="x", pady=10)
        
        # 生成按钮
        btn_frame = ttk.Frame(config_frame)
        btn_frame.pack(fill="x", pady=10)
        ttk.Button(btn_frame, text="生成标签", command=self.start_generation).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config).pack(side="left", padx=5)
        
        # 右侧预览区域
        preview_container = ttk.Frame(main_frame, width=500)
        preview_container.pack(side="right", fill="both", expand=True)
        
        preview_header = ttk.Frame(preview_container, padding=10, relief="groove", style="Header.TFrame")
        preview_header.pack(fill="x")
        ttk.Label(preview_header, text="预览", style="Section.TLabel").pack()
        
        # 修复预览区域滚动条问题
        preview_scroll_container = ttk.Frame(preview_container)
        preview_scroll_container.pack(fill="both", expand=True)
        
        self.preview_canvas = tk.Canvas(preview_scroll_container, bg="#f5f5f5", highlightthickness=0)
        self.preview_scrollbar = ttk.Scrollbar(preview_scroll_container, orient="vertical", command=self.preview_canvas.yview)
        self.preview_inner_frame = ttk.Frame(self.preview_canvas)
        
        self.preview_inner_frame.bind(
            "<Configure>",
            lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        )
        
        self.preview_canvas.create_window((0, 0), window=self.preview_inner_frame, anchor="nw")
        self.preview_canvas.configure(yscrollcommand=self.preview_scrollbar.set)
        
        self.preview_canvas.pack(side="left", fill="both", expand=True)
        self.preview_scrollbar.pack(side="right", fill="y")
        
        self.preview_label = ttk.Label(self.preview_inner_frame)
        self.preview_label.pack(pady=20, padx=20)
        
        # 状态栏
        self.status_bar = ttk.Frame(self.root, height=25, relief="sunken")
        self.status_bar.pack(fill="x", side="bottom")
        self.status_label = ttk.Label(self.status_bar, text="就绪", font=("Arial", 9))
        self.status_label.pack(side="left", padx=10)
        
        # 更新颜色按钮背景
        self.update_color_buttons()
        
    def update_color_buttons(self):
        # 创建自定义样式
        style = ttk.Style()
        style.configure("Color.TButton", background=self.bg_color)
        style.configure("TextColor.TButton", background=self.text_color)
        style.configure("QRColor.TButton", background=self.qr_color)
        
        self.bg_color_btn.configure(style="Color.TButton")
        self.text_color_btn.configure(style="TextColor.TButton")
        self.qr_color_btn.configure(style="QRColor.TButton")
        
    def choose_color(self, color_type):
        color = colorchooser.askcolor(title=f"选择{color_type}颜色")[1]
        if color:
            if color_type == "bg_color":
                self.bg_color = color
            elif color_type == "text_color":
                self.text_color = color
            elif color_type == "qr_color":
                self.qr_color = color
            
            self.update_color_buttons()
            self.update_preview()
    
    def set_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir = dir_path
            self.output_dir_label.config(text=dir_path)
            self.update_status(f"输出目录设置为: {dir_path}")
    
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.df = pd.read_csv(file_path)
                else:
                    self.df = pd.read_excel(file_path)
                
                self.file_label.config(text=os.path.basename(file_path))
                self.total_rows = len(self.df)
                self.total_rows_label.config(text=str(self.total_rows))
                self.preview_spin.config(from_=1, to=self.total_rows)
                self.preview_row = 0
                
                # 清空旧内容
                for widget in self.field_scrollable_frame.winfo_children():
                    widget.destroy()
                
                # 重置数据
                self.field_order = list(self.df.columns)
                self.field_display_types = {}
                self.field_prefixes = {}
                self.field_suffixes = {}
                self.field_font_sizes = {}
                self.field_colors = {}
                
                # 尝试加载字段配置
                self.load_field_config()
                
                # 字段配置
                self._render_field_config()
                
                self.update_status(f"成功导入文件: {os.path.basename(file_path)}, 共 {len(self.df)} 行数据")
                self.update_preview()
            except Exception as e:
                messagebox.showerror("导入错误", f"导入失败：{str(e)}")
                self.update_status(f"导入失败: {str(e)}")
    
    def load_field_config(self):
        # 如果配置中有字段设置，则加载
        if 'field_order' in self.config:
            # 只保留当前数据集中存在的字段
            self.field_order = [col for col in self.config['field_order'] if col in self.df.columns or col in self.custom_fields]
            # 添加新字段到末尾
            for col in self.df.columns:
                if col not in self.field_order:
                    self.field_order.append(col)
            
            # 添加自定义字段
            if 'custom_fields' in self.config:
                self.custom_fields = self.config['custom_fields']
                for col in self.custom_fields:
                    if col not in self.field_order:
                        self.field_order.append(col)
        
        # 加载字段显示类型
        if 'field_display_types' in self.config:
            for col, display_type in self.config['field_display_types'].items():
                if col in self.df.columns or col in self.custom_fields:
                    self.field_display_types[col] = tk.StringVar(value=display_type)
        
        # 加载前后缀和字体大小
        for field_dict, config_key in [(self.field_prefixes, 'field_prefixes'), 
                                      (self.field_suffixes, 'field_suffixes'),
                                      (self.field_font_sizes, 'field_font_sizes')]:
            if config_key in self.config:
                for col, value in self.config[config_key].items():
                    if col in self.df.columns or col in self.custom_fields:
                        if config_key == 'field_font_sizes':
                            self.field_font_sizes[col] = value
                        else:
                            # 对于前缀和后缀，创建Entry控件
                            entry = ttk.Entry()
                            entry.insert(0, value)
                            if config_key == 'field_prefixes':
                                self.field_prefixes[col] = entry
                            else:
                                self.field_suffixes[col] = entry
        
        # 加载字段颜色
        if 'field_colors' in self.config:
            for col, color in self.config['field_colors'].items():
                if col in self.df.columns or col in self.custom_fields:
                    self.field_colors[col] = color
    
    def _render_field_config(self):
        # 填充字段列表
        self.content_listbox.delete(0, tk.END)
        for col in self.field_order:
            self.content_listbox.insert(tk.END, col)
        
        # 为每个字段创建配置UI
        for col in self.field_order:
            frame = ttk.Frame(self.field_scrollable_frame, padding=5)
            frame.pack(fill="x", pady=3)
            
            # 字段名称
            is_custom = col in self.custom_fields
            label_text = f"{col} (自定义)" if is_custom else f"{col}:"
            ttk.Label(frame, text=label_text, width=15, anchor="e").pack(side="left")
            
            # 展示形式
            var = self.field_display_types.get(col, tk.StringVar(value="text"))
            self.field_display_types[col] = var
            rb_frame = ttk.Frame(frame)
            rb_frame.pack(side="left", padx=5)
            ttk.Radiobutton(rb_frame, text="文本", value="text", variable=var, 
                           command=self.update_preview).pack(side="left")
            ttk.Radiobutton(rb_frame, text="二维码", value="qrcode", variable=var, 
                           command=self.update_preview).pack(side="left", padx=(10, 0))
            
            # 前置内容
            ttk.Label(frame, text="前缀:").pack(side="left", padx=(10, 0))
            prefix_entry = self.field_prefixes.get(col, ttk.Entry(frame, width=8))
            if not isinstance(prefix_entry, ttk.Entry):
                entry = ttk.Entry(frame, width=8)
                entry.insert(0, prefix_entry)
                prefix_entry = entry
            prefix_entry.pack(side="left", padx=2)
            self.field_prefixes[col] = prefix_entry
            
            # 后置内容
            ttk.Label(frame, text="后缀:").pack(side="left", padx=(10, 0))
            suffix_entry = self.field_suffixes.get(col, ttk.Entry(frame, width=8))
            if not isinstance(suffix_entry, ttk.Entry):
                entry = ttk.Entry(frame, width=8)
                entry.insert(0, suffix_entry)
                suffix_entry = entry
            suffix_entry.pack(side="left", padx=2)
            self.field_suffixes[col] = suffix_entry
            
            # 自定义字段内容输入框
            if is_custom:
                ttk.Label(frame, text="内容:").pack(side="left", padx=(10, 0))
                content_entry = ttk.Entry(frame, width=15)
                content_entry.insert(0, self.custom_fields.get(col, ""))
                content_entry.pack(side="left", padx=2)
                content_entry.bind("<KeyRelease>", lambda e, c=col: self.update_custom_field(c, e.widget.get()))
            
            # 字体大小
            ttk.Label(frame, text="字体:").pack(side="left", padx=(10, 0))
            font_size_value = self.field_font_sizes.get(col, "16")
            font_size_combo = ttk.Combobox(frame, values=[8, 10, 12, 14, 16, 18, 20, 24, 28, 32], width=4, state="readonly")
            font_size_combo.set(str(font_size_value))
            font_size_combo.pack(side="left", padx=2)
            self.field_font_sizes[col] = font_size_combo
            font_size_combo.bind("<<ComboboxSelected>>", lambda e, col=col: self.update_preview())
            
            # 字体颜色
            color_btn = ttk.Button(frame, text="颜色", width=6, 
                                 command=lambda c=col: self.choose_field_color(c))
            color_btn.pack(side="left", padx=(10, 0))
            
            # 设置默认颜色
            if col not in self.field_colors:
                self.field_colors[col] = self.text_color
            
            # 绑定事件
            prefix_entry.bind("<KeyRelease>", lambda e, col=col: self.update_preview())
            suffix_entry.bind("<KeyRelease>", lambda e, col=col: self.update_preview())
            
        # 更新滚动区域
        self.field_canvas.update_idletasks()
        self.field_canvas.config(scrollregion=self.field_canvas.bbox("all"))
    
    def update_custom_field(self, field_name, content):
        """更新自定义字段的内容"""
        self.custom_fields[field_name] = content
        self.update_preview()
    
    def choose_field_color(self, field_name):
        color = colorchooser.askcolor(title=f"选择 {field_name} 颜色")[1]
        if color:
            self.field_colors[field_name] = color
            self.update_preview()
    
    def move_up(self):
        try:
            index = self.content_listbox.curselection()[0]
            if index > 0:
                self.field_order[index], self.field_order[index - 1] = self.field_order[index - 1], self.field_order[index]
                self.update_field_list()
                self.content_listbox.select_set(index - 1)
                self.update_preview()
        except IndexError:
            pass
    
    def move_down(self):
        try:
            index = self.content_listbox.curselection()[0]
            if index < len(self.field_order) - 1:
                self.field_order[index], self.field_order[index + 1] = self.field_order[index + 1], self.field_order[index]
                self.update_field_list()
                self.content_listbox.select_set(index + 1)
                self.update_preview()
        except IndexError:
            pass
    
    def move_top(self):
        try:
            index = self.content_listbox.curselection()[0]
            if index > 0:
                item = self.field_order.pop(index)
                self.field_order.insert(0, item)
                self.update_field_list()
                self.content_listbox.select_set(0)
                self.update_preview()
        except IndexError:
            pass
    
    def move_bottom(self):
        try:
            index = self.content_listbox.curselection()[0]
            if index < len(self.field_order) - 1:
                item = self.field_order.pop(index)
                self.field_order.append(item)
                self.update_field_list()
                self.content_listbox.select_set(len(self.field_order) - 1)
                self.update_preview()
        except IndexError:
            pass
    
    def remove_field(self):
        """删除选中的字段"""
        try:
            index = self.content_listbox.curselection()[0]
            field_name = self.field_order[index]
            
            # 如果是自定义字段，从custom_fields中删除
            if field_name in self.custom_fields:
                del self.custom_fields[field_name]
            
            # 从字段列表中删除
            self.field_order.pop(index)
            
            # 更新字段列表
            self.update_field_list()
            
            # 重新渲染字段配置
            for widget in self.field_scrollable_frame.winfo_children():
                widget.destroy()
            self._render_field_config()
            
            self.update_status(f"已删除字段: {field_name}")
            self.update_preview()
        except IndexError:
            messagebox.showwarning("删除字段", "请先选择一个字段")
    
    def add_custom_field(self):
        """添加新的自定义字段"""
        field_name = simpledialog.askstring("添加自定义字段", "请输入字段名称:", parent=self.root)
        if field_name:
            if field_name in self.field_order:
                messagebox.showwarning("添加字段", f"字段 '{field_name}' 已存在!")
                return
            
            # 添加到字段列表
            self.field_order.append(field_name)
            self.custom_fields[field_name] = ""  # 默认内容为空
            
            # 更新字段列表
            self.update_field_list()
            
            # 重新渲染字段配置
            for widget in self.field_scrollable_frame.winfo_children():
                widget.destroy()
            self._render_field_config()
            
            self.update_status(f"已添加自定义字段: {field_name}")
            self.update_preview()
    
    def update_field_list(self):
        self.content_listbox.delete(0, tk.END)
        for col in self.field_order:
            self.content_listbox.insert(tk.END, col)
    
    def update_size(self):
        try:
            self.label_width = int(self.width_entry.get())
            self.label_height = int(self.height_entry.get())
            self.qr_size = int(self.qr_size_entry.get())
            self.update_status("尺寸已更新")
            self.update_preview()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
    
    def start_generation(self):
        if self.df is None:
            messagebox.showerror("错误", "请先导入Excel文件！")
            return
        
        # 禁用生成按钮
        self.progress['value'] = 0
        self.update_status("正在生成标签...")
        
        # 使用线程生成标签，避免界面冻结
        threading.Thread(target=self.generate_labels, daemon=True).start()
    
    def generate_labels(self):
        total = len(self.df)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder = os.path.join(self.output_dir, f"labels_{timestamp}")
        os.makedirs(output_folder, exist_ok=True)
        
        for idx, row in self.df.iterrows():
            try:
                img = Image.new('RGB', (self.label_width, self.label_height), color=self.bg_color)
                draw = ImageDraw.Draw(img)
                current_y = 20
                
                for col in self.field_order:
                    # 获取字段内容
                    if col in self.custom_fields:
                        content = self.custom_fields[col]  # 自定义字段内容
                    else:
                        content = str(row[col])  # Excel数据内容
                    
                    prefix = self.field_prefixes[col].get()
                    suffix = self.field_suffixes[col].get()
                    full_content = f"{prefix}{content}{suffix}"
                    
                    if self.field_display_types[col].get() == "qrcode":
                        qr = qrcode.QRCode(version=1, box_size=5, border=2)
                        qr.add_data(full_content)
                        qr.make(fit=True)
                        
                        # 创建彩色二维码
                        qr_img = qr.make_image(fill_color=self.qr_color, back_color=self.bg_color)
                        qr_img = qr_img.resize((self.qr_size, self.qr_size))
                        x = (self.label_width - self.qr_size) // 2
                        img.paste(qr_img, (x, current_y))
                        current_y += self.qr_size + 20
                    else:
                        font_size = int(self.field_font_sizes[col].get())
                        try:
                            font = ImageFont.truetype("msyh.ttc", font_size)
                        except:
                            font = ImageFont.load_default()
                        
                        text_color = self.field_colors.get(col, self.text_color)
                        text_width = draw.textlength(full_content, font=font)
                        x = (self.label_width - text_width) // 2
                        draw.text((x, current_y), full_content, fill=text_color, font=font)
                        current_y += font_size + 10
                
                # 保存标签
                img.save(os.path.join(output_folder, f"label_{idx+1}.png"))
                
                # 更新进度
                progress = (idx + 1) / total * 100
                self.root.after(10, lambda v=progress: self.progress.configure(value=v))
                
            except Exception as e:
                self.root.after(10, lambda e=e: self.update_status(f"生成第 {idx+1} 行时出错: {str(e)}"))
        
        self.root.after(10, lambda: self.update_status(f"成功生成 {total} 个标签到: {output_folder}"))
        self.root.after(10, lambda: messagebox.showinfo("完成", f"已生成 {total} 个标签到:\n{output_folder}"))
    
    def update_preview(self, event=None):
        if self.df is None or not self.field_order:
            return
        
        try:
            row_idx = int(self.preview_spin.get()) - 1
            if row_idx < 0 or row_idx >= len(self.df):
                row_idx = 0
                self.preview_spin.set(1)
            
            self.preview_row = row_idx
            sample_row = self.df.iloc[row_idx] if self.df is not None else {}
            
            img = Image.new('RGB', (self.label_width, self.label_height), color=self.bg_color)
            draw = ImageDraw.Draw(img)
            current_y = 20
            
            for col in self.field_order:
                # 获取字段内容
                if col in self.custom_fields:
                    content = self.custom_fields[col]  # 自定义字段内容
                else:
                    content = str(sample_row.get(col, ""))  # Excel数据内容
                
                prefix = self.field_prefixes[col].get()
                suffix = self.field_suffixes[col].get()
                full_content = f"{prefix}{content}{suffix}"
                
                if self.field_display_types.get(col, tk.StringVar(value="text")).get() == "qrcode":
                    qr = qrcode.QRCode(version=1, box_size=5, border=2)
                    qr.add_data(full_content)
                    qr.make(fit=True)
                    qr_img = qr.make_image(fill_color=self.qr_color, back_color=self.bg_color)
                    qr_img = qr_img.resize((self.qr_size, self.qr_size))
                    x = (self.label_width - self.qr_size) // 2
                    img.paste(qr_img, (x, current_y))
                    current_y += self.qr_size + 20
                else:
                    try:
                        font_size = int(self.field_font_sizes[col].get())
                        font = ImageFont.truetype("msyh.ttc", font_size)
                    except:
                        font = ImageFont.load_default()
                    
                    text_color = self.field_colors.get(col, self.text_color)
                    text_width = draw.textlength(full_content, font=font)
                    x = (self.label_width - text_width) // 2
                    draw.text((x, current_y), full_content, fill=text_color, font=font)
                    current_y += font_size + 10
            
            # 添加边框
            border_img = Image.new('RGB', (self.label_width + 20, self.label_height + 20), color="#f0f0f0")
            border_img.paste(img, (10, 10))
            
            preview_img = ImageTk.PhotoImage(border_img)
            self.preview_label.config(image=preview_img)
            self.preview_label.image = preview_img
            
            self.update_status(f"预览第 {row_idx+1} 行")
            
            # 更新预览区域滚动
            self.preview_canvas.update_idletasks()
            self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))
            
        except Exception as e:
            self.update_status(f"预览错误: {str(e)}")
    
    def export_config(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            self.save_config()
            try:
                with open("label_config.json", "r") as src, open(file_path, "w") as dst:
                    dst.write(src.read())
                self.update_status(f"配置已导出到: {file_path}")
            except Exception as e:
                messagebox.showerror("导出错误", f"导出失败: {str(e)}")
    
    def import_config(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "r") as f:
                    self.config = json.load(f)
                
                # 应用配置
                self.label_width = self.config.get('label_width', 300)
                self.label_height = self.config.get('label_height', 400)
                self.qr_size = self.config.get('qr_size', 150)
                self.output_dir = self.config.get('output_dir', os.getcwd())
                self.bg_color = self.config.get('bg_color', '#FFFFFF')
                self.text_color = self.config.get('text_color', '#000000')
                self.qr_color = self.config.get('qr_color', '#000000')
                self.custom_fields = self.config.get('custom_fields', {})
                
                # 更新UI
                self.width_entry.delete(0, tk.END)
                self.width_entry.insert(0, str(self.label_width))
                self.height_entry.delete(0, tk.END)
                self.height_entry.insert(0, str(self.label_height))
                self.qr_size_entry.delete(0, tk.END)
                self.qr_size_entry.insert(0, str(self.qr_size))
                self.output_dir_label.config(text=self.output_dir)
                self.update_color_buttons()
                
                # 如果有数据，重新渲染字段配置
                if self.df is not None:
                    self.load_field_config()
                    self._render_field_config()
                    self.update_preview()
                
                self.update_status(f"配置已导入: {os.path.basename(file_path)}")
                messagebox.showinfo("成功", "配置导入成功！")
                
            except Exception as e:
                messagebox.showerror("导入错误", f"导入失败: {str(e)}")
    
    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于标签生成器")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        
        ttk.Label(about_window, text="标签生成器", font=("Arial", 16, "bold")).pack(pady=20)
        ttk.Label(about_window, text="版本: 2.0", font=("Arial", 12)).pack(pady=5)
        ttk.Label(about_window, text="开发人员: UI/程序大师", font=("Arial", 12)).pack(pady=5)
        ttk.Label(about_window, text="© 2023 高级标签生成工具", font=("Arial", 10)).pack(pady=20)
        
        info_text = tk.Text(about_window, height=6, width=40, font=("Arial", 9), wrap="word")
        info_text.pack(pady=10, padx=20)
        info_text.insert(tk.END, "此工具用于从Excel数据生成自定义标签。支持文本和二维码，可自定义尺寸、颜色、字体和布局。")
        info_text.config(state="disabled")
        
        ttk.Button(about_window, text="确定", command=about_window.destroy).pack(pady=10)
    
    def update_status(self, message):
        self.status_label.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = LabelGeneratorApp(root)
    root.mainloop()