import os
import sys
import json
import importlib
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
from PIL import Image, ImageTk

# ==================== Đường dẫn ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLASS_FILE = os.path.join(BASE_DIR, "core", "data", "classification.xlsx")
COMP_PIC_DIR = os.path.join(BASE_DIR, "component_picture")
REF_PIC_DIR = os.path.join(BASE_DIR, "ref")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LANG_DIR = os.path.join(BASE_DIR, "lang")

# ==================== Lớp FilteredCombobox ====================
class FilteredCombobox(ttk.Combobox):
    """Combobox có hỗ trợ filter khi nhập."""
    def __init__(self, master=None, **kwargs):
        self._full_values = kwargs.pop('values', [])
        super().__init__(master, **kwargs)
        self._current_filter = ""
        self._var = tk.StringVar()
        self.configure(textvariable=self._var)
        self.bind('<KeyRelease>', self._on_keyrelease)
        self.bind('<Button-1>', self._on_click)
        self.code_map = {}
    
    def set_full_values(self, values):
        self._full_values = values
        self['values'] = values
    
    def _on_click(self, event):
        self['values'] = self._full_values
        self._current_filter = ""
    
    def set_code_map(self, code_map):
        """Lưu mapping từ display sang code"""
        self.code_map = code_map
    
    def get_code(self):
        """Lấy code tương ứng với giá trị display hiện tại"""
        selected_display = self._var.get()
        return self.code_map.get(selected_display, selected_display)

    def _on_keyrelease(self, event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab'):
            return
        text = self._var.get().lower()
        if text != self._current_filter:
            self._current_filter = text
            if not text:
                self['values'] = self._full_values
            else:
                filtered = [item for item in self._full_values if text in item.lower()]
                self['values'] = filtered
            if self['values']:
                self.event_generate('<Down>')
    
    def get(self):
        return self._var.get()
    
    def set(self, value):
        self._var.set(value)
        self._current_filter = ""
        self['values'] = self._full_values

# ==================== Quản lý ngôn ngữ ====================
class LanguageManager:
    def __init__(self):
        self.current_lang = "vi"
        self.strings = {}
        self.load_language()
    
    def load_language(self):
        lang_file = os.path.join(LANG_DIR, f"{self.current_lang}.json")
        if os.path.exists(lang_file):
            with open(lang_file, 'r', encoding='utf-8') as f:
                self.strings = json.load(f)
        else:
            self.strings = {
                "app_title": "Tạo Part Number và Description",
                "menu_file": "File",
                "menu_exit": "Thoát",
                "menu_tools": "Công cụ",
                "menu_search": "Tìm Component",
                "menu_settings": "Cài đặt",
                "menu_help": "Trợ giúp",
                "menu_guide": "Hướng dẫn",
                "menu_about": "Giới thiệu",
                "group": "Group",
                "component": "Component:",
                "classification": "Classification",
                "sub_class": "Sub-classification",
                "params": "Nhập các tham số",
                "result": "Kết quả",
                "part_number": "Part Number",
                "description": "Description",
                "parameter": "Tham số",
                "reference": "Quy tắc tham chiếu",
                "component_image": "Linh kiện:",
                "ref_image": "Ảnh linh kiện",
                "select_classification": "Vui lòng chọn đầy đủ phân loại",
                "no_logic": "Không có logic cho linh kiện này",
                "search_window_title": "Tìm kiếm linh kiện",
                "settings_title": "Cài đặt",
                "font": "Font:",
                "font_size": "Cỡ chữ:",
                "resizable": "Cho phép co dãn giao diện",
                "language": "Ngôn ngữ:",
                "apply": "Áp dụng",
                "cancel": "Hủy",
                "guide_title": "Hướng dẫn sử dụng",
                "guide_text": "Chưa có gì để hiển thị ở đây cả, hihi",
                "about_title": "Giới thiệu",
                "about_text": "Phần mềm tạo Part Number và Description\n\nPhiên bản 1.0\nTác giả: Hoàng PT\nEmail: hoang.pt@example.com\n\n© 2025 - Công ty ABC",
                "ok": "Đóng"
            }
    
    def set_language(self, lang):
        self.current_lang = lang
        self.load_language()
    
    def get(self, key):
        return self.strings.get(key, key)

lang = LanguageManager()

# ==================== Đọc dữ liệu phân loại ====================
def load_classification():
    try:
        df = pd.read_excel(CLASS_FILE, sheet_name="class", dtype=str)
        df = df.fillna("")
        for col in ["Component", "Classification", "Sub-classification", "Code"]:
            if col in df.columns:
                df[col] = df[col].astype(str)
        return df
    except Exception as e:
        print(f"Lỗi đọc classification: {e}")
        return pd.DataFrame(columns=["Component", "Classification", "Sub-classification", "Code"])

CLASS_DF = load_classification()

# ==================== Cấu hình ứng dụng ====================
class AppConfig:
    def __init__(self):
        self.font_name = "Arial"
        self.font_size = 11
        self.resizable = False
        self.language = "vi"
        self.load()
    
    def load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.font_name = data.get('font_name', self.font_name)
                    self.font_size = data.get('font_size', self.font_size)
                    self.resizable = data.get('resizable', self.resizable)
                    self.language = data.get('language', self.language)
            except:
                pass
    
    def save(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                'font_name': self.font_name,
                'font_size': self.font_size,
                'resizable': self.resizable,
                'language': self.language
            }, f, indent=2)

config = AppConfig()

# ==================== Hàm áp dụng font ====================
def apply_font_to_all(root, font_name, font_size):
    style = ttk.Style()
    style.configure('.', font=(font_name, font_size))
    style.configure('TLabel', font=(font_name, font_size))
    style.configure('TButton', font=(font_name, font_size))
    style.configure('TEntry', font=(font_name, font_size))
    style.configure('TCombobox', font=(font_name, font_size))
    style.configure('TNotebook.Tab', font=(font_name, font_size))
    def update_font(widget):
        try:
            if isinstance(widget, (tk.Entry, tk.Text, tk.Listbox, tk.Button, tk.Label)):
                widget.configure(font=(font_name, font_size))
        except:
            pass
        for child in widget.winfo_children():
            update_font(child)
    update_font(root)

# ==================== Cửa sổ tìm component ====================
class SearchWindow(Toplevel):
    def __init__(self, parent, df):
        super().__init__(parent)
        self.title(lang.get("search_window_title"))
        self.geometry("800x500")
        self.df = df
        self.filtered_df = df.copy()
        self.create_widgets()
    
    def create_widgets(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        self.search_vars = {}
        columns = list(self.df.columns)
        for i, col in enumerate(columns):
            ttk.Label(top_frame, text=col).grid(row=0, column=i, padx=2)
            var = tk.StringVar()
            var.trace('w', lambda *args, col=col: self.filter_data(col))
            entry = ttk.Entry(top_frame, textvariable=var, width=15)
            entry.grid(row=1, column=i, padx=2)
            self.search_vars[col] = var
        
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        style = ttk.Style()
        style.configure("Custom.Treeview", relief="solid", borderwidth=1)
        style.configure("Custom.Treeview.Heading", font=('Arial', 10, 'bold'), relief="solid")
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style="Custom.Treeview")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor='center')
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.load_data()
    
    def filter_data(self, col):
        mask = pd.Series([True]*len(self.df))
        for col_name, var in self.search_vars.items():
            txt = var.get().strip()
            if txt:
                mask &= self.df[col_name].str.contains(txt, case=False, na=False)
        self.filtered_df = self.df[mask]
        self.load_data()
    
    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for _, row in self.filtered_df.iterrows():
            self.tree.insert('', tk.END, values=list(row))

# ==================== Cửa sổ cài đặt ====================
class SettingsWindow(Toplevel):
    def __init__(self, parent, config_obj, apply_callback):
        super().__init__(parent)
        self.title(lang.get("settings_title"))
        self.geometry("400x350")
        self.config = config_obj
        self.apply_callback = apply_callback
        self.create_widgets()
    
    def create_widgets(self):
        row = 0
        ttk.Label(self, text=lang.get("font")+":").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.font_var = tk.StringVar(value=self.config.font_name)
        fonts = ['Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma']
        font_combo = ttk.Combobox(self, textvariable=self.font_var, values=fonts, state='readonly')
        font_combo.grid(row=row, column=1, padx=5, pady=5, sticky=tk.W)
        row += 1
        
        ttk.Label(self, text=lang.get("font_size")+":").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.size_var = tk.IntVar(value=self.config.font_size)
        sizes = [9,10,11,12,13,14]
        size_combo = ttk.Combobox(self, textvariable=self.size_var, values=sizes, state='readonly')
        size_combo.grid(row=row, column=1, padx=5, pady=5, sticky=tk.W)
        row += 1
        
        self.resizable_var = tk.BooleanVar(value=self.config.resizable)
        check = ttk.Checkbutton(self, text=lang.get("resizable"), variable=self.resizable_var)
        check.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        row += 1
        
        ttk.Label(self, text=lang.get("language")+":").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.lang_var = tk.StringVar(value=self.config.language)
        langs = ['vi', 'en']
        lang_combo = ttk.Combobox(self, textvariable=self.lang_var, values=langs, state='readonly')
        lang_combo.grid(row=row, column=1, padx=5, pady=5, sticky=tk.W)
        row += 1
        
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text=lang.get("apply"), command=self.apply).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text=lang.get("cancel"), command=self.destroy).pack(side=tk.LEFT, padx=10)
    
    def apply(self):
        self.config.font_name = self.font_var.get()
        self.config.font_size = self.size_var.get()
        self.config.resizable = self.resizable_var.get()
        self.config.language = self.lang_var.get()
        self.config.save()
        lang.set_language(self.config.language)
        self.apply_callback()
        self.destroy()

# ==================== Cửa sổ About ====================
class AboutWindow(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(lang.get("about_title"))
        self.geometry("300x250")
        text = lang.get("about_text")
        label = ttk.Label(self, text=text, justify=tk.CENTER)
        label.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        ttk.Button(self, text=lang.get("ok"), command=self.destroy).pack(pady=10)

class Guide(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(lang.get("guide_title"))
        self.geometry("300x250")
        text = lang.get("guide_text")
        label = ttk.Label(self, text=text, justify=tk.CENTER)
        label.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        ttk.Button(self, text=lang.get("ok"), command=self.destroy).pack(pady=10)

# ==================== GUI Chính ====================
class PartNumberGenerator:
    def __init__(self, root):
        self.root = root
        root.title(lang.get("app_title"))
        root.geometry("1000x700")
        
        apply_font_to_all(root, config.font_name, config.font_size)
        root.resizable(config.resizable, config.resizable)
        
        # Biến
        self.group_var = tk.StringVar()
        self.component_var = tk.StringVar()
        self.classification_var = tk.StringVar()
        self.sub_class_var = tk.StringVar()
        self.current_class_code = None
        self.component_logic_instance = None
        self.param_widgets = []     # list các label, widget, typ
        self.current_group = "Other"  # Mặc định là Other
        
        self.create_menu()
        
        # Frame chính
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Hàng trên: Nhập thông tin và tham số
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=5)
        
        # Frame trái: Nhập thông tin
        self.info_frame = ttk.LabelFrame(top_frame, text=lang.get("params"), padding="5")
        self.info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        
        self.lbl_group = ttk.Label(self.info_frame, text=lang.get("group"))
        self.lbl_group.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.group_combo = FilteredCombobox(self.info_frame, textvariable=self.group_var, values=["ABM", "CNN/CKR", "Other"])
        self.group_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_change)
        
        self.lbl_component = ttk.Label(self.info_frame, text=lang.get("component"))
        self.lbl_component.grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.component_combo = FilteredCombobox(self.info_frame, textvariable=self.component_var)
        self.component_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        self.component_combo.bind("<<ComboboxSelected>>", self.on_component_change)
        
        self.lbl_class = ttk.Label(self.info_frame, text=lang.get("classification"))
        self.lbl_class.grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.class_combo = FilteredCombobox(self.info_frame, textvariable=self.classification_var)
        self.class_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        self.class_combo.bind("<<ComboboxSelected>>", self.on_class_change)
        
        self.lbl_sub = ttk.Label(self.info_frame, text=lang.get("sub_class"))
        self.lbl_sub.grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.sub_class_combo = FilteredCombobox(self.info_frame, textvariable=self.sub_class_var)
        self.sub_class_combo.grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)
        self.sub_class_combo.bind("<<ComboboxSelected>>", self.on_change)
        
        # Frame phải: Nhập tham số
        self.param_frame = ttk.LabelFrame(top_frame, text=lang.get("params"), padding="5")
        self.param_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5,0))
        
        # Kết quả
        result_frame = ttk.LabelFrame(main_frame, text=lang.get("result"), padding="5")
        result_frame.pack(fill=tk.X, pady=5)
        
        # Hiển thị Part number
        self.lbl_pn = ttk.Label(result_frame, text=lang.get("part_number"))
        self.lbl_pn.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.pn_entry = ttk.Entry(result_frame, font=("Arial", 11, "bold"), foreground="blue")
        self.pn_entry.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # Hiển thị description
        self.lbl_desc = ttk.Label(result_frame, text=lang.get("description"))
        self.lbl_desc.grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.desc_entry = ttk.Entry(result_frame, width=80)
        self.desc_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)

        # Hiển thị các parameter
        self.lbl_para = ttk.Label(result_frame, text=lang.get("parameter"))
        self.lbl_para.grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.para_entry = ttk.Entry(result_frame, width=80)
        self.para_entry.grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        result_frame.columnconfigure(1, weight=1)
        
        # Hình ảnh
        image_frame = ttk.LabelFrame(main_frame, text=lang.get("reference"), padding="5")
        image_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        comp_img_frame = ttk.Frame(image_frame)
        comp_img_frame.pack(side=tk.LEFT, padx=10, fill=tk.BOTH, expand=True)
        self.lbl_comp_img = ttk.Label(comp_img_frame, text=lang.get("component_image"))
        self.lbl_comp_img.pack()
        self.comp_img_label = ttk.Label(comp_img_frame, relief=tk.SUNKEN)
        self.comp_img_label.pack(fill=tk.BOTH, expand=True)
        
        ref_img_frame = ttk.Frame(image_frame)
        ref_img_frame.pack(side=tk.RIGHT, padx=10, fill=tk.BOTH, expand=True)
        self.lbl_ref_img = ttk.Label(ref_img_frame, text=lang.get("ref_image"))
        self.lbl_ref_img.pack()
        self.ref_img_label = ttk.Label(ref_img_frame, relief=tk.SUNKEN)
        self.ref_img_label.pack(fill=tk.BOTH, expand=True)
        
        self.load_components()
        self.set_window_icon()
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.get("menu_file"), menu=file_menu)
        file_menu.add_command(label=lang.get("menu_exit"), command=self.root.quit)
        
        # Search menu
        search_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.get("menu_search"), menu=search_menu)
        search_menu.add_command(label=lang.get("menu_search"), command=self.open_search)

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.get("menu_tools"), menu=tools_menu)
        #tools_menu.add_command(label=lang.get("menu_search"), command=self.open_search)
        tools_menu.add_command(label=lang.get("menu_settings"), command=self.open_settings)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=lang.get("menu_help"), menu=help_menu)
        help_menu.add_command(label=lang.get("menu_guide"), command=self.open_guide)
        help_menu.add_command(label=lang.get("menu_about"), command=self.open_about)
    
    def update_language(self):
        self.root.title(lang.get("app_title"))
        self.info_frame.config(text=lang.get("params"))
        self.param_frame.config(text=lang.get("params"))
        self.lbl_group.config(text=lang.get("group"))
        self.lbl_component.config(text=lang.get("component"))
        self.lbl_class.config(text=lang.get("classification"))
        self.lbl_sub.config(text=lang.get("sub_class"))
        self.lbl_pn.config(text=lang.get("part_number"))
        self.lbl_desc.config(text=lang.get("description"))
        self.lbl_comp_img.config(text=lang.get("component_image"))
        self.lbl_ref_img.config(text=lang.get("ref_image"))
        self.update_menu()
    
    def update_menu(self):
        self.root.config(menu=tk.Menu())
        self.create_menu()
    
    def open_search(self):
        if CLASS_DF is not None and not CLASS_DF.empty:
            SearchWindow(self.root, CLASS_DF)
        else:
            messagebox.showwarning("Lỗi", "Không có dữ liệu classification!")
    
    def open_settings(self):
        def apply():
            apply_font_to_all(self.root, config.font_name, config.font_size)
            self.root.resizable(config.resizable, config.resizable)
            self.update_language()
        SettingsWindow(self.root, config, apply)
    
    def open_about(self):
        AboutWindow(self.root)

    def open_guide(self):
        Guide(self.root)
    
    def set_window_icon(self):
        try:
            icon_png = os.path.join(BASE_DIR, "pictures", "icon", "icon.png")
            img = Image.open(icon_png)
            img = img.resize((32, 32), Image.Resampling.LANCZOS)
            icon_photo = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, icon_photo)
            print("Đã đặt icon cho ứng dụng")
        except Exception as e:
            print(f"Không thể đặt icon: {e}")
    
    def load_components(self):
        if CLASS_DF.empty:
            self.component_combo.set_full_values([])
            return
        comps = sorted([str(v) for v in CLASS_DF["Component"].unique() if str(v).strip() != ""])
        self.component_combo.set_full_values(comps)
    
    def on_group_change(self, event=None):
        self.current_group = self.group_combo.get()
        self.update_parameter_frame()
        self.update_result()
    
    def on_component_change(self, event=None):
        component = self.component_combo.get()
        if not component:
            return
        
        # Xóa các widget tham số cũ
        for w in self.param_frame.winfo_children():
            w.destroy()
        self.param_widgets.clear()
        
        # Load module logic tương ứng
        try:
            module_name = f"core.logic.{component.lower()}"
            core_path = os.path.join(BASE_DIR, "core")
            if core_path not in sys.path:
                sys.path.insert(0, core_path)
            self.component_logic = importlib.import_module(module_name)
            
            possible = [component.capitalize()+"Logic", component.upper()+"Logic", component.lower()+"logic", component.title()+"Logic"]
            logic_class = None
            for name in possible:
                if hasattr(self.component_logic, name):
                    logic_class = getattr(self.component_logic, name)
                    break
            if logic_class:
                self.component_logic_instance = logic_class()
            else:
                self.component_logic_instance = self.component_logic
        except ImportError as e:
            print(f"Không tìm thấy module cho {component}: {e}")
            self.component_logic_instance = None
            ttk.Label(self.param_frame, text=lang.get("no_logic"), foreground="red").pack()
            return
        
        if CLASS_DF.empty:
            return
        df_comp = CLASS_DF[CLASS_DF["Component"] == component]
        classes = sorted(df_comp["Classification"].unique())
        self.class_combo.set_full_values(classes)
        self.class_combo.set("")
        self.sub_class_combo.set("")
        self.sub_class_combo.set_full_values([])
        self.on_change()
    
    def on_class_change(self, event=None):
        comp = self.component_combo.get()
        cls = self.class_combo.get()
        if not comp or not cls or CLASS_DF.empty:
            return
        df_sub = CLASS_DF[(CLASS_DF["Component"] == comp) & (CLASS_DF["Classification"] == cls)]
        subs = sorted(df_sub["Sub-classification"].unique())
        self.sub_class_combo.set_full_values(subs)
        self.sub_class_combo.set("")
        self.on_change()
    
    def on_change(self, event=None):
        self.update_parameter_frame()
        self.update_result()
    
    def get_current_class_code(self):
        comp = self.component_combo.get()
        cls = self.class_combo.get()
        sub = self.sub_class_combo.get()
        if not comp or not cls or CLASS_DF.empty:
            return None
        mask = (CLASS_DF["Component"] == comp) & (CLASS_DF["Classification"] == cls)
        if sub:
            mask &= (CLASS_DF["Sub-classification"] == sub)
        matched = CLASS_DF[mask]
        if len(matched) == 0 and not sub:
            mask2 = (CLASS_DF["Component"] == comp) & (CLASS_DF["Classification"] == cls) & (CLASS_DF["Sub-classification"] == "")
            matched = CLASS_DF[mask2]
        if len(matched) == 0:
            return None
        return matched.iloc[0]["Code"]
    
    def update_parameter_frame(self):
        # Xóa các widget cũ
        for w in self.param_frame.winfo_children():
            w.destroy()
        self.param_widgets.clear()
        
        # Lấy class-code
        class_code = self.get_current_class_code()
        if class_code is None:
            ttk.Label(self.param_frame, text=lang.get("select_classification"), foreground="red").pack()
            return
        self.current_class_code = class_code
        
        # Kiểm tra component logic instance
        if self.component_logic_instance is None:
            ttk.Label(self.param_frame, text=lang.get("no_logic"), foreground="red").pack()
            return
        
        # Gọi get_parameter_fields với group hiện tại
        if hasattr(self.component_logic_instance, 'get_parameter_fields'):
            fields = self.component_logic_instance.get_parameter_fields(class_code, self.current_group)
        else:
            fields = []
        
        # Tạo các widget
        for i, field_info in enumerate(fields):
            if len(field_info) == 4:
                label, typ, display_list, code_map = field_info
            else:
                label, typ, source = field_info
                display_list, code_map = None, None
                if typ == "combo" and source and hasattr(self.component_logic_instance, 'sheet_data'):
                    df = self.component_logic_instance.sheet_data.get(source, pd.DataFrame())
                    if not df.empty:
                        display_list = df["display"].tolist()
                        code_map = dict(zip(df["display"], df["code"]))
            
            ttk.Label(self.param_frame, text=label+":").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            
            if typ == "combo" and display_list:
                combo = FilteredCombobox(self.param_frame)
                combo.set_full_values(display_list)
                combo.set_code_map(code_map)
                combo.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
                combo.bind("<<ComboboxSelected>>", self.on_param_change)
                #combo.code_map = code_map
                self.param_widgets.append((label, combo, "combo"))
            else:
                entry = ttk.Entry(self.param_frame)
                entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
                entry.bind("<KeyRelease>", self.on_param_change)
                self.param_widgets.append((label, entry, "entry"))
    
    def on_param_change(self, event=None):
        self.update_result()
    
    def update_result(self):
        if not self.component_logic_instance or not self.current_class_code:
            return
        
        # Thu thập values từ các widget
        code_values = {}        # code cho part number
        display_values = {}     # giá trị hiển thị cho description

        for label, widget, typ in self.param_widgets:
            if typ == "combo":
                display_val = widget.get()
                code_val = widget.get_code() if hasattr(widget, 'get_code') else display_val
                code_values[label] = code_val
                display_values[label] = display_val
            else:  # entry
                val = widget.get().strip()
                code_values[label] = val
                display_values[label] = val
        
        # Gọi generate_part_number với group hiện tại
        pn = ""
        if hasattr(self.component_logic_instance, 'generate_part_number'):
            pn = self.component_logic_instance.generate_part_number(
                self.current_class_code, code_values, self.current_group)
        
        # Gọi generate_description với group hiện tại
        desc = ""
        if hasattr(self.component_logic_instance, 'generate_description'):
            desc = self.component_logic_instance.generate_description(
                self.current_class_code, display_values, code_values, self.current_group)

        # Gọi generate_parameter với group hiện tại
        para = ""
        if hasattr(self.component_logic_instance, 'generate_parameter'):
            para = self.component_logic_instance.generate_parameter(
                self.current_class_code, display_values, self.current_group)
        
        self.pn_entry.delete(0, tk.END)
        self.pn_entry.insert(0, pn)
        
        self.desc_entry.delete(0, tk.END)
        self.desc_entry.insert(0, desc)
        
        self.para_entry.delete(0, tk.END)
        self.para_entry.insert(0, para)
        
        self.update_images(pn)
    
    def update_images(self, pn):
        comp = self.component_combo.get()
        if comp:
            for ext in ['.jpg','.png']:
                path = os.path.join(COMP_PIC_DIR, comp.lower()+ext)
                if os.path.exists(path):
                    self.load_image(path, self.comp_img_label)
                    break
        if pn:
            for ext in ['.jpg','.png']:
                path = os.path.join(REF_PIC_DIR, pn+ext)
                if os.path.exists(path):
                    self.load_image(path, self.ref_img_label)
                    break
    
    def load_image(self, path, label):
        try:
            img = Image.open(path)
            img.thumbnail((200,200))
            photo = ImageTk.PhotoImage(img)
            label.config(image=photo)
            label.image = photo
        except:
            label.config(image="", text="Không có ảnh")

# ==================== Chạy chương trình ====================
if __name__ == "__main__":
    lang.set_language(config.language)
    root = tk.Tk()
    app = PartNumberGenerator(root)
    root.mainloop()