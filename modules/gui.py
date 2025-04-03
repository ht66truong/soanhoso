# Thư viện chuẩn
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import sys
import json
import logging
from datetime import datetime
import requests
from tkinter import messagebox

# Thư viện bên thứ ba
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES

# Module nội bộ
from modules.utils import ToolTip, create_popup, number_to_words, create_popup_with_notebook
from modules.export import (ExportManager)


# Thiết lập thư mục AppData và logging
appdata_dir = "AppData"
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)
logging.basicConfig(filename=os.path.join(appdata_dir, "app.log"), level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logging.getLogger("docxtpl").setLevel(logging.ERROR)


class DataEntryApp:
    def __init__(self, root):
        """Khởi tạo ứng dụng nhập liệu hồ sơ kinh doanh."""
        self.root = root
        self.root.title("Ứng dụng nhập liệu hồ sơ Kinh Doanh v6.6")
        window_width = 1280
        window_height = 768
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        #self.root.state('zoomed')  # Tùy chọn: Phóng to tối đa khi khởi động       
        self.root.configure(bg="#f8f9fa")
        logging.info("Ứng dụng khởi động")

        # Cấu hình file trong thư mục AppData
        self.appdata_dir = appdata_dir
        self.configs_file = os.path.join(self.appdata_dir, "configs.json")
        self.templates_dir = "templates"
        self.backup_dir = "backup"

        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)

        self.default_fields = [
            "ten_doanh_nghiep", "ten_doanh_nghiep_bang_tieng_anh", "ten_viet_tat",
            "ma_so_doanh_nghiep", "ngay_cap_mst", "so_nha_ten_duong", "xa_phuong", "quan_huyen",
            "tinh_thanh_pho", "so_dien_thoai", "von_đieu_le", "ho_ten", "gioi_tinh", "sinh_ngay",
            "dan_toc", "quoc_tich", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han", "dia_chi_thuong_tru", "dia_chi_lien_lac",
            "ho_ten_uq", "gioi_tinh_uq", "sinh_ngay_uq", "so_cccd_uq", "ngay_cap_uq", "noi_cap_uq", "dia_chi_lien_lac_uq", "sdt_uq", "email_uq"
        ]
        # Định nghĩa các cột cho bảng thông tin thành viên
        self.member_columns = [
            "ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich", "loai_giay_to",
            "so_cccd", "ngay_cap", "noi_cap",
            "ngay_het_han", "dia_chi_thuong_tru", "dia_chi_lien_lac", "von_gop", "ty_le_gop", "ngay_gop_von"
        ]

        self.configs = {}
        self.current_config_name = None
        self.field_groups = {}
        self.fields = []
        self.saved_entries = []
        self.drag_item = None
        self.entries = {}
        self.labels = {}
        self.current_tab_index = 0

        # Initialize manager classes first
        self.config_manager = ConfigManager(self)
        self.field_manager = FieldManager(self)
        self.tab_manager = TabManager(self)
        self.data_manager = DataManager(self)
        self.member_manager = MemberManager(self)
        self.industry_manager = IndustryManager(self)
        self.template_manager = TemplateManager(self)
        self.export_manager = ExportManager(self)
        self.backup_manager = BackupManager(self)
        
        

        # Khởi tạo top_frame trước
        self.top_frame = ttk.Frame(self.root)
        self.top_frame.pack(side="top", fill="x", padx=5, pady=5)
        
        self.main_frame = ttk.Frame(root, padding=5)
        self.main_frame.pack(fill="both", expand=True)

        # Config Frame
        self.config_frame = ttk.LabelFrame(self.main_frame, text="Quản lý cấu hình", padding=5)
        self.config_frame.pack(fill="x", pady=(0, 5))

        self.config_var = tk.StringVar()
        self.config_dropdown = ttk.Combobox(self.config_frame, textvariable=self.config_var, values=list(self.configs.keys()), state="readonly", width=15)
        self.config_dropdown.pack(side="left", padx=5)
        ToolTip(self.config_dropdown, "Chọn cấu hình để thêm/sửa/xóa")
        self.config_dropdown.bind("<<ComboboxSelected>>", self.load_selected_config)

        def resource_path(relative_path): # Hàm hỗ trợ lấy đường dẫn tài nguyên
            """Lấy đường dẫn tuyệt đối đến tài nguyên, hoạt động cho cả mã nguồn và file thực thi PyInstaller."""
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # Load icons
        try:
            add_icon = Image.open(resource_path("icon/add_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            delete_icon = Image.open(resource_path("icon/delete_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            save_icon = Image.open(resource_path("icon/save_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            clear_icon = Image.open(resource_path("icon/clear_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            import_icon = Image.open(resource_path("icon/imxls_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            xls_icon = Image.open(resource_path("icon/xls_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            preview_icon = Image.open(resource_path("icon/preview_icon.png")).resize((40, 40), Image.Resampling.LANCZOS)
            export_icon = Image.open(resource_path("icon/export_icon.png")).resize((40, 40), Image.Resampling.LANCZOS)
            remove_icon = Image.open(resource_path("icon/remove_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            edit_icon = Image.open(resource_path("icon/edit_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            search_icon = Image.open(resource_path("icon/search_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            find_icon = Image.open(resource_path("icon/find_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            restorebackup_icon = Image.open(resource_path("icon/restorebackup_icon.png")).resize((20, 20), Image.Resampling.LANCZOS)
            self.add_icon_img = ImageTk.PhotoImage(add_icon)
            self.delete_icon_img = ImageTk.PhotoImage(delete_icon)
            self.save_icon_img = ImageTk.PhotoImage(save_icon)
            self.clear_icon_img = ImageTk.PhotoImage(clear_icon)
            self.import_icon_img = ImageTk.PhotoImage(import_icon)
            self.xls_icon_img = ImageTk.PhotoImage(xls_icon)
            self.preview_icon_img = ImageTk.PhotoImage(preview_icon)
            self.export_icon_img = ImageTk.PhotoImage(export_icon)
            self.remove_icon_img = ImageTk.PhotoImage(remove_icon)
            self.edit_icon_img = ImageTk.PhotoImage(edit_icon)
            self.search_icon_img = ImageTk.PhotoImage(search_icon)
            self.find_icon_img = ImageTk.PhotoImage(find_icon)
            self.restorebackup_icon_img = ImageTk.PhotoImage(restorebackup_icon)
        except Exception as e:
            logging.error(f"Lỗi khi tải biểu tượng: {str(e)}")
            messagebox.showerror("Lỗi", "Không thể tải biểu tượng, kiểm tra thư mục icon!")

        #Quản lý cấu hình
        self.add_config_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_new_config, style="info.TButton")
        ToolTip(self.add_config_button, "Thêm cấu hình")
        self.add_config_button.pack(side="left", padx=5)
        self.delete_config_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_config, style="danger.TButton")
        ToolTip(self.delete_config_button, "Xóa cấu hình")
        self.delete_config_button.pack(side="left", padx=5)
        self.rename_config_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_config, style="warning.TButton")
        ToolTip(self.rename_config_button, "Sửa tên cấu hình")
        self.rename_config_button.pack(side="left", padx=5)

        #Quản lý tab
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=10, fill="y")
        #ttk.Label(self.config_frame, text="Quản lý tab:").pack(side="left", padx=5)
        self.tab_var = tk.StringVar()
        self.tab_dropdown = ttk.Combobox(self.config_frame, textvariable=self.tab_var, state="readonly", width=20)
        ToolTip(self.tab_dropdown, "Chọn tab để thêm/sửa/xóa")
        self.tab_dropdown.pack(side="left", padx=5)
        self.add_tab_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_tab, style="info.TButton")
        ToolTip(self.add_tab_button, "Thêm tab")
        self.add_tab_button.pack(side="left", padx=5)
        self.delete_tab_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_tab, style="danger.TButton")
        ToolTip(self.delete_tab_button, "Xóa tab")
        self.delete_tab_button.pack(side="left", padx=5)
        self.rename_tab_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_tab, style="warning.TButton")
        ToolTip(self.rename_tab_button, "Sửa tên tab")
        self.rename_tab_button.pack(side="left", padx=5)

        #Quản lý trường
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=10, fill="y")
        #ttk.Label(self.config_frame, text="Quản lý trường:").pack(side="left", padx=5)
        self.field_var = tk.StringVar()
        self.field_dropdown = ttk.Combobox(self.config_frame, textvariable=self.field_var, state="readonly", width=20)
        ToolTip(self.field_dropdown, "Chọn trường để thêm/sửa/xóa")
        self.field_dropdown.pack(side="left", padx=5)
        # Thêm nút "Thêm trường"
        self.add_field_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_field, style="info.TButton")
        ToolTip(self.add_field_button, "Thêm trường")
        self.add_field_button.pack(side="left", padx=5)
        self.delete_field_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_selected_field, style="danger.TButton")
        ToolTip(self.delete_field_button, "Xóa trường")
        self.delete_field_button.pack(side="left", padx=5)
        self.rename_field_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_selected_field, style="warning.TButton")
        ToolTip(self.rename_field_button, "Sửa tên trường")
        self.rename_field_button.pack(side="left", padx=5)

        #Quản lý template"
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=10, fill="y")
        #ttk.Label(self.config_frame, text="Quản lý mẫu:").pack(side="left", padx=5) 
        self.template_frame = ttk.Frame(self.config_frame)
        self.template_frame.pack(side="left", fill="x", padx=5)
        ToolTip(self.template_frame, "Kéo thả để thêm mẫu/sắp xếp")
        self.template_tree = ttk.Treeview(self.template_frame, columns=(), show="tree", height=5, selectmode="extended")
        self.template_tree.column("#0", width=200)
        self.template_tree.pack(fill="both", expand=True)
        self.update_template_tree()
        self.template_frame.drop_target_register(DND_FILES)
        self.template_frame.dnd_bind('<<Drop>>', self.drop_template_files)
        self.template_tree.bind("<Button-1>", self.start_drag)
        self.template_tree.bind("<B1-Motion>", self.drag_template)
        self.template_tree.bind("<ButtonRelease-1>", self.drop_template)
        self.template_tree.bind("<Button-3>", self.show_template_context_menu)

        # Control Frame
        self.control_frame = ttk.LabelFrame(self.main_frame, text="Quản lý dữ liệu", padding=10)
        self.control_frame.pack(fill="x", pady=10)

        # Thêm khung tìm kiếm
        search_frame = ttk.Frame(self.control_frame)
        search_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="w")
        ToolTip(search_frame, "Tìm kiếm dữ liệu")
        ttk.Label(search_frame, text="Tìm kiếm:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_combobox = ttk.Combobox(search_frame, textvariable=self.search_var, width=20, state="normal")
        self.search_combobox.pack(side="left", padx=5)
        self.search_combobox["values"] = [entry["name"] for entry in self.saved_entries]  # Gán danh sách dữ liệu
        self.search_combobox.bind("<<ComboboxSelected>>", self.load_data_from_search)  # Gắn sự kiện
        self.search_combobox.bind("<KeyRelease>", self.update_search_suggestions)
        ttk.Button(search_frame, image=self.find_icon_img, command=self.search_data, style="primary.TButton").pack(side="left", padx=5)

        #ttk.Label(self.control_frame, text="Chọn công ty:").grid(row=0, column=2, padx=5, pady=5)
        self.load_data_var = tk.StringVar()
        self.load_data_dropdown = ttk.Combobox(self.control_frame, textvariable=self.load_data_var, state="readonly", width=10)
        self.load_data_dropdown.grid(row=0, column=3, padx=5, pady=5)
        ToolTip(self.load_data_dropdown, "Chọn công ty")
        self.load_data_dropdown.bind("<<ComboboxSelected>>", self.load_selected_entry)

        self.add_data_button = ttk.Button(self.control_frame, image=self.add_icon_img, command=self.add_entry_data, style="info.TButton")
        ToolTip(self.add_data_button, "Thêm dữ liệu")
        self.add_data_button.grid(row=0, column=4, padx=5, pady=5)
        self.delete_data_button = ttk.Button(self.control_frame, image=self.delete_icon_img, command=self.delete_entry_data, style="danger.TButton")
        ToolTip(self.delete_data_button, "Xóa dữ liệu")
        self.delete_data_button.grid(row=0, column=5, padx=5, pady=5)
        self.rename_data_button = ttk.Button(self.control_frame, image=self.edit_icon_img, command=self.rename_entry_data, style="warning.TButton")
        ToolTip(self.rename_data_button, "Sửa tên dữ liệu")
        self.rename_data_button.grid(row=0, column=6, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=7, padx=10, pady=5, sticky="ns")
        #self.edit_data_button = ttk.Button(self.control_frame, image=self.save_icon_img, command=self.save_entry_data, style="success.TButton")
        #ToolTip(self.edit_data_button, "Lưu thông tin")
        #self.edit_data_button.grid(row=0, column=6, padx=5, pady=5)
        self.clear_data_button = ttk.Button(self.control_frame, image=self.clear_icon_img, command=self.clear_entries, style="danger.TButton")
        ToolTip(self.clear_data_button, "Xóa thông tin")
        self.clear_data_button.grid(row=0, column=8, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=9, padx=10, pady=5, sticky="ns")
        ttk.Label(self.control_frame, text="Xuất Nhập dữ liệu:").grid(row=0, column=10, padx=5, pady=5)
        self.export_excel_button = ttk.Button(self.control_frame, image=self.xls_icon_img, command=self.export_data, style="primary.TButton")
        ToolTip(self.export_excel_button, "Xuất dữ liệu")
        self.export_excel_button.grid(row=0, column=11, padx=5, pady=5)
        self.import_data_button = ttk.Button(self.control_frame, image=self.import_icon_img, command=self.import_from_file, style="primary.TButton")
        ToolTip(self.import_data_button, "Nhập dữ liệu")
        self.import_data_button.grid(row=0, column=12, padx=5, pady=5)
        
        self.restore_data_button = ttk.Button(self.control_frame, image=self.restorebackup_icon_img, command=self.restore_from_backup, style="primary.TButton")
        ToolTip(self.restore_data_button, "Khôi phục từ sao lưu")
        self.restore_data_button.grid(row=0, column=13, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=15, padx=10, pady=5, sticky="ns")
        self.show_placeholder_button = ttk.Button(self.control_frame, image=self.search_icon_img, command=self.show_placeholder_popup, style="danger.TButton")
        ToolTip(self.show_placeholder_button, "Hiển thị danh sách placeholder")
        self.show_placeholder_button.grid(row=0, column=16, padx=5, pady=5)

        # Notebook với tiêu đề "Quản lý nhập liệu"
        self.notebook_frame = ttk.LabelFrame(self.main_frame, text="Quản lý nhập liệu", padding=5)
        self.notebook_frame.pack(fill="both", expand=True, pady=10)
        self.notebook = ttk.Notebook(self.notebook_frame)
        self.notebook.pack(fill="both", expand=True)
        # Gắn sự kiện khi tab thay đổi
        self.notebook.bind("<<NotebookTabChanged>>", lambda event: self.update_field_dropdown())
        
        # Export Frame
        self.export_frame = ttk.Frame(self.main_frame)
        self.export_frame.pack(fill="x", pady=10)
        self.export_frame.grid_columnconfigure(0, weight=1)
        self.export_frame.grid_columnconfigure(3, weight=1)
        self.preview_word_button = ttk.Button(self.export_frame, image=self.preview_icon_img, command=self.preview_word, style="secondary.TButton")
        ToolTip(self.preview_word_button, "Xem trước")
        self.preview_word_button.grid(row=0, column=1, padx=5)
        self.export_file_button = ttk.Button(self.export_frame, image=self.export_icon_img, command=self.export_file, style="danger.TButton")
        ToolTip(self.export_file_button, "Xuất file")
        self.export_file_button.grid(row=0, column=2, padx=5)

        # Styles
        style = ttk.Style()
        style.theme_use("flatly")
        style.configure("TButton", font=("Segoe UI", 10), borderwidth=1, relief="flat", padding=5)
        style.configure("info.TButton", background="#4582ec", foreground="white", bordercolor="#4582ec", borderradius=90)
        style.map("info.TButton", background=[("active", "#3572d8")])
        style.configure("danger.TButton", background="#adb5bd", foreground="black", bordercolor="#adb5bd", borderradius=90, focusthickness=0)
        style.map("danger.TButton", background=[("active", "#9ba3ab")], bordercolor=[("active", "#9ba3ab"), ("focus", "#adb5bd")], foreground=[("active", "black")])
        style.configure("warning.TButton", background="#f0ad4e", foreground="black", bordercolor="#f0ad4e", borderradius=90)
        style.map("warning.TButton", background=[("active", "#e89b3c")])
        style.configure("success.TButton", background="#02b875", foreground="white", bordercolor="#02b875", borderradius=90)
        style.map("success.TButton", background=[("active", "#019c62")])
        style.configure("primary.TButton", background="#4582ec", foreground="white", bordercolor="#4582ec", borderradius=90)
        style.map("primary.TButton", background=[("active", "#3572d8")])
        style.configure("secondary.TButton", background="#4582ec", foreground="white", bordercolor="#4582ec", borderradius=90)
        style.map("secondary.TButton", background=[("active", "#3572d8")])
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TCombobox", font=("Segoe UI", 10))

        self.config_manager.load_configs()
        if not self.config_manager.configs:
            self.config_manager.initialize_default_config()
        self.config_dropdown["values"] = list(self.config_manager.configs.keys())
        if self.config_manager.configs:
            self.config_dropdown.set(list(self.config_manager.configs.keys())[0])
            self.config_manager.load_selected_config(None)
        
        self.root.after(600000, self.backup_manager.auto_backup)
     
    def load_selected_config(self, event):
        self.config_manager.load_selected_config(event)

    def load_selected_entry(self, event):
        """Tải dữ liệu được chọn từ dropdown hoặc tìm kiếm."""
        selected_name = self.load_data_var.get()
        for entry in self.saved_entries:
            if entry["name"] == selected_name:
                for field, value in entry["data"].items():
                    if field in self.entries:
                        self.entries[field].delete(0, tk.END)
                        self.entries[field].insert(0, value)
                if hasattr(self, 'industry_tree'):  # Tải dữ liệu ngành nghề
                    self.industry_manager.load_industry_data()
                if hasattr(self, 'additional_industry_tree'):  # Tải dữ liệu ngành bổ sung
                    self.industry_manager.load_additional_industry_data()
                if hasattr(self, 'removed_industry_tree'):  # Tải dữ liệu ngành giảm
                    self.industry_manager.load_removed_industry_data()
                if hasattr(self, 'adjusted_industry_tree'):  # Tải dữ liệu ngành điều chỉnh
                    self.industry_manager.load_adjusted_industry_data()
                if hasattr(self, 'member_tree'):  # Tải dữ liệu thành viên
                    self.member_manager.load_member_data()
                break

    def update_field_dropdown(self):
        """Cập nhật danh sách trường trong dropdown dựa trên tab hiện tại."""
        if not self.notebook.tabs():  # Check if there are no tabs
            self.field_dropdown["values"] = []
            self.field_var.set("")
            return

        current_tab = self.notebook.tab(self.notebook.select(), "text")
        fields = self.field_groups.get(current_tab, [])
        self.field_dropdown["values"] = fields
        self.field_var.set(fields[0] if fields else "")

    def add_new_config(self):
        self.config_manager.add_new_config()

    def delete_config(self):
        self.config_manager.delete_config()

    def rename_config(self):
        self.config_manager.rename_config()

    def add_field(self):
        self.field_manager.add_field()

    def delete_selected_field(self):
        self.field_manager.delete_selected_field()

    def rename_selected_field(self):
        self.field_manager.rename_selected_field()

    def rename_field(self, field):
        self.field_manager.rename_field(field)

    def delete_field(self, field):
        self.field_manager.delete_field(field)

    def add_tab(self):
        self.tab_manager.add_tab()

    def delete_tab(self):
        self.tab_manager.delete_tab()

    def rename_tab(self):
        self.tab_manager.rename_tab()

    def clear_entries(self):
        self.data_manager.clear_entries()

    def clear_tabs(self): 
        self.tab_manager.clear_tabs()

    def add_entry_context_menu(self, entry):
        """Thêm menu ngữ cảnh cho ô nhập liệu."""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Sao chép", command=lambda: entry.event_generate("<<Copy>>"))
        context_menu.add_command(label="Dán", command=lambda: entry.event_generate("<<Paste>>"))
        context_menu.add_command(label="Xóa", command=lambda: entry.delete(0, tk.END))
        # Lưu menu vào thuộc tính context_menu của entry
        entry.context_menu = context_menu

    def add_entry_data(self):
        self.data_manager.add_entry_data()

    def save_entry_data(self):
        self.data_manager.save_entry_data()

    def delete_entry_data(self):
        self.data_manager.delete_entry_data()

    def rename_entry_data(self):
        self.data_manager.rename_entry_data()

    def create_member_tab(self):
        self.member_manager.create_member_tab()

    def load_member_data(self):
        self.member_manager.load_member_data()

    def add_member(self):
        self.member_manager.add_member()

    def delete_member(self):
        self.member_manager.delete_member()

    def edit_member(self):
        self.member_manager.edit_member()

    def view_member_details(self, event=None):
        self.member_manager.view_member_details(event)

    def start_drag_member(self, event):
        self.member_manager.start_drag_member(event)

    def drag_member(self, event):
        self.member_manager.drag_member(event)

    def drop_member(self, event):
        self.member_manager.drop_member(event)

    def create_industry_tab(self):
        self.industry_manager.create_industry_tab()

    def load_industry_data(self):
        self.industry_manager.load_industry_data()

    def add_industry(self):
        self.industry_manager.add_industry()

    def delete_industry(self):
        self.industry_manager.delete_industry()

    def edit_industry(self, event=None):
        self.industry_manager.edit_industry(event)

    def view_industry_details(self, event=None):
        self.industry_manager.view_industry_details(event)

    def set_main_industry(self):
        self.industry_manager.set_main_industry()

    def update_template_tree(self):
        self.template_manager.update_template_tree()

    def drop_template_files(self, event):
        self.template_manager.drop_template_files(event)

    def add_multiple_templates(self):
        self.template_manager.add_multiple_templates()

    def delete_template(self):
        self.template_manager.delete_template()

    def show_template_context_menu(self, event):
        self.template_manager.show_template_context_menu(event)

    def start_drag(self, event):
        self.template_manager.start_drag(event)

    def drag_template(self, event):
        self.template_manager.drag_template(event)

    def drop_template(self, event):
        self.template_manager.drop_template(event)

    def export_data(self):
        self.export_manager.export_data()

    def import_from_file(self):
        self.export_manager.import_from_file()

    def show_placeholder_popup(self):
        self.export_manager.show_placeholder_popup()

    def copy_placeholder_from_popup(self, placeholder_list):
        self.export_manager.copy_placeholder_from_popup(placeholder_list)

    def export_placeholders(self):
        self.export_manager.export_placeholders()

    def check_template_placeholders(self, doc_paths, data_lower):
        self.export_manager.check_template_placeholders(doc_paths, data_lower)

    def preview_word(self):
        self.export_manager.preview_word()

    def export_file(self):
        self.export_manager.export_file()

    def show_export_popup(self, export_type):
        self.export_manager.show_export_popup(export_type)

    def export_preview(self, doc_paths, data_lower, mode):
        self.export_manager.export_preview(doc_paths, data_lower, mode)

    def export_to_word(self, doc_paths, data_lower, mode):
        self.export_manager.export_to_word(doc_paths, data_lower, mode)

    def export_to_pdf(self, doc_paths, data_lower, mode):
        self.export_manager.export_to_pdf(doc_paths, data_lower, mode)

    def auto_backup(self):
        self.backup_manager.auto_backup()

    def restore_from_backup(self):
        self.backup_manager.restore_from_backup()

    def search_data(self):
        """Tìm kiếm dữ liệu dựa trên từ khóa."""
        keyword = self.search_var.get().lower()
        if not keyword:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập từ khóa tìm kiếm!")
            return

        # Lọc dữ liệu dựa trên từ khóa
        results = []
        for entry in self.saved_entries:
            if keyword in entry["name"].lower() or any(keyword in str(value).lower() for value in entry["data"].values()):
                results.append(entry)

        if not results:
            messagebox.showinfo("Kết quả", "Không tìm thấy dữ liệu phù hợp!")
        else:
            # Hiển thị kết quả tìm kiếm (có thể cập nhật Treeview hoặc popup)
            self.display_search_results(results)
    
    def load_data_from_search(self, event):
        """Tải dữ liệu từ dropdown tìm kiếm."""
        selected_name = self.search_var.get()
        if not selected_name:
            return

        # Đồng bộ giá trị vào load_data_var
        self.load_data_var.set(selected_name)

        # Tìm và tải dữ liệu tương ứng
        for entry in self.saved_entries:
            if entry["name"] == selected_name:
                # Điền dữ liệu vào các trường nhập liệu
                for field, value in entry["data"].items():
                    if field in self.entries:
                        self.entries[field].delete(0, tk.END)
                        self.entries[field].insert(0, value)

                # Tải dữ liệu ngành nghề (nếu có)
                if hasattr(self, 'industry_tree'):
                    self.industry_tree.delete(*self.industry_tree.get_children())
                    for industry in entry["data"].get("nganh_nghe", []):
                        self.industry_tree.insert(
                            "", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else "")
                        )

                # Tải dữ liệu thành viên (nếu có)
                if hasattr(self, 'member_tree'):
                    self.member_tree.delete(*self.member_tree.get_children())
                    for member in entry["data"].get("thanh_vien", []):
                        self.member_tree.insert(
                            "", "end", values=(
                                member.get("ho_ten", ""),
                                member.get("gioi_tinh", ""),
                                member.get("ngay_sinh", ""),
                                member.get("chuc_danh", "")
                            )
                        )

                # Xóa nội dung trong ô tìm kiếm
                self.search_var.set("")
                messagebox.showinfo("Thành công", f"Đã tải dữ liệu '{selected_name}'!")
                break

    def display_search_results(self, results):
        """Hiển thị kết quả tìm kiếm trong popup."""
        popup = create_popup(self.root, "Kết quả tìm kiếm", 600, 400)
        tree = ttk.Treeview(popup, columns=("name", "details"), show="headings")
        tree.heading("name", text="Tên")
        tree.heading("details", text="Chi tiết")
        tree.column("name", width=200)
        tree.column("details", width=350)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for result in results:
            tree.insert("", "end", values=(result["name"], str(result["data"])))

        def load_selected_data():
            """Tải dữ liệu được chọn vào khu vực 'Quản lý nhập liệu'."""
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một mục để tải!")
                return
            selected_name = tree.item(selected_item[0])["values"][0]
            for entry in self.saved_entries:
                if entry["name"] == selected_name:
                    self.load_data_var.set(selected_name)
                    self.load_selected_entry(None)
                    popup.destroy()
                    self.search_var.set("")
                    messagebox.showinfo("Thành công", f"Đã tải dữ liệu '{selected_name}'!")
                    break

        # Gắn sự kiện nhấn đúp chuột vào Treeview
        def on_double_click(event):
            load_selected_data()

        tree.bind("<Double-1>", on_double_click)

        ttk.Button(popup, text="Chọn", command=load_selected_data, style="primary.TButton").pack(side="left", padx=10, pady=10)
        ttk.Button(popup, text="Đóng", command=popup.destroy, style="danger.TButton").pack(side="right", padx=10, pady=10)

    def update_treeview_with_results(self, results):
        """Cập nhật Treeview chính với kết quả tìm kiếm."""
        self.template_tree.delete(*self.template_tree.get_children())
        for result in results:
            self.template_tree.insert("", "end", text=result["name"], values=(result["data"]))
    
    def update_search_suggestions(self, event):
        """Cập nhật danh sách gợi ý trong ô tìm kiếm."""
        input_text = self.search_var.get().lower()
        suggestions = [
            entry["name"]
            for entry in self.saved_entries
            if input_text in entry["name"].lower()
        ]
        self.search_combobox["values"] = suggestions     

class ConfigManager:
    def __init__(self, app):
        self.app = app
        self.configs = {}
        self.current_config_name = None
        self.configs_file = os.path.join(app.appdata_dir, "configs.json")
        if os.path.exists(self.configs_file):  # Check if the config file exists
            self.load_configs()
        else:
            self.initialize_default_config()

    def initialize_default_config(self):
        """Khởi tạo cấu hình mặc định nếu không có file configs."""
        self.configs = {
            "Mặc định": {
                "field_groups": {
                    "Thông tin công ty": self.app.default_fields[0:11],
                    "Thông tin ĐDPL": self.app.default_fields[11:22],
                    "Thông tin thành viên": [],
                    "Thông tin uỷ quyền": self.app.default_fields[22:],
                    "Ngành nghề kinh doanh": []
                    
                },
                "templates": {},
                "entries": []
            }
        }
        self.save_configs()
        self.current_config_name = "Mặc định"
        self.app.field_groups = self.configs["Mặc định"]["field_groups"]
        self.app.saved_entries = self.configs["Mặc định"]["entries"]
        self.app.fields = self.app.default_fields.copy()
        
    def load_configs(self):
        """Tải cấu hình từ file JSON trong AppData."""
        try:
            if os.path.exists(self.configs_file):
                with open(self.configs_file, 'r', encoding='utf-8') as f:
                    self.configs = json.load(f)
            else:
                self.initialize_default_config()
        except Exception as e:
            logging.error(f"Lỗi khi tải configs: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể tải cấu hình: {str(e)}")

    def load_selected_config(self, event):
        """Tải cấu hình được chọn từ dropdown."""
        self.current_config_name = self.app.config_var.get()
        self.app.field_groups = self.configs.get(self.current_config_name, {}).get("field_groups", {})
        self.app.saved_entries = self.configs.get(self.current_config_name, {}).get("entries", [])
        
        self.app.fields = [field for fields in self.app.field_groups.values() for field in fields if fields]
        self.app.clear_tabs()
        self.app.tab_manager.create_tabs()
        self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
        self.app.tab_dropdown.set(list(self.app.field_groups.keys())[0] if self.app.field_groups else "")
        
        # Xóa giá trị trong dropdown và không tự động chọn dữ liệu đầu tiên
        self.app.load_data_dropdown.set("")
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        self.app.update_template_tree()
        
        # Không tự động tải dữ liệu đầu tiên
        self.app.update_field_dropdown()  # Cập nhật danh sách trường khi tải cấu hình

    def save_configs(self):
        """Lưu cấu hình (bao gồm field_groups, templates, và entries) vào file JSON trong AppData."""
        try:
            with open(self.configs_file, 'w', encoding='utf-8') as f:
                json.dump(self.configs, f, ensure_ascii=False, indent=4)
            logging.info("Đã lưu cấu hình")
        except Exception as e:
            logging.error(f"Lỗi khi lưu configs: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể lưu cấu hình: {str(e)}")

    def add_new_config(self):
        config_name = simpledialog.askstring("Thêm cấu hình", "Nhập tên cấu hình mới:")
        if config_name and config_name not in self.configs:
            self.configs[config_name] = {
                "field_groups": {
                    "Thông tin công ty": self.app.default_fields[0:11],
                    "Thông tin ĐDPL": self.app.default_fields[11:22],
                    "Thông tin thành viên": [],
                    "Thông tin uỷ quyền": self.app.default_fields[22:],
                    "Ngành nghề kinh doanh": []  
                },
                "templates": {},
                "entries": []
            }
            self.save_configs()
            self.app.config_dropdown["values"] = list(self.configs.keys())
            self.app.config_dropdown.set(config_name)
            self.load_selected_config(None)
            logging.info(f"Thêm cấu hình '{config_name}'")

    def delete_config(self):
        if not self.current_config_name:
            return
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa cấu hình '{self.current_config_name}' không?"):
            del self.configs[self.current_config_name]
            self.save_configs()
            self.app.config_dropdown["values"] = list(self.configs.keys())
            if self.configs:
                self.app.config_dropdown.set(list(self.configs.keys())[0])
                self.load_selected_config(None)
            else:
                self.initialize_default_config()
                self.app.config_dropdown.set("Mặc định")
                self.load_selected_config(None)
            logging.info(f"Xóa cấu hình '{self.current_config_name}'")

    def rename_config(self):
        old_name = self.current_config_name
        if not old_name:
            return
        new_name = simpledialog.askstring("Sửa tên cấu hình", "Nhập tên mới:", initialvalue=old_name)
        if new_name and new_name != old_name and new_name not in self.configs:
            self.configs[new_name] = self.configs.pop(old_name)
            self.current_config_name = new_name
            self.save_configs()
            self.app.config_dropdown["values"] = list(self.configs.keys())
            self.app.config_dropdown.set(new_name)
            messagebox.showinfo("Thành công", f"Đã đổi tên thành '{new_name}'!")
            logging.info(f"Đổi tên cấu hình từ '{old_name}' thành '{new_name}'")

class FieldManager:
    def __init__(self, app):
        self.app = app

    def add_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        new_field = simpledialog.askstring("Thêm trường", "Nhập tên trường mới:")
        if new_field:
            new_field = new_field.strip()
            if not new_field:
                messagebox.showwarning("Cảnh báo", "Tên trường không được để trống!")
                return
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")  # Lấy tab hiện tại từ notebook
            if not current_tab:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab trước khi thêm trường!")
                return
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Trường này đã tồn tại trong tab hiện tại!")
                return
            self.app.field_groups[current_tab].append(new_field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Thêm trường '{new_field}' vào tab '{current_tab}'")
            
    def delete_selected_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        field = self.app.field_var.get()
        if not field:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn trường để xóa!")
            return
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        if field in self.app.field_groups[current_tab]:
            self.app.field_groups[current_tab].remove(field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Xóa trường '{field}' khỏi tab '{current_tab}'")

    def rename_selected_field(self):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        old_field = self.app.field_var.get()
        if not old_field:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn trường để sửa tên!")
            return
        new_field = simpledialog.askstring("Sửa tên trường", "Nhập tên mới cho trường:", initialvalue=old_field)
        if new_field and new_field != old_field:
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Tên trường đã tồn tại trong tab này!")
                return
            idx = self.app.field_groups[current_tab].index(old_field)
            self.app.field_groups[current_tab][idx] = new_field
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Sửa tên trường từ '{old_field}' thành '{new_field}' trong tab '{current_tab}'")

    def rename_field(self, field):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        old_field = field
        new_field = simpledialog.askstring("Sửa tên trường", "Nhập tên mới cho trường:", initialvalue=old_field)
        if new_field and new_field != old_field:
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
            if new_field in self.app.field_groups[current_tab]:
                messagebox.showwarning("Cảnh báo", "Tên trường đã tồn tại trong tab này!")
                return
            idx = self.app.field_groups[current_tab].index(old_field)
            self.app.field_groups[current_tab][idx] = new_field
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Sửa tên trường từ '{old_field}' thành '{new_field}' trong tab '{current_tab}'")

    def delete_field(self, field):
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        if field in self.app.field_groups[current_tab]:
            self.app.field_groups[current_tab].remove(field)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"][current_tab] = self.app.field_groups[current_tab]
            self.app.config_manager.save_configs()
            self.app.config_manager.load_selected_config(None)
            logging.info(f"Xóa trường '{field}' khỏi tab '{current_tab}'")

class TabManager:
    def __init__(self, app):
        self.app = app

    def create_tabs(self):
        """Tạo các tab dựa trên field_groups, bao gồm tab Thông tin thành viên và Ngành nghề kinh doanh."""
        tab_names = list(self.app.field_groups.keys())
        for i, tab_name in enumerate(tab_names):
            if tab_name == "Ngành nghề kinh doanh":
                self.app.industry_manager.create_industry_tab()
            elif tab_name == "Ngành bổ sung":
                self.app.industry_manager.create_additional_industry_tab()
            elif tab_name == "Ngành giảm":
                self.app.industry_manager.create_removed_industry_tab()
            elif tab_name == "Ngành điều chỉnh":
                self.app.industry_manager.create_adjusted_industry_tab()
            elif tab_name == "Thông tin thành viên":
                self.app.member_manager.create_member_tab()
            else:
                tab = ttk.Frame(self.app.notebook)
                self.app.notebook.add(tab, text=tab_name)
                canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
                scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
                scrollable_frame = ttk.Frame(canvas, padding=5)
                scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
                canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                scrollable_frame.grid_columnconfigure(0, weight=1)
                scrollable_frame.grid_columnconfigure(1, weight=2)
                
                # Đảm bảo mỗi tab có đủ trường để kiểm tra cuộn (tạm thời thêm nhiều trường giả)
                fields = self.app.field_groups[tab_name]
                for j, field in enumerate(fields):
                    display_name = field.replace('_', ' ').title()
                    label_width = max(60, len(display_name) // 2)
                    label = ttk.Label(scrollable_frame, text=f"{display_name}:", width=label_width, anchor="e", wraplength=500)
                    label.grid(row=j, column=0, padx=5, pady=2, sticky="e")
                    self.app.labels[field] = label
                    entry = ttk.Entry(scrollable_frame, width=50)
                    entry.grid(row=j, column=1, padx=5, pady=2, sticky="ew")
                    self.app.entries[field] = entry

                    # Gắn sự kiện tự động lưu
                    entry.bind("<KeyRelease>", lambda event: self.app.data_manager.save_entry_data())
                    #entry.bind("<FocusOut>", lambda event: self.app.data_manager.save_entry_data())

                    # Gọi phương thức để thêm menu ngữ cảnh cho ô nhập liệu
                    self.app.add_entry_context_menu(entry)

                    # Lấy menu ngữ cảnh hiện có của entry
                    context_menu = entry.context_menu if hasattr(entry, 'context_menu') else tk.Menu(self.app.root, tearoff=0)
                    entry.context_menu = context_menu

                    # Thêm các tùy chọn quản lý trường vào menu ngữ cảnh
                    context_menu.add_separator()
                    context_menu.add_command(label="Thêm trường", image=self.app.add_icon_img, compound="left", command=self.app.field_manager.add_field)
                    context_menu.add_command(label="Xóa trường", image=self.app.delete_icon_img, compound="left", command=lambda f=field: self.app.field_manager.delete_field(f))
                    context_menu.add_command(label="Sửa tên trường", image=self.app.edit_icon_img, compound="left", command=lambda f=field: self.app.field_manager.rename_field(f))
                    
                    # Hàm hiển thị menu ngữ cảnh
                    def show_context_menu(event, menu=context_menu):
                        menu.post(event.x_root, event.y_root)

                    # Gán menu ngữ cảnh cho cả label và entry
                    label.bind("<Button-3>", show_context_menu)
                    entry.bind("<Button-3>", show_context_menu)

                    if field == "von_đieu_le":
                        def update_von_dieu_le_bang_chu(event):
                            von_dieu_le_value = self.app.entries["von_đieu_le"].get()
                            von_dieu_le_bang_chu = number_to_words(von_dieu_le_value)
                            if "von_đieu_le_bang_chu" in self.app.entries:
                                self.app.entries["von_đieu_le_bang_chu"].delete(0, tk.END)
                                self.app.entries["von_đieu_le_bang_chu"].insert(0, von_dieu_le_bang_chu)
                        entry.bind("<KeyRelease>", update_von_dieu_le_bang_chu)

                    if field == "so_tien":
                        def update_so_tien_bang_chu(event):
                            so_tien_value = self.app.entries["so_tien"].get()
                            so_tien_bang_chu = number_to_words(so_tien_value)
                            if "so_tien_bang_chu" in self.app.entries:
                                self.app.entries["so_tien_bang_chu"].delete(0, tk.END)
                                self.app.entries["so_tien_bang_chu"].insert(0, so_tien_bang_chu)
                        entry.bind("<KeyRelease>", update_so_tien_bang_chu)

                # Thêm sự kiện cuộn chuột cho từng canvas riêng biệt
                def on_mousewheel(event, c=canvas):  # Truyền canvas cụ thể vào hàm
                    c.yview_scroll(int(-1 * (event.delta / 120)), "units")
                canvas.bind("<Enter>", lambda e, c=canvas: c.bind_all("<MouseWheel>", lambda evt: on_mousewheel(evt, c)))
                canvas.bind("<Leave>", lambda e, c=canvas: c.unbind_all("<MouseWheel>"))

        if self.app.current_tab_index < len(tab_names):
            self.app.notebook.select(self.app.current_tab_index)
            
        self.app.update_field_dropdown()  # Cập nhật danh sách trường khi tạo tab

    def add_tab(self):
        tab_name = simpledialog.askstring("Thêm tab", "Nhập tên tab mới:")
        if tab_name and tab_name not in self.app.field_groups:
            self.app.field_groups[tab_name] = []
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            self.app.clear_tabs()
            self.create_tabs()
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(tab_name)
            messagebox.showinfo("Thành công", f"Đã thêm tab '{tab_name}'!")
            logging.info(f"Thêm tab '{tab_name}'")

    def delete_tab(self):
        selected_tab = self.app.tab_var.get()
        if not selected_tab:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab để xóa!")
            return
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa tab '{selected_tab}' không?"):
            for field in self.app.field_groups[selected_tab]:
                self.app.fields.remove(field)
            del self.app.field_groups[selected_tab]
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            self.clear_tabs()
            self.create_tabs()
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(list(self.app.field_groups.keys())[0] if self.app.field_groups else "")
            messagebox.showinfo("Thành công", f"Đã xóa tab '{selected_tab}'!")
            logging.info(f"Xóa tab '{selected_tab}'")

    def rename_tab(self):
        old_tab = self.app.tab_var.get()
        if not old_tab:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tab để sửa tên!")
            return
        new_tab = simpledialog.askstring("Sửa tên tab", "Nhập tên mới:", initialvalue=old_tab)
        if new_tab and new_tab != old_tab and new_tab not in self.app.field_groups:
            self.app.field_groups[new_tab] = self.app.field_groups.pop(old_tab)
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["field_groups"] = self.app.field_groups
            self.app.config_manager.save_configs()
            self.clear_tabs()
            self.create_tabs()
            self.app.tab_dropdown["values"] = list(self.app.field_groups.keys())
            self.app.tab_dropdown.set(new_tab)
            messagebox.showinfo("Thành công", f"Đã đổi tên tab thành '{new_tab}'!")
            logging.info(f"Đổi tên tab từ '{old_tab}' thành '{new_tab}'")

    def clear_tabs(self): 
        """Xóa tất cả các tab trong notebook."""
        current_tab_index = self.app.notebook.index("current") if self.app.notebook.tabs() else 0
        for tab in self.app.notebook.tabs():
            self.app.notebook.forget(tab)
        self.app.notebook.update_idletasks()
        self.app.entries.clear()
        self.app.labels.clear()
        self.app.current_tab_index = current_tab_index

class DataManager:
    def __init__(self, app):
        self.app = app

    def add_entry_data(self):
        if not self.app.config_manager.current_config_name:
            return
        entry_name = simpledialog.askstring("Thêm dữ liệu", "Nhập tên cho bộ dữ liệu mới:")
        if entry_name:
            data = {field: "" for field in self.app.fields}
            data["nganh_nghe"] = [] 
            self.app.saved_entries.append({"name": entry_name, "data": data})
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
            self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
            self.app.config_manager.save_configs()
            self.app.load_data_dropdown.set(entry_name)
            self.app.load_selected_entry(None)
            messagebox.showinfo("Thành công", f"Đã thêm dữ liệu '{entry_name}'!")
            logging.info(f"Thêm dữ liệu '{entry_name}'")

    def save_entry_data(self):
        """Lưu dữ liệu từ các trường nhập liệu."""
        selected_name = self.app.load_data_var.get()
        if not selected_name:
            # Kiểm tra nếu người dùng thực sự nhập liệu
            current_data = {field: self.app.entries[field].get() for field in self.app.entries}
            if any(value.strip() for value in current_data.values()):  # Nếu có dữ liệu được nhập
                messagebox.showwarning("Cảnh báo", "Vui lòng khởi tạo dữ liệu trước!")
                self.app.add_data_button.invoke()  # Gọi nút "Thêm dữ liệu"
            return

        if not self.app.config_manager.current_config_name:
            return

        data = {field: self.app.entries[field].get() for field in self.app.entries}

        # Tự động thêm von_đieu_le_bang_chu vào dữ liệu
        if "von_đieu_le" in data:
            data["von_đieu_le_bang_chu"] = number_to_words(data["von_đieu_le"])

        # Tự động thêm so_tien_bang_chu vào dữ liệu
        if "so_tien" in data:
            data["so_tien_bang_chu"] = number_to_words(data["so_tien"])

        # Lấy danh sách ngành nghề từ industry_tree
        industries = []
        if hasattr(self.app, 'industry_tree'):
            for item in self.app.industry_tree.get_children():
                values = self.app.industry_tree.item(item)['values']
                industry = {
                    "ten_nganh": values[0],
                    "ma_nganh": values[1],
                    "la_nganh_chinh": values[2] == "X"
                }
                industries.append(industry)

        for i, entry in enumerate(self.app.saved_entries):
            if entry["name"] == selected_name:
                entry["data"].update(data)
                entry["data"]["nganh_nghe"] = industries  # Cập nhật danh sách ngành nghề
                self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
                self.app.config_manager.save_configs()
                logging.info(f"Lưu dữ liệu '{selected_name}'")
                break

    def delete_entry_data(self):
        if not self.app.config_manager.current_config_name:
            return
        selected_name = self.app.load_data_var.get()
        if not selected_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn dữ liệu để xóa!")
            return
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa dữ liệu '{selected_name}' không?"):
            self.app.saved_entries = [entry for entry in self.app.saved_entries if entry["name"] != selected_name]
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
            self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
            self.app.config_manager.save_configs()
            messagebox.showinfo("Thành công", f"Dữ liệu '{selected_name}' đã được xóa!")
            logging.info(f"Xóa dữ liệu '{selected_name}'")
            for entry in self.app.entries.values():
                entry.delete(0, tk.END)

    def rename_entry_data(self):
        if not self.app.config_manager.current_config_name:
            return
        selected_name = self.app.load_data_var.get()
        if not selected_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn dữ liệu để sửa tên!")
            return
        new_name = simpledialog.askstring("Sửa tên dữ liệu", "Nhập tên mới:", initialvalue=selected_name)
        if new_name and new_name != selected_name:
            if new_name in [entry["name"] for entry in self.app.saved_entries]:
                messagebox.showwarning("Cảnh báo", "Tên này đã tồn tại, vui lòng chọn tên khác!")
                return
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    entry["name"] = new_name
                    break
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
            self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
            self.app.load_data_dropdown.set(new_name)
            self.app.config_manager.save_configs()
            messagebox.showinfo("Thành công", f"Đã đổi tên dữ liệu thành '{new_name}'!")
            logging.info(f"Đổi tên dữ liệu từ '{selected_name}' thành '{new_name}'")

    def clear_entries(self):
        """Xóa thông tin trong tab hiện tại."""
        # Lấy tab hiện tại
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        
        # Lấy danh sách các trường trong tab hiện tại
        fields_in_current_tab = self.app.field_groups.get(current_tab, [])
        
        # Xóa dữ liệu trong các trường thuộc tab hiện tại
        for field in fields_in_current_tab:
            if field in self.app.entries:
                self.app.entries[field].delete(0, tk.END)
        
        logging.info(f"Đã xóa thông tin trong tab '{current_tab}'")

    def add_entry_data_from_import(self, entry_data):
        """Thêm dữ liệu từ file Excel vào danh sách."""
        if not self.app.config_manager.current_config_name:
            return

        entry_name = entry_data.get("ten_doanh_nghiep", f"Dữ liệu {len(self.app.saved_entries) + 1}")
        if any(entry["name"] == entry_name for entry in self.app.saved_entries):
            entry_name = f"{entry_name}_{len(self.app.saved_entries) + 1}"

        self.app.saved_entries.append({"name": entry_name, "data": entry_data})
        self.app.config_manager.configs[self.app.config_manager.current_config_name]["entries"] = self.app.saved_entries
        self.app.load_data_dropdown["values"] = [entry["name"] for entry in self.app.saved_entries]
        self.app.config_manager.save_configs()
        logging.info(f"Thêm dữ liệu từ file Excel: {entry_name}")

class MemberManager:
    def __init__(self, app):
        self.app = app

    def create_member_tab(self):
        """Tạo tab Thông tin thành viên với Treeview và scrollbar dọc."""
        tab_name = "Thông tin thành viên"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)
        
        # Tạo Canvas và Scrollbar để hỗ trợ cuộn dọc
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar_y = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set)
        
        # Đặt vị trí Canvas và Scrollbar dọc
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar_y.pack(side="right", fill="y")
        
        # Tạo frame chứa Treeview
        tree_frame = ttk.Frame(scrollable_frame)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Chỉ hiển thị các cột quan trọng
        self.app.displayed_columns = ["ho_ten", "so_cccd", "dia_chi_thuong_tru", "von_gop", "ty_le_gop", "la_chu_tich"]
        self.app.member_tree = ttk.Treeview(tree_frame, columns=self.app.displayed_columns, show="headings", height=15)  # Tăng height để hiển thị nhiều hàng hơn
        
        # Định nghĩa tiêu đề và chiều rộng cột
        column_widths = {
            "ho_ten": 200,      # Họ tên rộng hơn để dễ đọc
            "so_cccd": 100,
            "dia_chi_thuong_tru": 400,
            "von_gop": 120,
            "ty_le_gop": 100,
            "la_chu_tich": 80,  # Cột Chủ tịch nhỏ gọn
        }
        for col in self.app.displayed_columns:
            self.app.member_tree.heading(col, text=col.replace('_', ' ').title())
            width = column_widths.get(col, 100)
            self.app.member_tree.column(col, width=width, anchor="center")
        
        # Đặt Treeview vào frame
        self.app.member_tree.pack(side="top", fill="both", expand=True)
        
        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm thành viên", command=self.add_member)
        context_menu.add_command(label="Sửa thành viên", command=self.edit_member)
        context_menu.add_command(label="Xóa thành viên", command=self.delete_member)
        context_menu.add_command(label="Xem chi tiết", command=self.view_member_details)

        def show_context_menu(event):
            selected = self.app.member_tree.identify_row(event.y)
            if selected:
                self.app.member_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.member_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.member_tree.bind("<Double-1>", self.view_member_details)

        # Gắn các sự kiện kéo thả
        self.app.member_tree.bind("<Button-1>", self.start_drag_member)
        self.app.member_tree.bind("<B1-Motion>", self.drag_member)
        self.app.member_tree.bind("<ButtonRelease-1>", self.drop_member)
        
        # Tạo frame cho các nút
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Thêm thành viên", command=self.add_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa thành viên", command=self.edit_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa thành viên", command=self.delete_member).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_member_details).pack(side="left", padx=10, expand=True)
        
        # Kích hoạt cuộn chuột dọc
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Gọi load_member_data để hiển thị dữ liệu ngay khi tab được tạo
        self.load_member_data()

    def load_member_data(self):
        """Tải dữ liệu thành viên vào Treeview."""
        self.app.member_tree.delete(*self.app.member_tree.get_children())
        selected_name = self.app.load_data_var.get()
        members = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                members = entry["data"].get("thanh_vien", [])
                break
        for member in members:
            values = [member.get(col, "") for col in self.app.displayed_columns]
            values[self.app.displayed_columns.index("la_chu_tich")] = "X" if member.get("la_chu_tich", False) else ""
            self.app.member_tree.insert("", "end", values=values)

    def add_member(self):
        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Thêm thành viên", 450, 550, tabs)

        # Thêm checkbox "Chức danh chủ tịch" vào dưới cùng của popup
        chairman_var = tk.BooleanVar(value=False)
        chairman_check = ttk.Checkbutton(popup, text="Chức danh chủ tịch", variable=chairman_var)
        chairman_check.pack(pady=10)

        # Nút xác nhận
        def confirm_add():
            member = {col: entries[col].get() for col in self.app.member_columns}
            member["la_chu_tich"] = chairman_var.get()
            selected_name = self.app.load_data_var.get()
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    members = entry["data"].setdefault("thanh_vien", [])
                    # Nếu thành viên mới là chủ tịch, bỏ chọn chủ tịch cũ
                    if member["la_chu_tich"]:
                        for existing_member in members:
                            if existing_member.get("la_chu_tich", False):
                                existing_member["la_chu_tich"] = False
                        members.insert(0, member)  # Đặt chủ tịch lên đầu
                    else:
                        members.append(member)
                    self.app.config_manager.save_configs()
                    self.load_member_data()
                    popup.destroy()
                    break

        ttk.Button(popup, text="Thêm", command=confirm_add).pack(pady=10)

    def delete_member(self):
        """Xóa nhiều thành viên đã chọn."""
        selected_items = self.app.member_tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một thành viên để xóa!")
            return

        if messagebox.askyesno("Xác nhận", f"Bạn có muốn xóa {len(selected_items)} thành viên đã chọn không?"):
            selected_name = self.app.load_data_var.get()
            if not selected_name:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một mục dữ liệu trước!")
                return

            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    if "thanh_vien" not in entry["data"] or not entry["data"]["thanh_vien"]:
                        messagebox.showwarning("Cảnh báo", "Không có thành viên nào để xóa!")
                        return

                    indices = [self.app.member_tree.index(item) for item in selected_items]
                    indices.sort(reverse=True)

                    for idx in indices:
                        try:
                            entry["data"]["thanh_vien"].pop(idx)
                        except IndexError:
                            continue

                    self.app.config_manager.save_configs()
                    self.load_member_data()
                    break
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy mục dữ liệu tương ứng!")

    def edit_member(self):
        """Sửa thông tin thành viên đã chọn."""
        selected_item = self.app.member_tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thành viên để sửa!")
            return
        idx = self.app.member_tree.index(selected_item)
        selected_name = self.app.load_data_var.get()
        member = None
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                member = entry["data"]["thanh_vien"][idx]
                break

        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Sửa thành viên", 450, 550, tabs)

        # Điền dữ liệu hiện tại vào các trường
        for field, entry in entries.items():
            entry.insert(0, member.get(field, ""))

        # Thêm checkbox "Chức danh chủ tịch"
        chairman_var = tk.BooleanVar(value=member.get("la_chu_tich", False))
        chairman_check = ttk.Checkbutton(popup, text="Chức danh chủ tịch", variable=chairman_var)
        chairman_check.pack(pady=5)

        # Nút xác nhận
        def confirm_edit():
            updated_member = {col: entries[col].get() for col in self.app.member_columns}
            updated_member["la_chu_tich"] = chairman_var.get()
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    members = entry["data"]["thanh_vien"]
                    # Nếu thành viên được sửa thành chủ tịch, bỏ chọn chủ tịch cũ
                    if updated_member["la_chu_tich"] and not member.get("la_chu_tich", False):
                        for i, existing_member in enumerate(members):
                            if i != idx and existing_member.get("la_chu_tich", False):
                                existing_member["la_chu_tich"] = False
                        # Đưa thành viên lên đầu danh sách nếu là chủ tịch
                        members.pop(idx)
                        members.insert(0, updated_member)
                    else:
                        members[idx] = updated_member
                    self.app.config_manager.save_configs()
                    self.load_member_data()
                    popup.destroy()
                    break

        ttk.Button(popup, text="Lưu", command=confirm_edit).pack(pady=10)

    def view_member_details(self, event=None):
        """Hiển thị chi tiết thông tin thành viên."""
        selected_item = self.app.member_tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thành viên để xem chi tiết!")
            return
        idx = self.app.member_tree.index(selected_item)
        selected_name = self.app.load_data_var.get()
        member = None
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                member = entry["data"]["thanh_vien"][idx]
                break

        tabs = {
            "Thông tin cá nhân": ["ho_ten", "gioi_tinh", "ngay_sinh", "dan_toc", "quoc_tich"],
            "Thông tin giấy tờ": ["loai_giay_to", "so_cccd", "ngay_cap", "noi_cap", "ngay_het_han"],
            "Thông tin địa chỉ": ["dia_chi_thuong_tru", "dia_chi_lien_lac"],
            "Thông tin vốn góp": ["von_gop", "ty_le_gop", "ngay_gop_von"]
        }
        popup, entries = create_popup_with_notebook(self.app.root, "Chi tiết thành viên", 450, 550, tabs)

        # Điền dữ liệu hiện tại vào các trường và đặt chúng ở trạng thái chỉ đọc
        for field, entry in entries.items():
            entry.insert(0, member.get(field, ""))
            entry.config(state="readonly")

        # Hiển thị trạng thái chủ tịch
        ttk.Label(popup, text="Chức danh: " + ("Chủ tịch" if member.get("la_chu_tich", False) else "Thành viên")).pack(pady=5)

        # Nút đóng
        ttk.Button(popup, text="Chỉnh sửa", command=lambda: [popup.destroy(), self.edit_member()]).pack(side="left", padx=10, expand=True)
        ttk.Button(popup, text="Đóng", command=popup.destroy).pack(side="right", pady=10, expand=True)

    def start_drag_member(self, event):
        """Bắt đầu kéo một thành viên."""
        item = self.app.member_tree.identify_row(event.y)
        if item:
            self.app.drag_item = item

    def drag_member(self, event):
        """Di chuyển thành viên khi kéo."""
        if self.app.drag_item:
            self.app.member_tree.selection_set(self.app.drag_item)

    def drop_member(self, event):
        """Thả thành viên vào vị trí mới và cập nhật danh sách."""
        if self.app.drag_item:
            target = self.app.member_tree.identify_row(event.y)
            if target and target != self.app.drag_item:
                selected_name = self.app.load_data_var.get()
                for entry in self.app.saved_entries:
                    if entry["name"] == selected_name:
                        members = entry["data"].get("thanh_vien", [])
                        dragged_idx = self.app.member_tree.index(self.app.drag_item)
                        target_idx = self.app.member_tree.index(target)
                        dragged_member = members[dragged_idx]
                        # Nếu thành viên được kéo là chủ tịch, giữ nó ở đầu
                        if dragged_member.get("la_chu_tich", False):
                            messagebox.showwarning("Cảnh báo", "Chủ tịch phải ở vị trí đầu tiên!")
                            break
                        # Di chuyển thành viên trong danh sách
                        members.insert(target_idx, members.pop(dragged_idx))
                        self.app.config_manager.save_configs()
                        self.load_member_data()
                        break
            self.app.drag_item = None

class IndustryManager:
    def __init__(self, app):
        self.app = app

    def create_industry_tab(self):
        """Tạo tab Ngành nghề kinh doanh với Treeview và các nút quản lý."""
        tab_name = "Ngành nghề kinh doanh"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)
        
        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="center")  # Căn giữa scrollable_frame
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Căn giữa canvas và scrollbar trong tab
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)  # Thêm padding để tạo khoảng cách
        scrollbar.pack(side="right", fill="y")
        
        # Tăng kích thước Treeview và căn giữa
        self.app.industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        
        # Tăng chiều rộng cột để hiển thị nội dung rõ ràng hơn
        self.app.industry_tree.column("ten_nganh", width=600)  # Tăng từ 400 lên 600
        self.app.industry_tree.column("ma_nganh", width=150)   # Tăng từ 100 lên 150
        self.app.industry_tree.column("la_nganh_chinh", width=150)  # Tăng từ 100 lên 150
        
        # Căn giữa Treeview trong scrollable_frame
        self.app.industry_tree.pack(fill="x", expand=True, padx=10, pady=10)
        
        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.industry_tree.identify_row(event.y)
            if selected:
                self.app.industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.industry_tree.bind("<Double-1>", self.view_industry_details)
        
        # Tạo frame cho các nút và căn giữa
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)
        
        # Tạo các nút và đặt chúng trong button_frame (xóa nút "Đặt ngành chính")
        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details).pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Load dữ liệu ngành nghề
        self.load_industry_data()

    def load_industry_data(self):
        """Tải dữ liệu ngành nghề kinh doanh vào Treeview."""
        self.app.industry_tree.delete(*self.app.industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        main_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                main_industries = entry["data"].get("nganh_nghe", [])
                break

        for industry in main_industries:
            self.app.industry_tree.insert(
                "", "end",
                values=(
                    industry["ten_nganh"],
                    industry["ma_nganh"],
                    "Có" if industry.get("la_nganh_chinh", False) else "Không"
                )
            )
            
    def get_current_tab_tree_and_data(self):
        """Xác định Treeview và danh sách ngành dựa trên tab hiện tại."""
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        selected_name = self.app.load_data_var.get()

        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                if current_tab == "Ngành bổ sung":
                    return self.app.additional_industry_tree, entry["data"].get("nganh_bo_sung", [])
                elif current_tab == "Ngành giảm":
                    return self.app.removed_industry_tree, entry["data"].get("nganh_giam", [])
                elif current_tab == "Ngành điều chỉnh":
                    return self.app.adjusted_industry_tree, entry["data"].get("nganh_dieu_chinh", [])
                else:  # Mặc định là "Ngành nghề kinh doanh"
                    return self.app.industry_tree, entry["data"].get("nganh_nghe", [])
        return None, None

    def load_industry_data_for_current_tab(self, tree, industries):
        """Tải lại dữ liệu cho Treeview và danh sách ngành của tab hiện tại."""
        tree.delete(*tree.get_children())
        for industry in industries:
            tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry.get("la_nganh_chinh", False) else ""))
            self.app.config_manager.save_configs()
            
    def add_industry(self):
        """Thêm một ngành nghề mới với logic tùy chỉnh cho từng tab."""
        current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
        selected_name = self.app.load_data_var.get()

        if current_tab in ["Ngành giảm", "Ngành điều chỉnh"]:
            # Lấy danh sách ngành từ tab "Ngành nghề kinh doanh"
            main_industries = []
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    main_industries = entry["data"].get("nganh_nghe", [])
                    break

            if not main_industries:
                messagebox.showwarning("Cảnh báo", "Không có ngành nghề nào trong 'Ngành nghề kinh doanh' để thêm!")
                return

            # Tạo popup
            popup = create_popup(self.app.root, f"Thêm {current_tab.lower()}", 600, 250)
            ttk.Label(popup, text="Chọn ngành từ danh sách ngành nghề kinh doanh:").pack(pady=5)
            industry_var = tk.StringVar()
            industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90, state="readonly")
            industry_combo.pack(pady=5)

            # Tạo danh sách các ngành để hiển thị trong combobox
            combo_values = []
            for industry in main_industries:
                display_text = f"{industry['ma_nganh']} - {industry['ten_nganh']}"
                combo_values.append(display_text)
            industry_combo['values'] = combo_values

            def confirm_add():
                selected_industry = industry_var.get()
                if not selected_industry:
                    messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành để thêm!")
                    return

                ma_nganh, ten_nganh = selected_industry.split(" - ", 1)

                # Tìm ngành được chọn trong danh sách ngành chính
                selected_industry_data = None
                for industry in main_industries:
                    if industry["ma_nganh"] == ma_nganh:
                        selected_industry_data = industry.copy()
                        break

                if not selected_industry_data:
                    messagebox.showwarning("Cảnh báo", "Không tìm thấy thông tin ngành!")
                    return

                # Kiểm tra xem ngành đã tồn tại trong danh sách đích chưa
                for entry in self.app.saved_entries:
                    if entry["name"] == selected_name:
                        target_industries = entry["data"].setdefault(
                            "nganh_giam" if current_tab == "Ngành giảm" else "nganh_dieu_chinh", 
                            []
                        )
                        
                        # Kiểm tra trùng lặp
                        if any(industry["ma_nganh"] == ma_nganh for industry in target_industries):
                            messagebox.showwarning("Cảnh báo", f"Ngành này đã tồn tại trong {current_tab}!")
                            return

                        # Thêm ngành vào danh sách đích
                        target_industries.append(selected_industry_data)
                        self.app.config_manager.save_configs()

                        # Tải lại dữ liệu và cập nhật giao diện
                        if current_tab == "Ngành giảm":
                            self.load_removed_industry_data()
                            self.sync_main_industry_tab(
                                action="add",
                                current_tab=current_tab,
                                updated_industry=selected_industry_data
                            )
                        else:  # Ngành điều chỉnh
                            self.load_adjusted_industry_data()
                            self.sync_main_industry_tab(
                                action="add",
                                current_tab=current_tab,
                                updated_industry=selected_industry_data
                            )
                        popup.destroy()
                        break

            ttk.Button(popup, text="Thêm", command=confirm_add).pack(side="bottom", padx=10, pady=10)


        elif current_tab == "Ngành bổ sung":
            # Logic thêm ngành cho tab "Ngành bổ sung"
            self.add_industry_default_logic(current_tab)

            # Đồng bộ ngành bổ sung với "Ngành nghề kinh doanh"
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    industries = entry["data"].get("nganh_bo_sung", [])
                    if industries:
                        updated_industry = industries[-1]  # Lấy ngành vừa thêm
                        self.sync_main_industry_tab(action="add", current_tab=current_tab, updated_industry=updated_industry)
                        self.app.industry_manager.load_industry_data()  # Tải lại dữ liệu cho tab "Ngành nghề kinh doanh"
                    break

        else:
            # Logic mặc định cho các tab khác (ví dụ: "Ngành nghề kinh doanh")
            self.add_industry_default_logic(current_tab)

    def add_industry_default_logic(self, current_tab):
        """Logic mặc định để thêm ngành cho các tab khác."""
        # Đường dẫn đến file industry_codes.json trong thư mục AppData
        industry_codes_path = os.path.join(self.app.appdata_dir, "industry_codes.json")
        
        try:
            with open(industry_codes_path, "r", encoding="utf-8") as f:
                industry_codes = json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Lỗi", f"Không tìm thấy file {industry_codes_path}!")
            return

        # Tạo popup
        popup = create_popup(self.app.root, "Thêm ngành", 600, 250)
        popup.title("Thêm ngành")

        ttk.Label(popup, text="Chọn mã ngành và tên ngành:").pack(pady=5)
        industry_var = tk.StringVar()
        industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90)
        industry_combo.pack(pady=5)

        full_list = [f"{code['ma_nganh']} - {code['ten_nganh']}" for code in industry_codes]

        def update_list(*args):
            search_term = industry_var.get().lower()
            filtered_list = [item for item in full_list if search_term in item.lower()] if search_term else full_list
            industry_combo['values'] = filtered_list

        industry_combo.bind('<KeyRelease>', update_list)
        industry_combo['values'] = full_list

        ttk.Label(popup, text="Chi tiết ngành (nếu có):").pack(pady=5)
        chi_tiet_var = tk.StringVar()
        chi_tiet_entry = ttk.Entry(popup, textvariable=chi_tiet_var, width=90)
        chi_tiet_entry.pack(pady=5)

        main_industry_var = tk.BooleanVar(value=False)
        main_industry_check = ttk.Checkbutton(popup, text="Ngành chính", variable=main_industry_var)
        main_industry_check.pack(pady=5)

        def confirm_add():
            selected_industry = industry_var.get()
            chi_tiet = chi_tiet_var.get().strip()
            if not selected_industry:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn mã ngành và tên ngành!")
                return
            try:
                ma_nganh, ten_nganh = selected_industry.split(" - ", 1)
            except ValueError:
                messagebox.showerror("Lỗi", "Định dạng mã ngành không hợp lệ!")
                return

            if chi_tiet:
                ten_nganh = f"{ten_nganh} - {chi_tiet}"

            selected_name = self.app.load_data_var.get()
            for entry in self.app.saved_entries:
                if entry["name"] == selected_name:
                    # Xác định danh sách ngành theo tab hiện tại
                    if current_tab == "Ngành bổ sung":
                        industries = entry["data"].setdefault("nganh_bo_sung", [])
                    else:  # Mặc định là "Ngành nghề kinh doanh"
                        industries = entry["data"].setdefault("nganh_nghe", [])

                    # Kiểm tra trùng lặp mã ngành
                    if any(industry["ma_nganh"] == ma_nganh for industry in industries):
                        messagebox.showwarning("Cảnh báo", "Mã ngành này đã tồn tại!")
                        return

                    # Nếu ngành mới là ngành chính, bỏ chọn ngành chính cũ
                    if main_industry_var.get():
                        for existing_industry in industries:
                            if existing_industry.get("la_nganh_chinh", False):
                                existing_industry["la_nganh_chinh"] = False

                    # Tạo ngành mới
                    new_industry = {
                        "ma_nganh": ma_nganh,
                        "ten_nganh": ten_nganh,
                        "la_nganh_chinh": main_industry_var.get()
                    }

                    # Thêm ngành mới vào danh sách
                    industries.append(new_industry)

                    # Lưu cấu hình
                    self.app.config_manager.save_configs()

                    # Cập nhật hiển thị
                    if current_tab == "Ngành bổ sung":
                        self.load_additional_industry_data()
                        # Đồng bộ với tab "Ngành nghề kinh doanh"
                        self.sync_main_industry_tab(
                            action="add",
                            current_tab=current_tab,
                            updated_industry=new_industry
                        )
                    else:
                        self.load_industry_data()

                    popup.destroy()
                    break

        ttk.Button(popup, text="Thêm", command=confirm_add).pack(side="bottom", padx=10, pady=10)


    def delete_industry(self):
        """Xóa nhiều ngành nghề đã chọn."""
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một ngành nghề để xóa!")
            return

        if messagebox.askyesno("Xác nhận", f"Bạn có muốn xóa {len(selected_items)} ngành nghề đã chọn không?"):
            indices = [tree.index(item) for item in selected_items]
            indices.sort(reverse=True)  # Xóa từ cuối danh sách để tránh lỗi chỉ số
            removed_industries = []
            for idx in indices:
                try:
                    removed_industries.append(industries.pop(idx))
                except IndexError:
                    continue

            self.app.config_manager.save_configs()
            self.load_industry_data_for_current_tab(tree, industries)

            # Đồng bộ tab "Ngành nghề kinh doanh" nếu tab hiện tại là "Ngành bổ sung"
            current_tab = self.app.notebook.tab(self.app.notebook.select(), "text")
            if current_tab == "Ngành bổ sung":
                for removed_industry in removed_industries:
                    self.sync_main_industry_tab(
                        action="delete",
                        current_tab=current_tab,
                        updated_industry=removed_industry
                    )

    def edit_industry(self, event=None):
        """Sửa chi tiết ngành nghề với bố cục tương tự Thêm ngành và checkbox Ngành chính."""
        # Đường dẫn đến file industry_codes.json trong thư mục AppData
        industry_codes_path = os.path.join(self.app.appdata_dir, "industry_codes.json")
        
        # Tải danh sách mã ngành từ file
        try:
            with open(industry_codes_path, "r", encoding="utf-8") as f:
                industry_codes = json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Lỗi", f"Không tìm thấy file {industry_codes_path}! Vui lòng tạo file này chứa danh sách mã ngành trong thư mục AppData.")
            return

        # Xác định Treeview và danh sách ngành dựa trên tab hiện tại
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        # Kiểm tra xem ngành đã được chọn chưa
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để sửa!")
            return

        # Lấy chỉ số ngành được chọn
        idx = tree.index(selected_item)
        try:
            industry = industries[idx]
        except IndexError:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy ngành nghề để sửa!")
            return

        # Tạo popup sửa ngành
        popup = create_popup(self.app.root, "Sửa ngành nghề", 600, 250)
        ttk.Label(popup, text="Chọn mã ngành và tên ngành:").pack(pady=10)
        industry_var = tk.StringVar()

        # Tách chi tiết (nếu có) khỏi tên ngành
        ten_nganh = industry["ten_nganh"]
        chi_tiet = ""
        if " - " in ten_nganh:
            base_ten_nganh, chi_tiet = ten_nganh.rsplit(" - ", 1)
            for code in industry_codes:
                if code["ten_nganh"] == base_ten_nganh:
                    industry_var.set(f"{industry['ma_nganh']} - {base_ten_nganh}")
                    break
        else:
            industry_var.set(f"{industry['ma_nganh']} - {ten_nganh}")
        
        industry_combo = ttk.Combobox(popup, textvariable=industry_var, width=90)
        industry_combo.pack(pady=10)
        
        # Danh sách đầy đủ các mã ngành và tên ngành
        full_list = [f"{code['ma_nganh']} - {code['ten_nganh']}" for code in industry_codes]
        
        # Hàm cập nhật danh sách lọc
        def update_list(*args):
            search_term = industry_var.get().lower()  # Lấy chuỗi tìm kiếm và chuyển thành chữ thường
            if search_term:
                # Lọc các mục chứa chuỗi tìm kiếm (trong mã ngành hoặc tên ngành)
                filtered_list = [item for item in full_list if search_term in item.lower()]
            else:
                # Nếu không có chuỗi tìm kiếm, hiển thị toàn bộ danh sách
                filtered_list = full_list
            industry_combo['values'] = filtered_list  # Cập nhật danh sách trong combobox
        
        # Gắn sự kiện <KeyRelease> để gọi hàm update_list khi người dùng gõ
        industry_combo.bind('<KeyRelease>', update_list)
        
        # Khởi tạo combobox với danh sách đầy đủ ban đầu
        industry_combo['values'] = full_list
        
        ttk.Label(popup, text="Chi tiết ngành (nếu có):").pack(pady=5)
        chi_tiet_var = tk.StringVar(value=chi_tiet)
        chi_tiet_entry = ttk.Entry(popup, textvariable=chi_tiet_var, width=90)
        chi_tiet_entry.pack(pady=10)

        # Thêm checkbox "Ngành chính"
        main_industry_var = tk.BooleanVar(value=industry.get("la_nganh_chinh", False))
        main_industry_check = ttk.Checkbutton(popup, text="Ngành chính", variable=main_industry_var)
        main_industry_check.pack(pady=5)
        
        def confirm_edit():
            selected_industry = industry_var.get()
            chi_tiet = chi_tiet_var.get().strip()
            if not selected_industry:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn mã ngành và tên ngành!")
                return
            ma_nganh, ten_nganh = selected_industry.split(" - ", 1)
            if chi_tiet:
                ten_nganh = f"{ten_nganh} - {chi_tiet}"
            for entry in self.app.saved_entries:
                if entry["name"] == self.app.load_data_var.get():
                    # Nếu ngành được sửa thành ngành chính, bỏ chọn ngành chính cũ
                    if main_industry_var.get() and not industry.get("la_nganh_chinh", False):
                        for i, existing_industry in enumerate(industries):
                            if i != idx and existing_industry.get("la_nganh_chinh", False):
                                existing_industry["la_nganh_chinh"] = False
                    industries[idx] = {"ma_nganh": ma_nganh, "ten_nganh": ten_nganh, "la_nganh_chinh": main_industry_var.get()}
                    self.app.config_manager.save_configs()

                    # Tải lại dữ liệu cho tab hiện tại
                    self.load_industry_data_for_current_tab(tree, industries)
                    self.sync_main_industry_tab(
                        action="edit",
                        current_tab=self.app.notebook.tab(self.app.notebook.select(), "text"),
                        updated_industry={"ma_nganh": ma_nganh, "ten_nganh": ten_nganh, "la_nganh_chinh": main_industry_var.get()},
                        original_industry=industry
                    )
                    popup.destroy()
                    break
        
        ttk.Button(popup, text="Lưu", command=confirm_edit).pack(side="bottom", padx=10, pady=10, expand=True)

    def create_additional_industry_tab(self):
        """Tạo tab Ngành bổ sung với Treeview và các nút quản lý."""
        tab_name = "Ngành bổ sung"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.additional_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.additional_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.additional_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.additional_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.additional_industry_tree.column("ten_nganh", width=600)
        self.app.additional_industry_tree.column("ma_nganh", width=150)
        self.app.additional_industry_tree.column("la_nganh_chinh", width=150)
        self.app.additional_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)
        
        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.additional_industry_tree.identify_row(event.y)
            if selected:
                self.app.additional_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.additional_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.additional_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details).pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành bổ sung
        self.load_additional_industry_data()

    def create_removed_industry_tab(self):
        """Tạo tab Ngành giảm với Treeview và các nút quản lý."""
        tab_name = "Ngành giảm"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.removed_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.removed_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.removed_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.removed_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.removed_industry_tree.column("ten_nganh", width=600)
        self.app.removed_industry_tree.column("ma_nganh", width=150)
        self.app.removed_industry_tree.column("la_nganh_chinh", width=150)
        self.app.removed_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)

        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.removed_industry_tree.identify_row(event.y)
            if selected:
                self.app.removed_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.removed_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.removed_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details).pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành giảm
        self.load_removed_industry_data()

    def create_adjusted_industry_tab(self):
        """Tạo tab Ngành điều chỉnh với Treeview và các nút quản lý."""
        tab_name = "Ngành điều chỉnh"
        tab = ttk.Frame(self.app.notebook)
        self.app.notebook.add(tab, text=tab_name)

        # Tạo Canvas và Scrollbar để hỗ trợ cuộn
        canvas = tk.Canvas(tab, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding=5)
        scrollable_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Tạo Treeview
        self.app.adjusted_industry_tree = ttk.Treeview(scrollable_frame, columns=("ten_nganh", "ma_nganh", "la_nganh_chinh"), show="headings", height=15)
        self.app.adjusted_industry_tree.heading("ten_nganh", text="Tên ngành")
        self.app.adjusted_industry_tree.heading("ma_nganh", text="Mã ngành")
        self.app.adjusted_industry_tree.heading("la_nganh_chinh", text="Ngành chính")
        self.app.adjusted_industry_tree.column("ten_nganh", width=600)
        self.app.adjusted_industry_tree.column("ma_nganh", width=150)
        self.app.adjusted_industry_tree.column("la_nganh_chinh", width=150)
        self.app.adjusted_industry_tree.pack(fill="x", expand=True, padx=10, pady=10)

        # Tạo các nút quản lý
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10)

        # Tạo menu ngữ cảnh (context menu)
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm ngành", command=self.add_industry)
        context_menu.add_command(label="Sửa ngành", command=self.edit_industry)
        context_menu.add_command(label="Xóa ngành", command=self.delete_industry)
        context_menu.add_command(label="Ngành chính", command=self.set_main_industry)  # Thêm tùy chọn "Ngành chính"
        context_menu.add_command(label="Xem chi tiết", command=self.view_industry_details)

        def show_context_menu(event):
            selected = self.app.adjusted_industry_tree.identify_row(event.y)
            if selected:
                self.app.adjusted_industry_tree.selection_set(selected)
            context_menu.post(event.x_root, event.y_root)

        # Gắn sự kiện nhấp chuột phải để hiển thị menu ngữ cảnh
        self.app.adjusted_industry_tree.bind("<Button-3>", show_context_menu)
        
        # Gắn sự kiện nhấp đúp để xem chi tiết
        self.app.adjusted_industry_tree.bind("<Double-1>", self.view_industry_details)

        ttk.Button(button_frame, text="Thêm ngành", command=self.add_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xóa ngành", command=self.delete_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Sửa ngành", command=self.edit_industry).pack(side="left", padx=10, expand=True)
        ttk.Button(button_frame, text="Xem chi tiết", command=self.view_industry_details).pack(side="left", padx=10, expand=True)

        # Kích hoạt cuộn chuột
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Load dữ liệu ngành điều chỉnh
        self.load_adjusted_industry_data()
    
            
    def load_additional_industry_data(self):
        """Tải dữ liệu ngành bổ sung vào Treeview."""
        self.app.additional_industry_tree.delete(*self.app.additional_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        additional_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                additional_industries = entry["data"].get("nganh_bo_sung", [])
                break
        for industry in additional_industries:
            self.app.additional_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))

    def load_removed_industry_data(self):
        """Tải dữ liệu ngành giảm vào Treeview."""
        self.app.removed_industry_tree.delete(*self.app.removed_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        removed_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                removed_industries = entry["data"].get("nganh_giam", [])
                break
        for industry in removed_industries:
            self.app.removed_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))

    def load_adjusted_industry_data(self):
        """Tải dữ liệu ngành điều chỉnh vào Treeview."""
        self.app.adjusted_industry_tree.delete(*self.app.adjusted_industry_tree.get_children())
        selected_name = self.app.load_data_var.get()
        adjusted_industries = []
        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                adjusted_industries = entry["data"].get("nganh_dieu_chinh", [])
                break
        for industry in adjusted_industries:
            self.app.adjusted_industry_tree.insert("", "end", values=(industry["ten_nganh"], industry["ma_nganh"], "X" if industry["la_nganh_chinh"] else ""))
            
    def view_industry_details(self, event=None):
        """Hiển thị chi tiết thông tin ngành nghề."""
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để xem chi tiết!")
            return

        idx = tree.index(selected_item)
        try:
            industry = industries[idx]
        except IndexError:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy ngành nghề để xem chi tiết!")
            return

        # Tạo popup chi tiết ngành nghề
        popup = create_popup(self.app.root, "Chi tiết ngành nghề", 600, 250)
        ttk.Label(popup, text="Tên ngành:").pack(pady=5)
        ttk.Label(popup, text=industry.get("ten_nganh", ""), wraplength=450).pack(pady=5)

        ttk.Label(popup, text="Mã ngành:").pack(pady=5)
        ttk.Label(popup, text=industry.get("ma_nganh", "")).pack(pady=5)

        ttk.Label(popup, text="Ngành chính:").pack(pady=5)
        ttk.Label(popup, text="Có" if industry.get("la_nganh_chinh", False) else "Không").pack(pady=5)

        ttk.Button(popup, text="Đóng", command=popup.destroy).pack(side="right", padx=5, expand=True)
        ttk.Button(popup, text="Chỉnh sửa", command=lambda: [popup.destroy(), self.edit_industry()]).pack(side="left", padx=5, expand=True)
    def set_main_industry(self):
        """Đặt ngành nghề được chọn làm ngành chính."""
        # Xác định Treeview và danh sách ngành dựa trên tab hiện tại
        tree, industries = self.get_current_tab_tree_and_data()
        if not tree or not industries:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu ngành nghề!")
            return

        # Kiểm tra xem ngành đã được chọn chưa
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ngành nghề để đặt làm ngành chính!")
            return

        # Lấy chỉ số ngành được chọn
        idx = tree.index(selected_item)
        try:
            selected_industry = industries[idx]
        except IndexError:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy ngành nghề để đặt làm ngành chính!")
            return

        # Đặt ngành được chọn làm ngành chính và bỏ chọn ngành chính cũ
        for i, industry in enumerate(industries):
            industry["la_nganh_chinh"] = (i == idx)

        # Lưu cấu hình và tải lại dữ liệu
        self.app.config_manager.save_configs()
        self.load_industry_data_for_current_tab(tree, industries)

        messagebox.showinfo("Thành công", f"Đã đặt ngành '{selected_industry['ten_nganh']}' làm ngành chính!")

    def sync_main_industry_tab(self, action, current_tab, updated_industry=None, original_industry=None):
        """
        Đồng bộ dữ liệu của tab Ngành nghề kinh doanh dựa trên các thay đổi từ các tab khác.
        
        :param action: Loại hành động ("add", "edit", "delete").
        :param current_tab: Tab hiện tại ("Ngành bổ sung", "Ngành giảm", "Ngành điều chỉnh").
        :param updated_industry: Ngành mới được thêm hoặc sửa (nếu có).
        :param original_industry: Ngành cũ trước khi sửa (chỉ áp dụng cho "edit").
        """
        selected_name = self.app.load_data_var.get()

        for entry in self.app.saved_entries:
            if entry["name"] == selected_name:
                main_industries = entry["data"].setdefault("nganh_nghe", [])

                if current_tab == "Ngành bổ sung":
                    if action == "add" and updated_industry:
                        # Kiểm tra nếu ngành đã tồn tại trong "Ngành nghề kinh doanh"
                        if not any(industry["ma_nganh"] == updated_industry["ma_nganh"] for industry in main_industries):
                            main_industries.append(updated_industry)
                    
                    elif action == "edit" and updated_industry and original_industry:
                        # Tìm và cập nhật ngành tương ứng trong "Ngành nghề kinh doanh"
                        for industry in main_industries:
                            if industry["ma_nganh"] == original_industry["ma_nganh"]:
                                industry.update({
                                    "ma_nganh": updated_industry["ma_nganh"],
                                    "ten_nganh": updated_industry["ten_nganh"],
                                    "la_nganh_chinh": updated_industry["la_nganh_chinh"]
                                })
                                break

                    elif action == "delete" and updated_industry:
                        # Xóa ngành tương ứng khỏi "Ngành nghề kinh doanh"
                        main_industries = [
                            industry for industry in main_industries
                            if industry["ma_nganh"] != updated_industry["ma_nganh"]
                        ]

                elif current_tab == "Ngành giảm":
                    # Khi thêm ngành vào "Ngành giảm", xóa ngành tương ứng khỏi "Ngành nghề kinh doanh"
                    if action == "add" and updated_industry:
                        main_industries = [
                            industry for industry in main_industries
                            if industry["ma_nganh"] != updated_industry["ma_nganh"]
                        ]

                elif current_tab == "Ngành điều chỉnh":
                    # Khi sửa ngành trong "Ngành điều chỉnh", cập nhật thông tin ngành trong "Ngành nghề kinh doanh"
                    if action == "edit" and updated_industry and original_industry:
                        for industry in main_industries:
                            if industry["ma_nganh"] == original_industry["ma_nganh"]:
                                industry.update(updated_industry)
                                break

                # Cập nhật lại dữ liệu trong "Ngành nghề kinh doanh"
                entry["data"]["nganh_nghe"] = main_industries
                break

        # Lưu cấu hình và tải lại dữ liệu cho tab "Ngành nghề kinh doanh"
        self.app.config_manager.save_configs()
        self.load_industry_data()

class TemplateManager:
    def __init__(self, app):
        self.app = app

    def update_template_tree(self):
        """Cập nhật Treeview với templates của cấu hình hiện tại."""
        # Xóa tất cả các mục hiện tại trong Treeview
        for item in self.app.template_tree.get_children():
            self.app.template_tree.delete(item)

        # Kiểm tra cấu hình hiện tại và lấy danh sách templates
        if self.app.config_manager.current_config_name and self.app.config_manager.current_config_name in self.app.config_manager.configs:
            templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
            for template in templates.keys():
                # Loại bỏ đuôi .docx khỏi tên hiển thị nhưng lưu tên đầy đủ trong values
                display_name = os.path.splitext(template)[0]
                self.app.template_tree.insert("", "end", text=display_name, values=(template,))

    def drop_template_files(self, event):
        """Thêm template bằng kéo thả, sao chép file vào thư mục templates và tránh ghi đè."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        files = self.app.root.tk.splitlist(event.data)
        added_count = 0
        templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
        for file_path in files:
            if file_path.endswith('.docx') and os.path.exists(file_path):
                template_name = os.path.basename(file_path)
                base_name = os.path.splitext(template_name)[0]
                extension = os.path.splitext(template_name)[1]
                new_name = base_name + extension
                counter = 1
                target_path = os.path.join(self.app.templates_dir, new_name)
                while os.path.exists(target_path):
                    new_name = f"{base_name}_{counter}{extension}"
                    target_path = os.path.join(self.app.templates_dir, new_name)
                    counter += 1
                while new_name in templates:
                    new_name = f"{base_name}_{counter}{extension}"
                    target_path = os.path.join(self.app.templates_dir, new_name)
                    counter += 1
                try:
                    import shutil
                    shutil.copy2(file_path, target_path)
                    templates[new_name] = new_name
                    added_count += 1
                except Exception as e:
                    logging.error(f"Lỗi khi sao chép template {file_path}: {str(e)}")
                    messagebox.showerror("Lỗi", f"Không thể sao chép template: {str(e)}")
        if added_count > 0:
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = templates
            self.app.config_manager.save_configs()
            self.update_template_tree()
            messagebox.showinfo("Thành công", f"Đã thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'!")
            logging.info(f"Thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'")

    def add_multiple_templates(self):
        """Thêm nhiều template, sao chép file vào thư mục templates và tránh ghi đè."""
        if not self.app.config_manager.current_config_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cấu hình trước!")
            return
        template_paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")], title="Chọn nhiều template Word")
        if template_paths:
            added_count = 0
            templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
            for template_path in template_paths:
                if os.path.exists(template_path):
                    template_name = os.path.basename(template_path)
                    base_name = os.path.splitext(template_name)[0]
                    extension = os.path.splitext(template_name)[1]
                    new_name = base_name + extension
                    counter = 1
                    target_path = os.path.join(self.app.templates_dir, new_name)
                    while os.path.exists(target_path):
                        new_name = f"{base_name}_{counter}{extension}"
                        target_path = os.path.join(self.app.templates_dir, new_name)
                        counter += 1
                    while new_name in templates:
                        new_name = f"{base_name}_{counter}{extension}"
                        target_path = os.path.join(self.app.templates_dir, new_name)
                        counter += 1
                    try:
                        import shutil
                        shutil.copy2(template_path, target_path)
                        templates[new_name] = new_name
                        added_count += 1
                    except Exception as e:
                        logging.error(f"Lỗi khi sao chép template {template_path}: {str(e)}")
                        messagebox.showerror("Lỗi", f"Không thể sao chép template: {str(e)}")
            if added_count > 0:
                self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = templates
                self.app.config_manager.save_configs()
                self.update_template_tree()
                messagebox.showinfo("Thành công", f"Đã thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'!")
                logging.info(f"Thêm {added_count} template vào cấu hình '{self.app.config_manager.current_config_name}'")
   
    def delete_template(self):
        """Xóa template khỏi cấu hình hiện tại."""
        if not self.app.config_manager.current_config_name:
            return
        selected_items = self.app.template_tree.selection()
        if not selected_items:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn template để xóa!")
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa các template đã chọn không?"):
            templates = self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {})
            for item in selected_items:
                # Lấy tên đầy đủ từ values thay vì text
                full_template_name = self.app.template_tree.item(item)["values"][0]
                if full_template_name in templates:
                    del templates[full_template_name]
            self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = templates
            self.app.config_manager.save_configs()
            self.update_template_tree()
            messagebox.showinfo("Thành công", "Đã xóa các template đã chọn!")
            logging.info(f"Xóa template đã chọn khỏi cấu hình '{self.app.config_manager.current_config_name}'")

    def show_template_context_menu(self, event):
        context_menu = tk.Menu(self.app.root, tearoff=0)
        context_menu.add_command(label="Thêm template", command=self.add_multiple_templates)
        context_menu.add_command(label="Xóa", command=self.delete_template)
        context_menu.post(event.x_root, event.y_root)

    # kéo thả templates
    def start_drag(self, event):
        item = self.app.template_tree.identify_row(event.y)
        if item:
            self.app.drag_item = item

    def drag_template(self, event):
        if self.app.drag_item:
            self.app.template_tree.selection_set(self.app.drag_item)

    def drop_template(self, event):
        if self.app.drag_item:
            target = self.app.template_tree.identify_row(event.y)
            if target and target != self.app.drag_item:
                templates = list(self.app.config_manager.configs[self.app.config_manager.current_config_name].get("templates", {}).keys())
                # Lấy tên đầy đủ từ values thay vì text
                dragged_full_name = self.app.template_tree.item(self.app.drag_item)["values"][0]
                target_full_name = self.app.template_tree.item(target)["values"][0]
                dragged_idx = templates.index(dragged_full_name)
                target_idx = templates.index(target_full_name)
                templates.insert(target_idx, templates.pop(dragged_idx))
                new_templates = {templates[i]: templates[i] for i in range(len(templates))}
                self.app.config_manager.configs[self.app.config_manager.current_config_name]["templates"] = new_templates
                self.app.config_manager.save_configs()
                self.update_template_tree()
            self.app.drag_item = None

class BackupManager:
    def __init__(self, app):
        self.app = app

    def auto_backup(self):
        """Sao lưu tự động dữ liệu mỗi 10 phút vào thư mục backup."""
        try:
            backup_file = os.path.join(self.app.backup_dir, f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            with open(backup_file, 'w', encoding='utf-8') as f:
                json.dump(self.app.config_manager.configs, f, ensure_ascii=False, indent=4)
            logging.info(f"Sao lưu tự động: {backup_file}")
            self.app.root.after(600000, self.auto_backup)
        except Exception as e:
            logging.error(f"Lỗi khi sao lưu tự động: {str(e)}")

    def restore_from_backup(self):
        """Hiển thị popup để chọn và khôi phục dữ liệu từ file sao lưu, với tùy chọn xóa file cũ."""
        if not os.path.exists(self.app.backup_dir) or not os.listdir(self.app.backup_dir):
            messagebox.showinfo("Thông báo", "Chưa có file sao lưu nào!")
            return

        popup = create_popup(self.app.root, "Khôi phục từ sao lưu", 500, 500)
        ttk.Label(popup, text="Chọn file sao lưu để khôi phục:").pack(pady=5)

        backup_frame = ttk.Frame(popup)
        backup_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        backup_tree = ttk.Treeview(backup_frame, columns=("lastmod"), show="tree headings", height=10, selectmode="browse")
        backup_tree.heading("#0", text="Tên file")
        backup_tree.heading("lastmod", text="Ngày tạo")
        backup_tree.column("#0", width=250)
        backup_tree.column("lastmod", width=150)
        backup_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(backup_frame, orient="vertical", command=backup_tree.yview)
        scrollbar.pack(side="right", fill="y")
        backup_tree.configure(yscrollcommand=scrollbar.set)

        backup_files = sorted([f for f in os.listdir(self.app.backup_dir) if f.endswith(".json")], reverse=True)
        for backup_file in backup_files:
            file_path = os.path.join(self.app.backup_dir, backup_file)
            timestamp = datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
            backup_tree.insert("", "end", text=backup_file, values=(timestamp,))

        def confirm_restore():
            selected = backup_tree.selection()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một file sao lưu!")
                return
            backup_file = backup_tree.item(selected[0])["text"]
            file_path = os.path.join(self.app.backup_dir, backup_file)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.app.config_manager.configs = json.load(f)
                self.app.config_manager.save_configs()
                self.app.config_manager.load_selected_config(None)
                popup.destroy()
                messagebox.showinfo("Thành công", f"Đã khôi phục từ file: {backup_file}")
                logging.info(f"Khôi phục từ sao lưu: {backup_file}")
            except Exception as e:
                logging.error(f"Lỗi khi khôi phục: {str(e)}")
                messagebox.showerror("Lỗi", f"Không thể khôi phục: {str(e)}")

        def delete_old_backups():
            if messagebox.askyesno("Xác nhận", "Bạn có muốn xóa các file sao lưu cũ (giữ lại 10 file mới nhất)?"):
                backup_files_sorted = sorted(
                    [f for f in os.listdir(self.app.backup_dir) if f.endswith(".json")],
                    key=lambda x: os.path.getctime(os.path.join(self.app.backup_dir, x)),
                    reverse=True
                )
                files_to_delete = backup_files_sorted[10:]
                for file in files_to_delete:
                    os.remove(os.path.join(self.app.backup_dir, file))
                messagebox.showinfo("Thành công", f"Đã xóa {len(files_to_delete)} file sao lưu cũ!")
                logging.info(f"Xóa {len(files_to_delete)} file sao lưu cũ")
                for item in backup_tree.get_children():
                    backup_tree.delete(item)
                remaining_files = sorted([f for f in os.listdir(self.app.backup_dir) if f.endswith(".json")], reverse=True)
                for backup_file in remaining_files:
                    file_path = os.path.join(self.app.backup_dir, backup_file)
                    timestamp = datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    backup_tree.insert("", "end", text=backup_file, values=(timestamp,))
        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Khôi phục", command=confirm_restore, style="primary.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="Xóa file cũ", command=delete_old_backups, style="danger.TButton").pack(side="left", padx=5)
    