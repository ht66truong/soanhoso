# Thư viện chuẩn
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import logging
from datetime import datetime
from tkinter import messagebox

# Thư viện bên thứ ba
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES

# Module nội bộ
from modules.utils import ToolTip, create_popup
from modules.export import (ExportManager)
from modules.member import MemberManager
from modules.industry import IndustryManager
from modules.config import ConfigManager, BackupManager
from modules.data import DataManager, FieldManager, TabManager, TemplateManager
from modules.employee import EmployeeManager

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
        self.root.title("Ứng dụng nhập liệu hồ sơ Kinh Doanh v6.0.6")
        window_width = 1366
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

        self.default_fields = ["tên_doanh_nghiệp", "tên_nước_ngoài", "tên_viết_tắt", "mã_số_doanh_nghiệp", "ngày_cấp_mst", 
                               "số_nhà_tên_đường", "xã_phường", "quận_huyện", "tỉnh_thành_phố", "số_điện_thoại", "vốn_điều_lệ",
                               "họ_tên", "giới_tính", "ngày_sinh", "dân_tộc", "quốc_tịch", "số_cccd", "ngày_cấp", "nơi_cấp", "ngày_hết_hạn", "địa_chỉ_thường_trú", "địa_chỉ_liên_lạc",
                               "họ_tên_uq", "giới_tính_uq", "ngày_sinh_uq", "số_cccd_uq", "ngày_cấp_uq", "nơi_cấp_uq", "địa_chỉ_liên_lạc_uq", "sdt_uq", "email_uq"]
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
        self.employee_manager = EmployeeManager(self)
        
        

        # Khởi tạo top_frame trước
        self.top_frame = ttk.Frame(self.root)
        self.top_frame.pack(side="top", fill="x", padx=5, pady=5)
        
        self.main_frame = ttk.Frame(root, padding=5)
        self.main_frame.pack(fill="both", expand=True)

        # Config Frame
        self.config_frame = ttk.LabelFrame(self.main_frame, text="Quản lý cấu hình", padding=5)
        self.config_frame.pack(fill="x", pady=(0, 5))

        self.config_var = tk.StringVar()
        self.config_dropdown = ttk.Combobox(self.config_frame, textvariable=self.config_var, values=list(self.configs.keys()), state="readonly", width=20)
        self.config_dropdown.pack(side="left", padx=5)
        ToolTip(self.config_dropdown, "* Thông báo thay đổi: \n"
                                            "- Ngành, nghề kinh doanh \n"
                                            "- Thông tin đăng ký thuế \n \n"
                                        "* Đăng ký thay đổi: \n"
                                            "- Tên doanh nghiệp \n"
                                            "- Địa chỉ trụ sở chính \n"
                                            "- Thành viên công ty TNHH \n"
                                            "- Người đại diện theo pháp luật \n"
                                            "- Chủ sở hữu công ty TNHH 1TV \n"
                                            "- Vốn điều lệ của công ty, tỷ lệ vốn góp \n"
                                            "- Người đứng đầu chi nhánh/văn phòng đại diện/địa điểm kinh doanh \n")
        self.config_dropdown.bind("<<ComboboxSelected>>", self.load_selected_config)

        def resource_path(relative_path): # Hàm hỗ trợ lấy đường dẫn tài nguyên
            """Lấy đường dẫn tuyệt đối đến tài nguyên, hoạt động cho cả mã nguồn và file thực thi PyInstaller."""
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # Load icons
        try:
            add_icon = Image.open(resource_path("icon/add_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            delete_icon = Image.open(resource_path("icon/delete_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            save_icon = Image.open(resource_path("icon/save_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            clear_icon = Image.open(resource_path("icon/clear_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            import_icon = Image.open(resource_path("icon/imxls_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            xls_icon = Image.open(resource_path("icon/xls_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            preview_icon = Image.open(resource_path("icon/preview_icon.png")).resize((50, 50), Image.Resampling.LANCZOS)
            export_icon = Image.open(resource_path("icon/export_icon.png")).resize((50, 50), Image.Resampling.LANCZOS)
            remove_icon = Image.open(resource_path("icon/remove_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            edit_icon = Image.open(resource_path("icon/edit_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            search_icon = Image.open(resource_path("icon/search_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            find_icon = Image.open(resource_path("icon/find_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
            restorebackup_icon = Image.open(resource_path("icon/restorebackup_icon.png")).resize((22, 22), Image.Resampling.LANCZOS)
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
        self.add_config_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_new_config, bootstyle="primary-outline")
        ToolTip(self.add_config_button, "Thêm cấu hình")
        self.add_config_button.pack(side="left", padx=5)
        self.delete_config_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_config, bootstyle="danger-outline")
        ToolTip(self.delete_config_button, "Xóa cấu hình")
        self.delete_config_button.pack(side="left", padx=5)
        self.rename_config_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_config, bootstyle="warning-outline")
        ToolTip(self.rename_config_button, "Sửa tên cấu hình")
        self.rename_config_button.pack(side="left", padx=5)

        #Quản lý tab
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=10, fill="y")
        #ttk.Label(self.config_frame, text="Quản lý tab:").pack(side="left", padx=5)
        self.tab_var = tk.StringVar()
        self.tab_dropdown = ttk.Combobox(self.config_frame, textvariable=self.tab_var, state="readonly", width=20)
        ToolTip(self.tab_dropdown, "Chọn tab để thêm/sửa/xóa")
        self.tab_dropdown.pack(side="left", padx=5)
        self.add_tab_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_tab, bootstyle="primary-outline")
        ToolTip(self.add_tab_button, "Thêm tab")
        self.add_tab_button.pack(side="left", padx=5)
        self.delete_tab_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_tab, bootstyle="danger-outline")
        ToolTip(self.delete_tab_button, "Xóa tab")
        self.delete_tab_button.pack(side="left", padx=5)
        self.rename_tab_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_tab, bootstyle="warning-outline")
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
        self.add_field_button = ttk.Button(self.config_frame, image=self.add_icon_img, command=self.add_field, bootstyle="primary-outline")
        ToolTip(self.add_field_button, "Thêm trường")
        self.add_field_button.pack(side="left", padx=5)
        self.delete_field_button = ttk.Button(self.config_frame, image=self.delete_icon_img, command=self.delete_selected_field, bootstyle="danger-outline")
        ToolTip(self.delete_field_button, "Xóa trường")
        self.delete_field_button.pack(side="left", padx=5)
        self.rename_field_button = ttk.Button(self.config_frame, image=self.edit_icon_img, command=self.rename_selected_field, bootstyle="warning-outline")
        ToolTip(self.rename_field_button, "Sửa tên trường")
        self.rename_field_button.pack(side="left", padx=5)

        #Quản lý template"
        ttk.Separator(self.config_frame, orient="vertical").pack(side="left", padx=10, fill="y")
        #ttk.Label(self.config_frame, text="Quản lý mẫu:").pack(side="left", padx=5) 
        self.template_frame = ttk.Frame(self.config_frame)
        self.template_frame.pack(side="left", fill="x", padx=5)
        ToolTip(self.template_frame, "Kéo thả để thêm mẫu/sắp xếp")
        self.template_tree = ttk.Treeview(self.template_frame, columns=(), show="tree", height=5, selectmode="extended")
        self.template_tree.column("#0", width=300)
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
        ttk.Button(search_frame, image=self.find_icon_img, command=self.search_data, bootstyle="success-outline").pack(side="left", padx=5)

        #ttk.Label(self.control_frame, text="Chọn công ty:").grid(row=0, column=2, padx=5, pady=5)
        self.load_data_var = tk.StringVar()
        self.load_data_dropdown = ttk.Combobox(self.control_frame, textvariable=self.load_data_var, state="readonly", width=10)
        self.load_data_dropdown.grid(row=0, column=3, padx=5, pady=5)
        ToolTip(self.load_data_dropdown, "Chọn công ty")
        self.load_data_dropdown.bind("<<ComboboxSelected>>", self.load_selected_entry)

        self.add_data_button = ttk.Button(self.control_frame, image=self.add_icon_img, command=self.add_entry_data, bootstyle="primary-outline")
        ToolTip(self.add_data_button, "Thêm dữ liệu")
        self.add_data_button.grid(row=0, column=4, padx=5, pady=5)
        self.delete_data_button = ttk.Button(self.control_frame, image=self.delete_icon_img, command=self.delete_entry_data, bootstyle="danger-outline")
        ToolTip(self.delete_data_button, "Xóa dữ liệu")
        self.delete_data_button.grid(row=0, column=5, padx=5, pady=5)
        self.rename_data_button = ttk.Button(self.control_frame, image=self.edit_icon_img, command=self.rename_entry_data, bootstyle="warning-outline")
        ToolTip(self.rename_data_button, "Sửa tên dữ liệu")
        self.rename_data_button.grid(row=0, column=6, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=7, padx=10, pady=5, sticky="ns")
        #self.edit_data_button = ttk.Button(self.control_frame, image=self.save_icon_img, command=self.save_entry_data, bootstyle="success")
        #ToolTip(self.edit_data_button, "Lưu thông tin")
        #self.edit_data_button.grid(row=0, column=6, padx=5, pady=5)
        self.clear_data_button = ttk.Button(self.control_frame, image=self.clear_icon_img, command=self.clear_entries, bootstyle="dark-outline")
        ToolTip(self.clear_data_button, "Xóa thông tin")
        self.clear_data_button.grid(row=0, column=8, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=9, padx=10, pady=5, sticky="ns")
        ttk.Label(self.control_frame, text="Xuất Nhập dữ liệu:").grid(row=0, column=10, padx=5, pady=5)
        self.export_excel_button = ttk.Button(self.control_frame, image=self.xls_icon_img, command=self.export_data, bootstyle="success-outline")
        ToolTip(self.export_excel_button, "Xuất dữ liệu")
        self.export_excel_button.grid(row=0, column=11, padx=5, pady=5)
        self.import_data_button = ttk.Button(self.control_frame, image=self.import_icon_img, command=self.import_from_file, bootstyle="secondary-outline")
        ToolTip(self.import_data_button, "Nhập dữ liệu")
        self.import_data_button.grid(row=0, column=12, padx=5, pady=5)
        
        self.restore_data_button = ttk.Button(self.control_frame, image=self.restorebackup_icon_img, command=self.restore_from_backup, bootstyle="success-outline")
        ToolTip(self.restore_data_button, "Khôi phục từ sao lưu")
        self.restore_data_button.grid(row=0, column=13, padx=5, pady=5)

        ttk.Separator(self.control_frame, orient="vertical").grid(row=0, column=15, padx=10, pady=5, sticky="ns")
        self.show_placeholder_button = ttk.Button(self.control_frame, image=self.search_icon_img, command=self.show_placeholder_popup, bootstyle="Secondary-outline")
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
        self.preview_word_button = ttk.Button(self.export_frame, image=self.preview_icon_img, command=self.preview_word, bootstyle="secondary-outline")
        ToolTip(self.preview_word_button, "Xem trước")
        self.preview_word_button.grid(row=0, column=1, padx=5)
        self.export_file_button = ttk.Button(self.export_frame, image=self.export_icon_img, command=self.export_file, bootstyle="success-outline")
        ToolTip(self.export_file_button, "Xuất file")
        self.export_file_button.grid(row=0, column=2, padx=5)

        # Styles
        style = ttk.Style()
        style.theme_use("cosmo")
        style.configure("TButton", font=("Segoe UI", 10), borderwidth=1, relief="raised", padding=5)
        

        # Initialize default config if none exists
        if not self.config_manager.configs:
            self.config_manager.initialize_default_config()
            
        # Update dropdown with configs
        self.config_dropdown["values"] = list(self.config_manager.configs.keys())
        if self.config_manager.configs:
            config_name = list(self.config_manager.configs.keys())[0]
            self.config_dropdown.set(config_name)
            self.config_manager.current_config_name = config_name
            self.config_manager.load_selected_config(None)
        
        #self.check_and_migrate_data()

        self.root.after(600000, self.backup_manager.auto_backup)

    def create_toolbar(self):
        """Tạo thanh công cụ chính với các biểu tượng lớn hơn và trực quan hơn"""
        self.toolbar_frame = ttk.Frame(self.root, bootstyle="PRIMARY")
        self.toolbar_frame.pack(side="top", fill="x", padx=0, pady=0)
        
        # Nút tạo mới với biểu tượng lớn
        self.new_button = ttk.Button(self.toolbar_frame, text=" Tạo mới", image=self.clear_icon_img, compound="left", 
                               command=self.clear_entries, bootstyle="success-outline", padding=10)
        self.new_button.pack(side="left", padx=5, pady=3)
        ToolTip(self.new_button, "Tạo hồ sơ mới")
        
        # Nút lưu với biểu tượng lớn
        self.save_button = ttk.Button(self.toolbar_frame, text=" Lưu", image=self.save_icon_img, compound="left", 
                               command=self.add_entry_data, bootstyle="primary-outline", padding=10)
        self.save_button.pack(side="left", padx=5, pady=3)
        ToolTip(self.save_button, "Lưu dữ liệu hiện tại")
        
        ttk.Separator(self.toolbar_frame, orient="vertical", bootstyle="SECONDARY").pack(side="left", padx=10, fill="y")
        
        # Nút xuất với biểu tượng lớn
        self.export_button = ttk.Button(self.toolbar_frame, text=" Xuất Word", image=self.export_icon_img, compound="left", 
                                 command=self.export_file, bootstyle="primary-outline", padding=10)
        self.export_button.pack(side="left", padx=5, pady=3)
        ToolTip(self.export_button, "Xuất dữ liệu ra Word")
        
        self.preview_button = ttk.Button(self.toolbar_frame, text=" Xem trước", image=self.preview_icon_img, compound="left", 
                                   command=self.preview_word, bootstyle="warning-outline", padding=10)
        self.preview_button.pack(side="left", padx=5, pady=3)
        ToolTip(self.preview_button, "Xem trước kết quả")
        
        # Phần tìm kiếm ở bên phải thanh công cụ
        search_container = ttk.Frame(self.toolbar_frame, bootstyle="SECONDARY")
        search_container.pack(side="right", padx=10, pady=3)
        
        self.toolbar_search_var = tk.StringVar()
        self.toolbar_search_entry = ttk.Entry(search_container, textvariable=self.toolbar_search_var, width=20, font=("Segoe UI", 10), bootstyle="SECONDARY")
        self.toolbar_search_entry.pack(side="left", padx=2)
        
        search_button = ttk.Button(search_container, image=self.search_icon_img, command=self.search_data, bootstyle="dark", padding=8)
        search_button.pack(side="left", padx=2)
        ToolTip(search_button, "Tìm kiếm hồ sơ")

    def load_selected_config(self, event):
        self.config_manager.load_selected_config(event)

    def load_selected_entry(self, event):
        """Tải dữ liệu được chọn từ dropdown hoặc tìm kiếm."""
        selected_name = self.load_data_var.get()
        
        if not selected_name:
            return
            
        # Lấy dữ liệu từ cơ sở dữ liệu - đảm bảo lấy entry mới nhất
        entries = self.config_manager.db_manager.get_entries(
            self.config_manager.current_config_name
        )
        
        # Cập nhật saved_entries để đồng bộ với database
        self.saved_entries = entries
        
        # Tìm entry tương ứng và điền dữ liệu
        found = False
        for entry in entries:
            if entry["name"] == selected_name:
                found = True
                
                # Xóa dữ liệu hiện tại
                for field in self.entries:
                    self.entries[field].delete(0, tk.END)
                    
                # Điền dữ liệu mới
                for field, value in entry["data"].items():
                    if field in self.entries:
                        self.entries[field].delete(0, tk.END)
                        self.entries[field].insert(0, value)
                
                # Tải dữ liệu thành viên
                if hasattr(self, 'member_tree'):
                    self.member_manager.load_member_data()
                    
                # Tải dữ liệu ngành nghề cho tất cả các tab ngành
                if hasattr(self, 'industry_tree'):
                    self.industry_manager.load_industry_data()
                    
                # Thêm các phương thức tải dữ liệu cho các tab khác
                if hasattr(self, 'additional_industry_tree'):
                    self.industry_manager.load_additional_industry_data()
                    
                if hasattr(self, 'removed_industry_tree'):
                    self.industry_manager.load_removed_industry_data()
                    
                if hasattr(self, 'adjusted_industry_tree'):
                    self.industry_manager.load_adjusted_industry_data()
                
                # Log và thông báo
                logging.info(f"Đã tải dữ liệu: {selected_name}")
                break
                    
        if not found:
            messagebox.showwarning("Cảnh báo", f"Không tìm thấy dữ liệu cho '{selected_name}'")

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
        self.config_manager.delete_current_config()  # Change from delete_config

    def rename_config(self):
        self.config_manager.rename_current_config()  # Change from rename_config

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

    def add_entry_context_menu(self, entry):
        """Thêm menu ngữ cảnh cho ô nhập liệu."""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Sao chép", command=lambda: entry.event_generate("<<Copy>>"))
        context_menu.add_command(label="Dán", command=lambda: entry.event_generate("<<Paste>>"))
        context_menu.add_command(label="Xóa", command=lambda: entry.delete(0, tk.END))
        # Lưu menu vào thuộc tính context_menu của entry
        entry.context_menu = context_menu

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

        ttk.Button(popup, text="Chọn", command=load_selected_data, bootstyle="primary").pack(side="left", padx=10, pady=10)
        ttk.Button(popup, text="Đóng", command=popup.destroy, bootstyle="danger").pack(side="right", padx=10, pady=10)

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
